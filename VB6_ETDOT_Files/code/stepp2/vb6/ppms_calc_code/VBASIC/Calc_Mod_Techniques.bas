Attribute VB_Name = "Calc_Mod_Techniques"
Option Explicit






Const Calc_Mod_Techniques_decl_end = True


Function sc_AddTechs( _
    out_PropOrd As PropertyOrder_Type, _
    in_PropCode As Long) _
    As Boolean
On Error GoTo err_ThisFunc
  With out_PropOrd
    .Property_Code = in_PropCode
    Call Get_Complete_List_of_TechCodes(.Property_Code, .Technique_Code)
  End With
exit_normally_ThisFunc:
  sc_AddTechs = True
  Exit Function
exit_err_ThisFunc:
  sc_AddTechs = False
  Exit Function
err_ThisFunc:
  ''''Call Show_Trapped_Error("sc_AddTechs")
  Resume exit_err_ThisFunc
End Function
Function Project_UserHierarchy_SetDefaults( _
    Prj As Project_Type) _
    As Boolean
On Error GoTo err_ThisFunc
  With Prj
    With .UserHierarchy
      '
      ' SET NAME OF EACH PROPERTY SHEET.
      '
      ReDim .PropertySheetOrder(1 To 10)
      .PropertySheetOrder(1).Name = PROPERTYSHEETNAME_BASIC_CHEMICAL_INFO
      .PropertySheetOrder(2).Name = "General 1"
      .PropertySheetOrder(3).Name = "General 2"
      .PropertySheetOrder(4).Name = "Transport"
      .PropertySheetOrder(5).Name = "Partitioning/Equilibrium"
      .PropertySheetOrder(6).Name = "Fire and Explosion"
      .PropertySheetOrder(7).Name = "Oxygen Demand"
      .PropertySheetOrder(8).Name = "Aquatic Toxicity 1"
      .PropertySheetOrder(9).Name = "Aquatic Toxicity 2"
      .PropertySheetOrder(10).Name = PROPERTYSHEETNAME_CHEMICAL_NOTE
      '
      '////////////////////////////////////////////////////////////////////////////////////////
      '////////////////////////////////////////////////////////////////////////////////////////
      '///////  PROPERTY SHEET {BASIC CHEMICAL INFO}: DEFAULT TECHNIQUES  /////////////////////
      '////////////////////////////////////////////////////////////////////////////////////////
      '////////////////////////////////////////////////////////////////////////////////////////
      '
              '
              ' THERE ARE NO TECHNIQUES FOR THIS PROPERTY SHEET.
              '
      With .PropertySheetOrder(1)
        ReDim .PropertyOrder(0 To 0)
      End With
      '
      '////////////////////////////////////////////////////////////////////////////////////////
      '////////////////////////////////////////////////////////////////////////////////////////
      '///////  PROPERTY SHEET "GENERAL 1": DEFAULT TECHNIQUES  ///////////////////////////////
      '////////////////////////////////////////////////////////////////////////////////////////
      '////////////////////////////////////////////////////////////////////////////////////////
      '
      With .PropertySheetOrder(2)
        ReDim .PropertyOrder(1 To 8)
        Call sc_AddTechs(.PropertyOrder(1), PROPCODE_000_MOLEC_WEIGHT)
        Call sc_AddTechs(.PropertyOrder(2), PROPCODE_001_LIQDENS_298K)
        Call sc_AddTechs(.PropertyOrder(3), PROPCODE_002_LIQDENS_FOFT)
        Call sc_AddTechs(.PropertyOrder(4), PROPCODE_003_MELTING_POINT)
        Call sc_AddTechs(.PropertyOrder(5), PROPCODE_004_NBP)
        Call sc_AddTechs(.PropertyOrder(6), PROPCODE_005_VP_298K)
        Call sc_AddTechs(.PropertyOrder(7), PROPCODE_006_VP_FOFT)
        Call sc_AddTechs(.PropertyOrder(8), PROPCODE_007_HEAT_FORMATION)
      End With
      '
      '////////////////////////////////////////////////////////////////////////////////////////
      '////////////////////////////////////////////////////////////////////////////////////////
      '///////  PROPERTY SHEET "GENERAL 2": DEFAULT TECHNIQUES  ///////////////////////////////
      '////////////////////////////////////////////////////////////////////////////////////////
      '////////////////////////////////////////////////////////////////////////////////////////
      '
      With .PropertySheetOrder(3)
        ReDim .PropertyOrder(1 To 8)
        Call sc_AddTechs(.PropertyOrder(1), PROPCODE_008_LIQUID_HEAT_CAPACITY_FOFT)
        Call sc_AddTechs(.PropertyOrder(2), PROPCODE_009_VAPOR_HEAT_CAPACITY_FOFT)
        Call sc_AddTechs(.PropertyOrder(3), PROPCODE_010_HEAT_OF_VAPORIZATION_298K)
        Call sc_AddTechs(.PropertyOrder(4), PROPCODE_011_HEAT_OF_VAPORIZATION_NBP)
        Call sc_AddTechs(.PropertyOrder(5), PROPCODE_012_HEAT_OF_VAPORIZATION_FOFT)
        Call sc_AddTechs(.PropertyOrder(6), PROPCODE_013_CRITICAL_T)
        Call sc_AddTechs(.PropertyOrder(7), PROPCODE_014_CRITICAL_P)
        Call sc_AddTechs(.PropertyOrder(8), PROPCODE_038_CRITICAL_V)
      End With
      '
      '////////////////////////////////////////////////////////////////////////////////////////
      '////////////////////////////////////////////////////////////////////////////////////////
      '///////  PROPERTY SHEET "TRANSPORT": DEFAULT TECHNIQUES  ///////////////////////////////
      '////////////////////////////////////////////////////////////////////////////////////////
      '////////////////////////////////////////////////////////////////////////////////////////
      '
      With .PropertySheetOrder(4)
        ReDim .PropertyOrder(1 To 8)
        Call sc_AddTechs(.PropertyOrder(1), PROPCODE_015_DIFFUSIVITY_H2O)
        Call sc_AddTechs(.PropertyOrder(2), PROPCODE_016_DIFFUSIVITY_AIR)
        Call sc_AddTechs(.PropertyOrder(3), PROPCODE_017_SURFACE_TENSION_298K)
        Call sc_AddTechs(.PropertyOrder(4), PROPCODE_018_SURFACE_TENSION_FOFT)
        Call sc_AddTechs(.PropertyOrder(5), PROPCODE_019_VAPOR_VISCOSITY_FOFT)
        Call sc_AddTechs(.PropertyOrder(6), PROPCODE_020_LIQUID_VISCOSITY_FOFT)
        Call sc_AddTechs(.PropertyOrder(7), PROPCODE_021_LIQUID_THERMAL_CONDUC_FOFT)
        Call sc_AddTechs(.PropertyOrder(8), PROPCODE_022_VAPOR_THERMAL_CONDUC_FOFT)
      End With
      '
      '////////////////////////////////////////////////////////////////////////////////////////
      '////////////////////////////////////////////////////////////////////////////////////////
      '///////  PROPERTY SHEET "PARTITIONING/EQUILIBRIUM": DEFAULT TECHNIQUES  ////////////////
      '////////////////////////////////////////////////////////////////////////////////////////
      '////////////////////////////////////////////////////////////////////////////////////////
      '
      With .PropertySheetOrder(5)
        ReDim .PropertyOrder(1 To 7)
        Call sc_AddTechs(.PropertyOrder(1), PROPCODE_034_AC_CHEM_IN_H2O)
        Call sc_AddTechs(.PropertyOrder(2), PROPCODE_032_AC_H2O_IN_CHEM)
        Call sc_AddTechs(.PropertyOrder(3), PROPCODE_033_HENRY_CONSTANT)
        Call sc_AddTechs(.PropertyOrder(4), PROPCODE_039_SOL_LIMIT_CHEM_IN_H2O)
        Call sc_AddTechs(.PropertyOrder(5), PROPCODE_035_LOG_KOW)
        Call sc_AddTechs(.PropertyOrder(6), PROPCODE_036_LOG_KOC)
        Call sc_AddTechs(.PropertyOrder(7), PROPCODE_037_BIOCONC_FACTOR)
      End With
      '
      '////////////////////////////////////////////////////////////////////////////////////////
      '////////////////////////////////////////////////////////////////////////////////////////
      '///////  PROPERTY SHEET "FIRE AND EXPLOSION": DEFAULT TECHNIQUES  //////////////////////
      '////////////////////////////////////////////////////////////////////////////////////////
      '////////////////////////////////////////////////////////////////////////////////////////
      '
      With .PropertySheetOrder(6)
        ReDim .PropertyOrder(1 To 5)
        Call sc_AddTechs(.PropertyOrder(1), PROPCODE_023_UF_LIMIT)
        Call sc_AddTechs(.PropertyOrder(2), PROPCODE_024_LF_LIMIT)
        Call sc_AddTechs(.PropertyOrder(3), PROPCODE_025_FLASH_POINT)
        Call sc_AddTechs(.PropertyOrder(4), PROPCODE_026_AUTOIGNITION_T)
        Call sc_AddTechs(.PropertyOrder(5), PROPCODE_027_COMBUSTION_HEAT)
      End With
      '
      '////////////////////////////////////////////////////////////////////////////////////////
      '////////////////////////////////////////////////////////////////////////////////////////
      '///////  PROPERTY SHEET "OXYGEN DEMAND": DEFAULT TECHNIQUES  ///////////////////////////
      '////////////////////////////////////////////////////////////////////////////////////////
      '////////////////////////////////////////////////////////////////////////////////////////
      '
      With .PropertySheetOrder(7)
        ReDim .PropertyOrder(1 To 4)
        Call sc_AddTechs(.PropertyOrder(1), PROPCODE_028_CARBON_THOD)
        Call sc_AddTechs(.PropertyOrder(2), PROPCODE_029_COMBINED_THOD)
        Call sc_AddTechs(.PropertyOrder(3), PROPCODE_030_COD)
        Call sc_AddTechs(.PropertyOrder(4), PROPCODE_031_BCOD)
      End With
      '
      '////////////////////////////////////////////////////////////////////////////////////////
      '////////////////////////////////////////////////////////////////////////////////////////
      '///////  PROPERTY SHEET "AQUATIC TOXICITY 1": DEFAULT TECHNIQUES  //////////////////////
      '////////////////////////////////////////////////////////////////////////////////////////
      '////////////////////////////////////////////////////////////////////////////////////////
      '
      With .PropertySheetOrder(8)
        ReDim .PropertyOrder(1 To 8)
        Call sc_AddTechs(.PropertyOrder(1), PROPCODE_041_FMINNOW_48H_EC50)
        Call sc_AddTechs(.PropertyOrder(2), PROPCODE_042_FMINNOW_96H_EC50)
        Call sc_AddTechs(.PropertyOrder(3), PROPCODE_043_FMINNOW_24H_LC50)
        Call sc_AddTechs(.PropertyOrder(4), PROPCODE_044_FMINNOW_48H_LC50)
        Call sc_AddTechs(.PropertyOrder(5), PROPCODE_045_FMINNOW_96H_LC50)
        Call sc_AddTechs(.PropertyOrder(6), PROPCODE_046_SALMONIDAE_24H_LC50)
        Call sc_AddTechs(.PropertyOrder(7), PROPCODE_047_SALMONIDAE_48H_LC50)
        Call sc_AddTechs(.PropertyOrder(8), PROPCODE_048_SALMONIDAE_96H_LC50)
      End With
      '
      '////////////////////////////////////////////////////////////////////////////////////////
      '////////////////////////////////////////////////////////////////////////////////////////
      '///////  PROPERTY SHEET "AQUATIC TOXICITY 2": DEFAULT TECHNIQUES  //////////////////////
      '////////////////////////////////////////////////////////////////////////////////////////
      '////////////////////////////////////////////////////////////////////////////////////////
      '
      With .PropertySheetOrder(9)
        ReDim .PropertyOrder(1 To 6)
        Call sc_AddTechs(.PropertyOrder(1), PROPCODE_049_DMAGNA_24H_EC50)
        Call sc_AddTechs(.PropertyOrder(2), PROPCODE_050_DMAGNA_48H_EC50)
        Call sc_AddTechs(.PropertyOrder(3), PROPCODE_051_DMAGNA_24H_LC50)
        Call sc_AddTechs(.PropertyOrder(4), PROPCODE_052_DMAGNA_48H_LC50)
        Call sc_AddTechs(.PropertyOrder(5), PROPCODE_053_MYSID_96H_LC50)
        Call sc_AddTechs(.PropertyOrder(6), PROPCODE_054_ALTERNATE_SPECIES)
      End With
      '
      '////////////////////////////////////////////////////////////////////////////////////////
      '////////////////////////////////////////////////////////////////////////////////////////
      '///////  PROPERTY SHEET {CHEMICAL NOTE}: DEFAULT TECHNIQUES  ///////////////////////////
      '////////////////////////////////////////////////////////////////////////////////////////
      '////////////////////////////////////////////////////////////////////////////////////////
      '
              '
              ' THERE ARE NO TECHNIQUES FOR THIS PROPERTY SHEET.
              '
      With .PropertySheetOrder(10)
        ReDim .PropertyOrder(0 To 0)
      End With
    End With
  End With
  '
  ' EXIT OUT OF HERE.
  '
exit_normally_ThisFunc:
  Project_UserHierarchy_SetDefaults = True
  Exit Function
exit_err_ThisFunc:
  Project_UserHierarchy_SetDefaults = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("Project_UserHierarchy_SetDefaults")
  Resume exit_err_ThisFunc
End Function


Function Given_PropCode_Get_Is_FofT( _
    in_PropCode As Long, _
    out_Is_FofT As Boolean) _
    As Boolean
On Error GoTo err_ThisFunc
  out_Is_FofT = False
  Select Case in_PropCode
    '
    '////////////////////////////////////////////////////////////////////////////////////////
    '///////  PROPERTY SHEET "GENERAL 1": DEFAULT TECHNIQUES  ///////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '
    Case PROPCODE_000_MOLEC_WEIGHT: out_Is_FofT = False
    Case PROPCODE_001_LIQDENS_298K: out_Is_FofT = False
    Case PROPCODE_002_LIQDENS_FOFT: out_Is_FofT = True
    Case PROPCODE_003_MELTING_POINT: out_Is_FofT = False
    Case PROPCODE_004_NBP: out_Is_FofT = False
    Case PROPCODE_005_VP_298K: out_Is_FofT = False
    Case PROPCODE_006_VP_FOFT: out_Is_FofT = True
    Case PROPCODE_007_HEAT_FORMATION: out_Is_FofT = False
    '
    '////////////////////////////////////////////////////////////////////////////////////////
    '///////  PROPERTY SHEET "GENERAL 2": DEFAULT TECHNIQUES  ///////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '
    Case PROPCODE_008_LIQUID_HEAT_CAPACITY_FOFT: out_Is_FofT = True
    Case PROPCODE_009_VAPOR_HEAT_CAPACITY_FOFT: out_Is_FofT = True
    Case PROPCODE_010_HEAT_OF_VAPORIZATION_298K: out_Is_FofT = False
    Case PROPCODE_011_HEAT_OF_VAPORIZATION_NBP: out_Is_FofT = False
    Case PROPCODE_012_HEAT_OF_VAPORIZATION_FOFT: out_Is_FofT = True
    Case PROPCODE_013_CRITICAL_T: out_Is_FofT = False
    Case PROPCODE_014_CRITICAL_P: out_Is_FofT = False
    Case PROPCODE_038_CRITICAL_V: out_Is_FofT = False
    '
    '////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '///////  PROPERTY SHEET "TRANSPORT": DEFAULT TECHNIQUES  ///////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '
    Case PROPCODE_015_DIFFUSIVITY_H2O: out_Is_FofT = False
    Case PROPCODE_016_DIFFUSIVITY_AIR: out_Is_FofT = False
    Case PROPCODE_017_SURFACE_TENSION_298K: out_Is_FofT = False
    Case PROPCODE_018_SURFACE_TENSION_FOFT: out_Is_FofT = True
    Case PROPCODE_019_VAPOR_VISCOSITY_FOFT: out_Is_FofT = True
    Case PROPCODE_020_LIQUID_VISCOSITY_FOFT: out_Is_FofT = True
    Case PROPCODE_021_LIQUID_THERMAL_CONDUC_FOFT: out_Is_FofT = True
    Case PROPCODE_022_VAPOR_THERMAL_CONDUC_FOFT: out_Is_FofT = True
    '
    '////////////////////////////////////////////////////////////////////////////////////////
    '///////  PROPERTY SHEET "PARTITIONING/EQUILIBRIUM": DEFAULT TECHNIQUES  ////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '
    Case PROPCODE_034_AC_CHEM_IN_H2O: out_Is_FofT = False
    Case PROPCODE_032_AC_H2O_IN_CHEM: out_Is_FofT = False
    Case PROPCODE_033_HENRY_CONSTANT: out_Is_FofT = False
    Case PROPCODE_039_SOL_LIMIT_CHEM_IN_H2O: out_Is_FofT = False
    Case PROPCODE_035_LOG_KOW: out_Is_FofT = False
    Case PROPCODE_036_LOG_KOC: out_Is_FofT = False
    Case PROPCODE_037_BIOCONC_FACTOR: out_Is_FofT = False
    '
    '////////////////////////////////////////////////////////////////////////////////////////
    '///////  PROPERTY SHEET "FIRE AND EXPLOSION": DEFAULT TECHNIQUES  //////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '
    Case PROPCODE_023_UF_LIMIT: out_Is_FofT = False
    Case PROPCODE_024_LF_LIMIT: out_Is_FofT = False
    Case PROPCODE_025_FLASH_POINT: out_Is_FofT = False
    Case PROPCODE_026_AUTOIGNITION_T: out_Is_FofT = False
    Case PROPCODE_027_COMBUSTION_HEAT: out_Is_FofT = False
    '
    '////////////////////////////////////////////////////////////////////////////////////////
    '///////  PROPERTY SHEET "OXYGEN DEMAND": DEFAULT TECHNIQUES  ///////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '
    Case PROPCODE_028_CARBON_THOD: out_Is_FofT = False
    Case PROPCODE_029_COMBINED_THOD: out_Is_FofT = False
    Case PROPCODE_030_COD: out_Is_FofT = False
    Case PROPCODE_031_BCOD: out_Is_FofT = False
    '
    '////////////////////////////////////////////////////////////////////////////////////////
    '///////  PROPERTY SHEET "AQUATIC TOXICITY 1": DEFAULT TECHNIQUES  //////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '
    Case PROPCODE_041_FMINNOW_48H_EC50: out_Is_FofT = False
    Case PROPCODE_042_FMINNOW_96H_EC50: out_Is_FofT = False
    Case PROPCODE_043_FMINNOW_24H_LC50: out_Is_FofT = False
    Case PROPCODE_044_FMINNOW_48H_LC50: out_Is_FofT = False
    Case PROPCODE_045_FMINNOW_96H_LC50: out_Is_FofT = False
    Case PROPCODE_046_SALMONIDAE_24H_LC50: out_Is_FofT = False
    Case PROPCODE_047_SALMONIDAE_48H_LC50: out_Is_FofT = False
    Case PROPCODE_048_SALMONIDAE_96H_LC50: out_Is_FofT = False
    '
    '////////////////////////////////////////////////////////////////////////////////////////
    '///////  PROPERTY SHEET "AQUATIC TOXICITY 2": DEFAULT TECHNIQUES  //////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '
    Case PROPCODE_049_DMAGNA_24H_EC50: out_Is_FofT = False
    Case PROPCODE_050_DMAGNA_48H_EC50: out_Is_FofT = False
    Case PROPCODE_051_DMAGNA_24H_LC50: out_Is_FofT = False
    Case PROPCODE_052_DMAGNA_48H_LC50: out_Is_FofT = False
    Case PROPCODE_053_MYSID_96H_LC50: out_Is_FofT = False
    Case PROPCODE_054_ALTERNATE_SPECIES: out_Is_FofT = False
    Case Else:
      '
      ' THE VARIABLE out_Is_FofT DEFAULTS TO FALSE, AND AN ERROR
      ' MESSAGE IS RETURNED.
      '
      GoTo exit_err_ThisFunc
  End Select
exit_normally_ThisFunc:
  Given_PropCode_Get_Is_FofT = True
  Exit Function
exit_err_ThisFunc:
  Given_PropCode_Get_Is_FofT = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("Given_PropCode_Get_Is_FofT")
  Resume exit_err_ThisFunc
End Function
Function Given_PropCode_Get_UnitType_and_UnitBase( _
    in_PropCode As Long, _
    out_UnitType As String, _
    out_UnitBase As String) _
    As Boolean
On Error GoTo err_ThisFunc
  out_UnitBase = ""
  Select Case in_PropCode
    '
    '////////////////////////////////////////////////////////////////////////////////////////
    '///////  PROPERTY SHEET "GENERAL 1": DEFAULT TECHNIQUES  ///////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '
    Case PROPCODE_000_MOLEC_WEIGHT: out_UnitBase = "g/gmol": out_UnitType = "molecular_weight"
    Case PROPCODE_001_LIQDENS_298K: out_UnitBase = "kg/m³": out_UnitType = "density"
    Case PROPCODE_002_LIQDENS_FOFT: out_UnitBase = "kmol/m³": out_UnitType = "molar_density"
    Case PROPCODE_003_MELTING_POINT: out_UnitBase = "K": out_UnitType = "temperature"
    Case PROPCODE_004_NBP: out_UnitBase = "K": out_UnitType = "temperature"
    Case PROPCODE_005_VP_298K: out_UnitBase = "Pa": out_UnitType = "pressure"
    Case PROPCODE_006_VP_FOFT: out_UnitBase = "": out_UnitType = ""
    Case PROPCODE_007_HEAT_FORMATION: out_UnitBase = "J/kmol": out_UnitType = "molar_energy"
    '
    '////////////////////////////////////////////////////////////////////////////////////////
    '///////  PROPERTY SHEET "GENERAL 2": DEFAULT TECHNIQUES  ///////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '
    Case PROPCODE_008_LIQUID_HEAT_CAPACITY_FOFT: out_UnitBase = "J/kmol-K": out_UnitType = "molar_thermal_energy"
    Case PROPCODE_009_VAPOR_HEAT_CAPACITY_FOFT: out_UnitBase = "J/kmol-K": out_UnitType = "molar_thermal_energy"
    Case PROPCODE_010_HEAT_OF_VAPORIZATION_298K: out_UnitBase = "J/kmol": out_UnitType = "molar_energy"
    Case PROPCODE_011_HEAT_OF_VAPORIZATION_NBP: out_UnitBase = "J/kmol": out_UnitType = "molar_energy"
    Case PROPCODE_012_HEAT_OF_VAPORIZATION_FOFT: out_UnitBase = "J/kmol": out_UnitType = "molar_energy"
    Case PROPCODE_013_CRITICAL_T: out_UnitBase = "K": out_UnitType = "temperature"
    Case PROPCODE_014_CRITICAL_P: out_UnitBase = "Pa": out_UnitType = "pressure"
    Case PROPCODE_038_CRITICAL_V: out_UnitBase = "m³/kmol": out_UnitType = "molar_volume"
    '
    '////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '///////  PROPERTY SHEET "TRANSPORT": DEFAULT TECHNIQUES  ///////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '
    Case PROPCODE_015_DIFFUSIVITY_H2O: out_UnitBase = "cm²/s": out_UnitType = "diffusivity"
    Case PROPCODE_016_DIFFUSIVITY_AIR: out_UnitBase = "cm²/s": out_UnitType = "diffusivity"
    Case PROPCODE_017_SURFACE_TENSION_298K: out_UnitBase = "N/m": out_UnitType = "surface_tension"
    Case PROPCODE_018_SURFACE_TENSION_FOFT: out_UnitBase = "N/m": out_UnitType = "surface_tension"
    Case PROPCODE_019_VAPOR_VISCOSITY_FOFT: out_UnitBase = "Pa*s": out_UnitType = "viscosity"
    Case PROPCODE_020_LIQUID_VISCOSITY_FOFT: out_UnitBase = "Pa*s": out_UnitType = "viscosity"
    Case PROPCODE_021_LIQUID_THERMAL_CONDUC_FOFT: out_UnitBase = "W/m/K": out_UnitType = "thermal_conductivity"
    Case PROPCODE_022_VAPOR_THERMAL_CONDUC_FOFT: out_UnitBase = "W/m/K": out_UnitType = "thermal_conductivity"
    '
    '////////////////////////////////////////////////////////////////////////////////////////
    '///////  PROPERTY SHEET "PARTITIONING/EQUILIBRIUM": DEFAULT TECHNIQUES  ////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '
    Case PROPCODE_034_AC_CHEM_IN_H2O: out_UnitBase = "unit-less": out_UnitType = "activity_coefficient"
    Case PROPCODE_032_AC_H2O_IN_CHEM: out_UnitBase = "unit-less": out_UnitType = "activity_coefficient"
    Case PROPCODE_033_HENRY_CONSTANT: out_UnitBase = "": out_UnitType = ""
    Case PROPCODE_039_SOL_LIMIT_CHEM_IN_H2O: out_UnitBase = "": out_UnitType = ""
    Case PROPCODE_035_LOG_KOW: out_UnitBase = "unit-less": out_UnitType = "log_kow"
    Case PROPCODE_036_LOG_KOC: out_UnitBase = "cm3/g OC": out_UnitType = "log_koc"
    Case PROPCODE_037_BIOCONC_FACTOR: out_UnitBase = "unit-less": out_UnitType = "bioconc_factor"
    '
    '////////////////////////////////////////////////////////////////////////////////////////
    '///////  PROPERTY SHEET "FIRE AND EXPLOSION": DEFAULT TECHNIQUES  //////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '
    Case PROPCODE_023_UF_LIMIT: out_UnitBase = "vol% in air": out_UnitType = "percent_volume"
    Case PROPCODE_024_LF_LIMIT: out_UnitBase = "vol% in air": out_UnitType = "percent_volume"
    Case PROPCODE_025_FLASH_POINT: out_UnitBase = "K": out_UnitType = "temperature"
    Case PROPCODE_026_AUTOIGNITION_T: out_UnitBase = "K": out_UnitType = "temperature"
    Case PROPCODE_027_COMBUSTION_HEAT: out_UnitBase = "J/kmol": out_UnitType = "molar_energy"
    '
    '////////////////////////////////////////////////////////////////////////////////////////
    '///////  PROPERTY SHEET "OXYGEN DEMAND": DEFAULT TECHNIQUES  ///////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '
    Case PROPCODE_028_CARBON_THOD: out_UnitBase = "g O2/g chem": out_UnitType = "oxygen_demand"
    Case PROPCODE_029_COMBINED_THOD: out_UnitBase = "g O2/g chem": out_UnitType = "oxygen_demand"
    Case PROPCODE_030_COD: out_UnitBase = "g O2/g chem": out_UnitType = "oxygen_demand"
    Case PROPCODE_031_BCOD: out_UnitBase = "g O2/g chem": out_UnitType = "oxygen_demand"
    '
    '////////////////////////////////////////////////////////////////////////////////////////
    '///////  PROPERTY SHEET "AQUATIC TOXICITY 1": DEFAULT TECHNIQUES  //////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '
    Case PROPCODE_041_FMINNOW_48H_EC50: out_UnitBase = "mg/L": out_UnitType = "concentration"
    Case PROPCODE_042_FMINNOW_96H_EC50: out_UnitBase = "mg/L": out_UnitType = "concentration"
    Case PROPCODE_043_FMINNOW_24H_LC50: out_UnitBase = "mg/L": out_UnitType = "concentration"
    Case PROPCODE_044_FMINNOW_48H_LC50: out_UnitBase = "mg/L": out_UnitType = "concentration"
    Case PROPCODE_045_FMINNOW_96H_LC50: out_UnitBase = "mg/L": out_UnitType = "concentration"
    Case PROPCODE_046_SALMONIDAE_24H_LC50: out_UnitBase = "mg/L": out_UnitType = "concentration"
    Case PROPCODE_047_SALMONIDAE_48H_LC50: out_UnitBase = "mg/L": out_UnitType = "concentration"
    Case PROPCODE_048_SALMONIDAE_96H_LC50: out_UnitBase = "mg/L": out_UnitType = "concentration"
    '
    '////////////////////////////////////////////////////////////////////////////////////////
    '///////  PROPERTY SHEET "AQUATIC TOXICITY 2": DEFAULT TECHNIQUES  //////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '
    Case PROPCODE_049_DMAGNA_24H_EC50: out_UnitBase = "mg/L": out_UnitType = "concentration"
    Case PROPCODE_050_DMAGNA_48H_EC50: out_UnitBase = "mg/L": out_UnitType = "concentration"
    Case PROPCODE_051_DMAGNA_24H_LC50: out_UnitBase = "mg/L": out_UnitType = "concentration"
    Case PROPCODE_052_DMAGNA_48H_LC50: out_UnitBase = "mg/L": out_UnitType = "concentration"
    Case PROPCODE_053_MYSID_96H_LC50: out_UnitBase = "mg/L": out_UnitType = "concentration"
    Case PROPCODE_054_ALTERNATE_SPECIES: out_UnitBase = "mg/L": out_UnitType = "concentration"
    Case Else:
      '
      ' THE VARIABLE XXXX DEFAULTS TO "", INDICATING
      ' AN ERROR MESSAGE RETURNED.
      '
  End Select
  If (out_UnitBase = "") Then
    out_UnitBase = "(Invalid Property!)"
    GoTo exit_err_ThisFunc
  End If
exit_normally_ThisFunc:
  Given_PropCode_Get_UnitType_and_UnitBase = True
  Exit Function
exit_err_ThisFunc:
  Given_PropCode_Get_UnitType_and_UnitBase = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("Given_PropCode_Get_UnitType_and_UnitBase")
  Resume exit_err_ThisFunc
End Function
Function Given_PropCode_Get_Validity( _
    in_PropCode As Long, _
    out_IsValid As Boolean) _
    As Boolean
On Error GoTo err_ThisFunc
Dim out_Name As String
  If (False = _
      Given_PropCode_Get_Name(in_PropCode, out_Name)) Then
    '
    ' INVALID PROPERTY CODE!
    GoTo exit_err_ThisFunc
  End If
  '
  ' VALID PROPERTY CODE.
  GoTo exit_normally_ThisFunc
exit_normally_ThisFunc:
  Given_PropCode_Get_Validity = True
  Exit Function
exit_err_ThisFunc:
  Given_PropCode_Get_Validity = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("Given_PropCode_Get_Validity")
  Resume exit_err_ThisFunc
End Function
Function Given_PropCode_Get_Name( _
    in_PropCode As Long, _
    out_Name As String) _
    As Boolean
On Error GoTo err_ThisFunc
  out_Name = ""
  Select Case in_PropCode
    '
    '////////////////////////////////////////////////////////////////////////////////////////
    '///////  PROPERTY SHEET "GENERAL 1": DEFAULT TECHNIQUES  ///////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '
    Case PROPCODE_000_MOLEC_WEIGHT: out_Name = "Molecular Weight"
    Case PROPCODE_001_LIQDENS_298K: out_Name = "Liquid Density @ 298.15 K"
    Case PROPCODE_002_LIQDENS_FOFT: out_Name = "Liquid Density as f(T)"
    Case PROPCODE_003_MELTING_POINT: out_Name = "Melting Point"
    Case PROPCODE_004_NBP: out_Name = "Normal Boiling Point (NBP)"
    Case PROPCODE_005_VP_298K: out_Name = "Vapor Pressure @ 298.15 K"
    Case PROPCODE_006_VP_FOFT: out_Name = "Vapor Pressure as f(T)"
    Case PROPCODE_007_HEAT_FORMATION: out_Name = "Heat of Formation"
    '
    '////////////////////////////////////////////////////////////////////////////////////////
    '///////  PROPERTY SHEET "GENERAL 2": DEFAULT TECHNIQUES  ///////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '
    Case PROPCODE_008_LIQUID_HEAT_CAPACITY_FOFT: out_Name = "Liquid Heat Capacity as f(T)"
    Case PROPCODE_009_VAPOR_HEAT_CAPACITY_FOFT: out_Name = "Vapor Heat Capacity as f(T)"
    Case PROPCODE_010_HEAT_OF_VAPORIZATION_298K: out_Name = "Heat of Vaporization @ 298.15 K"
    Case PROPCODE_011_HEAT_OF_VAPORIZATION_NBP: out_Name = "Heat of Vaporization @ NBP"
    Case PROPCODE_012_HEAT_OF_VAPORIZATION_FOFT: out_Name = "Heat of Vaporization as f(T)"
    Case PROPCODE_013_CRITICAL_T: out_Name = "Critical Temperature"
    Case PROPCODE_014_CRITICAL_P: out_Name = "Critical Pressure"
    Case PROPCODE_038_CRITICAL_V: out_Name = "Critical Volume"
    '
    '////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '///////  PROPERTY SHEET "TRANSPORT": DEFAULT TECHNIQUES  ///////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '
    Case PROPCODE_015_DIFFUSIVITY_H2O: out_Name = "Diffusivity in Water"
    Case PROPCODE_016_DIFFUSIVITY_AIR: out_Name = "Diffusivity in Air"
    Case PROPCODE_017_SURFACE_TENSION_298K: out_Name = "Surface Tension @ 298.15 K"
    Case PROPCODE_018_SURFACE_TENSION_FOFT: out_Name = "Surface Tension as f(T)"
    Case PROPCODE_019_VAPOR_VISCOSITY_FOFT: out_Name = "Vapor Viscosity as f(T)"
    Case PROPCODE_020_LIQUID_VISCOSITY_FOFT: out_Name = "Liquid Viscosity as f(T)"
    Case PROPCODE_021_LIQUID_THERMAL_CONDUC_FOFT: out_Name = "Liquid Thermal Conductivity as f(T)"
    Case PROPCODE_022_VAPOR_THERMAL_CONDUC_FOFT: out_Name = "Vapor Thermal Conductivity as f(T)"
    '
    '////////////////////////////////////////////////////////////////////////////////////////
    '///////  PROPERTY SHEET "PARTITIONING/EQUILIBRIUM": DEFAULT TECHNIQUES  ////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '
    Case PROPCODE_034_AC_CHEM_IN_H2O: out_Name = "Activity Coefficient of Chemical in Water"
    Case PROPCODE_032_AC_H2O_IN_CHEM: out_Name = "Activity Coefficient of Water in Chemical"
    Case PROPCODE_033_HENRY_CONSTANT: out_Name = "Henry's Constant"
    Case PROPCODE_039_SOL_LIMIT_CHEM_IN_H2O: out_Name = "Solubility Limit of Chemical in Water"
    Case PROPCODE_035_LOG_KOW: out_Name = "Log Kow"
    Case PROPCODE_036_LOG_KOC: out_Name = "Log Koc"
    Case PROPCODE_037_BIOCONC_FACTOR: out_Name = "Bioconcentration Factor"
    '
    '////////////////////////////////////////////////////////////////////////////////////////
    '///////  PROPERTY SHEET "FIRE AND EXPLOSION": DEFAULT TECHNIQUES  //////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '
    Case PROPCODE_023_UF_LIMIT: out_Name = "Upper Flammability Limit"
    Case PROPCODE_024_LF_LIMIT: out_Name = "Lower Flammability Limit"
    Case PROPCODE_025_FLASH_POINT: out_Name = "Flash Point"
    Case PROPCODE_026_AUTOIGNITION_T: out_Name = "Autoignition Temperature"
    Case PROPCODE_027_COMBUSTION_HEAT: out_Name = "Heat of Combustion"
    '
    '////////////////////////////////////////////////////////////////////////////////////////
    '///////  PROPERTY SHEET "OXYGEN DEMAND": DEFAULT TECHNIQUES  ///////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '
    Case PROPCODE_028_CARBON_THOD: out_Name = "Carbonaceous ThOD"
    Case PROPCODE_029_COMBINED_THOD: out_Name = "Combined (C + N) ThOD"
    Case PROPCODE_030_COD: out_Name = "Chemical Oxygen Demand"
    Case PROPCODE_031_BCOD: out_Name = "Biochemical Oxygen Demand"
    '
    '////////////////////////////////////////////////////////////////////////////////////////
    '///////  PROPERTY SHEET "AQUATIC TOXICITY 1": DEFAULT TECHNIQUES  //////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '
    Case PROPCODE_041_FMINNOW_48H_EC50: out_Name = "Flathead Minnow, 48h, EC50"
    Case PROPCODE_042_FMINNOW_96H_EC50: out_Name = "Flathead Minnow, 96h, EC50"
    Case PROPCODE_043_FMINNOW_24H_LC50: out_Name = "Flathead Minnow, 24h, LC50"
    Case PROPCODE_044_FMINNOW_48H_LC50: out_Name = "Flathead Minnow, 48h, LC50"
    Case PROPCODE_045_FMINNOW_96H_LC50: out_Name = "Flathead Minnow, 96h, LC50"
    Case PROPCODE_046_SALMONIDAE_24H_LC50: out_Name = "Salmonidae, 24h, LC50"
    Case PROPCODE_047_SALMONIDAE_48H_LC50: out_Name = "Salmonidae, 48h, LC50"
    Case PROPCODE_048_SALMONIDAE_96H_LC50: out_Name = "Salmonidae, 96h, LC50"
    '
    '////////////////////////////////////////////////////////////////////////////////////////
    '///////  PROPERTY SHEET "AQUATIC TOXICITY 2": DEFAULT TECHNIQUES  //////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '
    Case PROPCODE_049_DMAGNA_24H_EC50: out_Name = "Daphnia Magna, 24h, EC50"
    Case PROPCODE_050_DMAGNA_48H_EC50: out_Name = "Daphnia Magna, 48h, EC50"
    Case PROPCODE_051_DMAGNA_24H_LC50: out_Name = "Daphnia Magna, 24h, LC50"
    Case PROPCODE_052_DMAGNA_48H_LC50: out_Name = "Daphnia Magna, 48h, LC50"
    Case PROPCODE_053_MYSID_96H_LC50: out_Name = "Mysid, 96h, LC50"
    Case PROPCODE_054_ALTERNATE_SPECIES: out_Name = "Alternate Species"
    Case Else:
      '
      ' THE VARIABLE XXXX DEFAULTS TO "", INDICATING
      ' AN ERROR MESSAGE RETURNED.
      '
  End Select
  If (out_Name = "") Then
    out_Name = "(Invalid Property!)"
    GoTo exit_err_ThisFunc
  End If
exit_normally_ThisFunc:
  Given_PropCode_Get_Name = True
  Exit Function
exit_err_ThisFunc:
  Given_PropCode_Get_Name = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("Given_PropCode_Get_Name")
  Resume exit_err_ThisFunc
End Function


Function Given_TechCode_Get_TechReference( _
    in_TechCode As Long, _
    out_TechReference As String) _
    As Boolean
On Error GoTo err_ThisFunc
  out_TechReference = ""
  
  
  '
  '  MORE TO COME! .......
  '




exit_normally_ThisFunc:
  Given_TechCode_Get_TechReference = True
  Exit Function
exit_err_ThisFunc:
  Given_TechCode_Get_TechReference = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("Given_TechCode_Get_TechReference")
  GoTo exit_err_ThisFunc
End Function
Function Given_TechCode_Get_TechCategory( _
    in_TechCode As Long, _
    out_TechCategory As Integer) _
    As Boolean
On Error GoTo err_ThisFunc
  out_TechCategory = -1
  Select Case in_TechCode
    '
    '////////////////////////////////////////////////////////////////////////////////////////
    '///////  GENERAL TECHNIQUES  ///////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '
    Case TECHCODE_ANY_000u_USER_INPUT: out_TechCategory = TECHCATEGORY_USER
    Case TECHCODE_ANY_991d_DB911: out_TechCategory = TECHCATEGORY_DATA
    Case TECHCODE_ANY_992d_DB801: out_TechCategory = TECHCATEGORY_DATA
    '
    '////////////////////////////////////////////////////////////////////////////////////////
    '///////  PROPERTY SHEET "GENERAL 1": DEFAULT TECHNIQUES  ///////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '
    Case TECHCODE_000_002e_UNIFAC: out_TechCategory = TECHCATEGORY_ESTIMATE
    Case TECHCODE_001_003e_BHIRUDS_1978: out_TechCategory = TECHCATEGORY_ESTIMATE
    Case TECHCODE_001_004e_RACKETT_1978: out_TechCategory = TECHCATEGORY_ESTIMATE
    Case TECHCODE_003_005e_TAFT_STAREK_1930: out_TechCategory = TECHCATEGORY_ESTIMATE
    Case TECHCODE_003_006e_LORENZ_HERZ_1922: out_TechCategory = TECHCATEGORY_ESTIMATE
    Case TECHCODE_006_007d_ANTOINELIKE_EXPRESSION: out_TechCategory = TECHCATEGORY_ESTIMATE
    '
    '////////////////////////////////////////////////////////////////////////////////////////
    '///////  PROPERTY SHEET "GENERAL 2": DEFAULT TECHNIQUES  ///////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '
    Case TECHCODE_010_008e_WATSON: out_TechCategory = TECHCATEGORY_ESTIMATE
    Case TECHCODE_011_009e_KLEIN_1949: out_TechCategory = TECHCATEGORY_ESTIMATE
    Case TECHCODE_011_010e_CHEN_PITZER_1965: out_TechCategory = TECHCATEGORY_ESTIMATE
    '
    '////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '///////  PROPERTY SHEET "TRANSPORT": DEFAULT TECHNIQUES  ///////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '
    Case TECHCODE_015_011e_HAYDUK_MINHAS_1982: out_TechCategory = TECHCATEGORY_ESTIMATE
    Case TECHCODE_015_012e_HAYDUK_LAUDIE_1974: out_TechCategory = TECHCATEGORY_ESTIMATE
    Case TECHCODE_015_013e_WILKE_CHANG: out_TechCategory = TECHCATEGORY_ESTIMATE
    Case TECHCODE_016_014e_WILKE_LEE_MOD: out_TechCategory = TECHCATEGORY_ESTIMATE
    Case TECHCODE_017_015e_BROCK_BIRD_1983: out_TechCategory = TECHCATEGORY_ESTIMATE
    '
    '////////////////////////////////////////////////////////////////////////////////////////
    '///////  PROPERTY SHEET "PARTITIONING/EQUILIBRIUM": DEFAULT TECHNIQUES  ////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '
    Case TECHCODE_034_016e_UNIFAC: out_TechCategory = TECHCATEGORY_ESTIMATE
    Case TECHCODE_034_017e_HANSCH_1968: out_TechCategory = TECHCATEGORY_ESTIMATE
    Case TECHCODE_032_018e_UNIFAC: out_TechCategory = TECHCATEGORY_ESTIMATE
    Case TECHCODE_039_020d_YAWS: out_TechCategory = TECHCATEGORY_DATA
    Case TECHCODE_039_019e_UNIFAC: out_TechCategory = TECHCATEGORY_ESTIMATE
    Case TECHCODE_039_021e_YALKOWSKY_1990: out_TechCategory = TECHCATEGORY_ESTIMATE
    Case TECHCODE_035_022e_KENAGA_GORING_1978: out_TechCategory = TECHCATEGORY_ESTIMATE
    Case TECHCODE_036_023e_BAKER_1994: out_TechCategory = TECHCATEGORY_ESTIMATE
    Case TECHCODE_037_024e_KOBAYSHI_1981: out_TechCategory = TECHCATEGORY_ESTIMATE
    Case TECHCODE_037_025e_KENAGA_GORING_1980: out_TechCategory = TECHCATEGORY_ESTIMATE
    '
    '////////////////////////////////////////////////////////////////////////////////////////
    '///////  PROPERTY SHEET "FIRE AND EXPLOSION": DEFAULT TECHNIQUES  //////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '
    Case TECHCODE_023_026d_MTU_FIREEXP_DATA: out_TechCategory = TECHCATEGORY_DATA
    Case TECHCODE_023_027d_MTU_GROUP_CONTRIB: out_TechCategory = TECHCATEGORY_DATA
    Case TECHCODE_023_028d_MTU_COMBUSTION_RXN: out_TechCategory = TECHCATEGORY_DATA
    Case TECHCODE_023_029d_PENN_GROUP_CONTRIB: out_TechCategory = TECHCATEGORY_DATA
    Case TECHCODE_024_030d_MTU_FIREEXP_DATA: out_TechCategory = TECHCATEGORY_DATA
    Case TECHCODE_024_031d_MTU_GROUP_CONTRIB: out_TechCategory = TECHCATEGORY_DATA
    Case TECHCODE_024_032d_PENN_GROUP_CONTRIB: out_TechCategory = TECHCATEGORY_DATA
    Case TECHCODE_024_033d_MTU_COMBUSTION_RXN: out_TechCategory = TECHCATEGORY_DATA
    Case TECHCODE_024_034d_MTU_FLASHPOINT_METH: out_TechCategory = TECHCATEGORY_DATA
    Case TECHCODE_025_035d_MTU_FIREEXP_DATA: out_TechCategory = TECHCATEGORY_DATA
    Case TECHCODE_025_036d_LFL_DATA: out_TechCategory = TECHCATEGORY_DATA
    Case TECHCODE_025_037d_MTU_LFL_GROUP_CONTRIB: out_TechCategory = TECHCATEGORY_DATA
    Case TECHCODE_025_038d_PENN_GROUP_CONTRIB: out_TechCategory = TECHCATEGORY_DATA
    Case TECHCODE_025_039d_MTU_LFL_COMBUSTION_RXN: out_TechCategory = TECHCATEGORY_DATA
    Case TECHCODE_026_040d_MTU_FIREEXP_DATA: out_TechCategory = TECHCATEGORY_DATA
    Case TECHCODE_026_041d_MTU_LOG_METHOD: out_TechCategory = TECHCATEGORY_DATA
    Case TECHCODE_026_042d_MTU_LINEAR_METHOD: out_TechCategory = TECHCATEGORY_DATA
    '
    '////////////////////////////////////////////////////////////////////////////////////////
    '///////  PROPERTY SHEET "OXYGEN DEMAND": DEFAULT TECHNIQUES  ///////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '
    Case TECHCODE_030_043e_MTU_DIPPR: out_TechCategory = TECHCATEGORY_ESTIMATE
    '
    '////////////////////////////////////////////////////////////////////////////////////////
    '///////  PROPERTY SHEET "AQUATIC TOXICITY 1": DEFAULT TECHNIQUES  //////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '
          '
          ' THERE ARE NO UNIQUE TECHNIQUES FOR THIS PROPERTY SHEET.
          '
    '
    '////////////////////////////////////////////////////////////////////////////////////////
    '///////  PROPERTY SHEET "AQUATIC TOXICITY 2": DEFAULT TECHNIQUES  //////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '
          '
          ' THERE ARE NO UNIQUE TECHNIQUES FOR THIS PROPERTY SHEET.
          '
    Case Else:
      '
      ' THE VARIABLE out_TechCategory DEFAULTS TO -1, INDICATING
      ' AN ERROR MESSAGE RETURNED.
      '
  End Select
  If (out_TechCategory = -1) Then
    GoTo exit_err_ThisFunc
  End If
exit_normally_ThisFunc:
  Given_TechCode_Get_TechCategory = True
  Exit Function
exit_err_ThisFunc:
  Given_TechCode_Get_TechCategory = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("Given_TechCode_Get_TechCategory")
  GoTo exit_err_ThisFunc
End Function
Function Given_TechCode_Get_Validity( _
    in_TechCode As Long, _
    out_IsValid As Boolean) _
    As Boolean
On Error GoTo err_ThisFunc
Dim out_Name As String
  If (False = _
      Given_TechCode_Get_Name(in_TechCode, out_Name)) Then
    '
    ' INVALID TECHNIQUE CODE!
    GoTo exit_err_ThisFunc
  End If
  '
  ' VALID TECHNIQUE CODE.
  GoTo exit_normally_ThisFunc
exit_normally_ThisFunc:
  Given_TechCode_Get_Validity = True
  Exit Function
exit_err_ThisFunc:
  Given_TechCode_Get_Validity = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("Given_TechCode_Get_Validity")
  Resume exit_err_ThisFunc
End Function
Function Given_TechCode_Get_Name( _
    in_TechCode As Long, _
    out_Name As String) _
    As Boolean
On Error GoTo err_ThisFunc
  out_Name = ""
  Select Case in_TechCode
    '
    '////////////////////////////////////////////////////////////////////////////////////////
    '///////  GENERAL TECHNIQUES  ///////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '
    Case TECHCODE_ANY_000u_USER_INPUT: out_Name = "User Input"
    Case TECHCODE_ANY_991d_DB911: out_Name = "DIPPR911 Database"
    Case TECHCODE_ANY_992d_DB801: out_Name = "DIPPR801 Database"
    '
    '////////////////////////////////////////////////////////////////////////////////////////
    '///////  PROPERTY SHEET "GENERAL 1": DEFAULT TECHNIQUES  ///////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '
    Case TECHCODE_000_002e_UNIFAC: out_Name = "UNIFAC"
    Case TECHCODE_001_003e_BHIRUDS_1978: out_Name = "Bhiruds (1978)"
    Case TECHCODE_001_004e_RACKETT_1978: out_Name = "Modified Rackett (1978)"
    Case TECHCODE_003_005e_TAFT_STAREK_1930: out_Name = "Taft and Starek (1930)"
    Case TECHCODE_003_006e_LORENZ_HERZ_1922: out_Name = "Lorenz and Herz (1922)"
    Case TECHCODE_006_007d_ANTOINELIKE_EXPRESSION: out_Name = "{ Antoine-like data }"
    '
    '////////////////////////////////////////////////////////////////////////////////////////
    '///////  PROPERTY SHEET "GENERAL 2": DEFAULT TECHNIQUES  ///////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '
    Case TECHCODE_010_008e_WATSON: out_Name = "Watson"
    Case TECHCODE_011_009e_KLEIN_1949: out_Name = "Klein (1949)"
    Case TECHCODE_011_010e_CHEN_PITZER_1965: out_Name = "Chen and Pitzer (1965)"
    '
    '////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '///////  PROPERTY SHEET "TRANSPORT": DEFAULT TECHNIQUES  ///////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '
    Case TECHCODE_015_011e_HAYDUK_MINHAS_1982: out_Name = "Hayduk and Minhas (1982)"
    Case TECHCODE_015_012e_HAYDUK_LAUDIE_1974: out_Name = "Hayduk and Laudie (1974)"
    Case TECHCODE_015_013e_WILKE_CHANG: out_Name = "Wilke and Chang"
    Case TECHCODE_016_014e_WILKE_LEE_MOD: out_Name = "Wilke and Lee Modification"
    Case TECHCODE_017_015e_BROCK_BIRD_1983: out_Name = "Brock and Bird (1983)"
    '
    '////////////////////////////////////////////////////////////////////////////////////////
    '///////  PROPERTY SHEET "PARTITIONING/EQUILIBRIUM": DEFAULT TECHNIQUES  ////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '
    Case TECHCODE_034_016e_UNIFAC: out_Name = "UNIFAC"
    Case TECHCODE_034_017e_HANSCH_1968: out_Name = "Hansch (1968)"
    Case TECHCODE_032_018e_UNIFAC: out_Name = "UNIFAC"
    Case TECHCODE_039_020d_YAWS: out_Name = "Yaws"
    Case TECHCODE_039_019e_UNIFAC: out_Name = "UNIFAC"
    Case TECHCODE_039_021e_YALKOWSKY_1990: out_Name = "Yalkowsky (1990)"
    Case TECHCODE_035_022e_KENAGA_GORING_1978: out_Name = "Kenaga and Goring (1978)"
    Case TECHCODE_036_023e_BAKER_1994: out_Name = "Baker (1994)"
    Case TECHCODE_037_024e_KOBAYSHI_1981: out_Name = "Kobayshi (1981)"
    Case TECHCODE_037_025e_KENAGA_GORING_1980: out_Name = "Kenaga and Goring (1980)"
    '
    '////////////////////////////////////////////////////////////////////////////////////////
    '///////  PROPERTY SHEET "FIRE AND EXPLOSION": DEFAULT TECHNIQUES  //////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '
    Case TECHCODE_023_026d_MTU_FIREEXP_DATA: out_Name = "MTU Fire & Explosion Data"
    Case TECHCODE_023_027d_MTU_GROUP_CONTRIB: out_Name = "MTU Group Contribution"
    Case TECHCODE_023_028d_MTU_COMBUSTION_RXN: out_Name = "MTU Combustion Reaction"
    Case TECHCODE_023_029d_PENN_GROUP_CONTRIB: out_Name = "Penn State Group Contribution"
    Case TECHCODE_024_030d_MTU_FIREEXP_DATA: out_Name = "MTU Fire & Explosion Data"
    Case TECHCODE_024_031d_MTU_GROUP_CONTRIB: out_Name = "MTU Group Contribution"
    Case TECHCODE_024_032d_PENN_GROUP_CONTRIB: out_Name = "Penn State Group Contribution"
    Case TECHCODE_024_033d_MTU_COMBUSTION_RXN: out_Name = "MTU Combustion Reaction"
    Case TECHCODE_024_034d_MTU_FLASHPOINT_METH: out_Name = "MTU FlashPoint Method"
    Case TECHCODE_025_035d_MTU_FIREEXP_DATA: out_Name = "MTU Fire & Explosion Data"
    Case TECHCODE_025_036d_LFL_DATA: out_Name = "LFL Data"
    Case TECHCODE_025_037d_MTU_LFL_GROUP_CONTRIB: out_Name = "MTU LFL Group Contribution"
    Case TECHCODE_025_038d_PENN_GROUP_CONTRIB: out_Name = "Penn State Group Contribution"
    Case TECHCODE_025_039d_MTU_LFL_COMBUSTION_RXN: out_Name = "MTU LFL Combustion Reaction"
    Case TECHCODE_026_040d_MTU_FIREEXP_DATA: out_Name = "MTU Fire & Explosion Data"
    Case TECHCODE_026_041d_MTU_LOG_METHOD: out_Name = "MTU Logarithmic Method"
    Case TECHCODE_026_042d_MTU_LINEAR_METHOD: out_Name = "MTU Linear Method"
    '
    '////////////////////////////////////////////////////////////////////////////////////////
    '///////  PROPERTY SHEET "OXYGEN DEMAND": DEFAULT TECHNIQUES  ///////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '
    Case TECHCODE_030_043e_MTU_DIPPR: out_Name = "MTU DIPPR"
    '
    '////////////////////////////////////////////////////////////////////////////////////////
    '///////  PROPERTY SHEET "AQUATIC TOXICITY 1": DEFAULT TECHNIQUES  //////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '
          '
          ' THERE ARE NO UNIQUE TECHNIQUES FOR THIS PROPERTY SHEET.
          '
    '
    '////////////////////////////////////////////////////////////////////////////////////////
    '///////  PROPERTY SHEET "AQUATIC TOXICITY 2": DEFAULT TECHNIQUES  //////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '
          '
          ' THERE ARE NO UNIQUE TECHNIQUES FOR THIS PROPERTY SHEET.
          '
    Case Else:
      '
      ' THE VARIABLE out_Name DEFAULTS TO "", INDICATING
      ' AN ERROR MESSAGE RETURNED.
      '
  End Select
  If (out_Name = "") Then
    out_Name = "(Invalid Technique!)"
    GoTo exit_err_ThisFunc
  End If
exit_normally_ThisFunc:
  Given_TechCode_Get_Name = True
  Exit Function
exit_err_ThisFunc:
  Given_TechCode_Get_Name = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("Given_TechCode_Get_Name")
  Resume exit_err_ThisFunc
End Function


Function sc_ElemFind( _
    in_List_PropCodes() As Long, _
    in_ThisCode As Long, _
    out_idx_Elem As Integer) _
    As Boolean
Dim i As Integer
Dim UB As Integer
  UB = UBound(in_List_PropCodes)
  For i = 1 To UB
    If (in_List_PropCodes(i) = in_ThisCode) Then
      out_idx_Elem = i
      sc_ElemFind = True
      Exit Function
    End If
  Next i
  out_idx_Elem = -1
  sc_ElemFind = False
End Function
Function sc_ElemAdd( _
    inout_List_PropCodes() As Long, _
    in_ThisCode As Long) _
    As Boolean
Dim UB As Integer
  UB = UBound(inout_List_PropCodes)
  If (UB = 0) Then
    ReDim inout_List_PropCodes(1 To 1)
  Else
    ReDim Preserve inout_List_PropCodes(1 To UB + 1)
  End If
  inout_List_PropCodes(UB + 1) = in_ThisCode
  sc_ElemAdd = True
End Function


Function Get_Complete_List_of_PropCodes( _
    out_List_PropCodes() As Long) _
    As Boolean
On Error GoTo err_ThisFunc
Dim aryThis() As Long
  ReDim aryThis(0 To 0)
  Call sc_ElemAdd(aryThis, PROPCODE_000_MOLEC_WEIGHT)
  Call sc_ElemAdd(aryThis, PROPCODE_001_LIQDENS_298K)
  Call sc_ElemAdd(aryThis, PROPCODE_002_LIQDENS_FOFT)
  Call sc_ElemAdd(aryThis, PROPCODE_003_MELTING_POINT)
  Call sc_ElemAdd(aryThis, PROPCODE_004_NBP)
  Call sc_ElemAdd(aryThis, PROPCODE_005_VP_298K)
  Call sc_ElemAdd(aryThis, PROPCODE_006_VP_FOFT)
  Call sc_ElemAdd(aryThis, PROPCODE_007_HEAT_FORMATION)
  Call sc_ElemAdd(aryThis, PROPCODE_008_LIQUID_HEAT_CAPACITY_FOFT)
  Call sc_ElemAdd(aryThis, PROPCODE_009_VAPOR_HEAT_CAPACITY_FOFT)
  Call sc_ElemAdd(aryThis, PROPCODE_010_HEAT_OF_VAPORIZATION_298K)
  Call sc_ElemAdd(aryThis, PROPCODE_011_HEAT_OF_VAPORIZATION_NBP)
  Call sc_ElemAdd(aryThis, PROPCODE_012_HEAT_OF_VAPORIZATION_FOFT)
  Call sc_ElemAdd(aryThis, PROPCODE_013_CRITICAL_T)
  Call sc_ElemAdd(aryThis, PROPCODE_014_CRITICAL_P)
  Call sc_ElemAdd(aryThis, PROPCODE_038_CRITICAL_V)
  Call sc_ElemAdd(aryThis, PROPCODE_015_DIFFUSIVITY_H2O)
  Call sc_ElemAdd(aryThis, PROPCODE_016_DIFFUSIVITY_AIR)
  Call sc_ElemAdd(aryThis, PROPCODE_017_SURFACE_TENSION_298K)
  Call sc_ElemAdd(aryThis, PROPCODE_018_SURFACE_TENSION_FOFT)
  Call sc_ElemAdd(aryThis, PROPCODE_019_VAPOR_VISCOSITY_FOFT)
  Call sc_ElemAdd(aryThis, PROPCODE_020_LIQUID_VISCOSITY_FOFT)
  Call sc_ElemAdd(aryThis, PROPCODE_021_LIQUID_THERMAL_CONDUC_FOFT)
  Call sc_ElemAdd(aryThis, PROPCODE_022_VAPOR_THERMAL_CONDUC_FOFT)
  Call sc_ElemAdd(aryThis, PROPCODE_034_AC_CHEM_IN_H2O)
  Call sc_ElemAdd(aryThis, PROPCODE_032_AC_H2O_IN_CHEM)
  Call sc_ElemAdd(aryThis, PROPCODE_033_HENRY_CONSTANT)
  Call sc_ElemAdd(aryThis, PROPCODE_039_SOL_LIMIT_CHEM_IN_H2O)
  Call sc_ElemAdd(aryThis, PROPCODE_035_LOG_KOW)
  Call sc_ElemAdd(aryThis, PROPCODE_036_LOG_KOC)
  Call sc_ElemAdd(aryThis, PROPCODE_037_BIOCONC_FACTOR)
  Call sc_ElemAdd(aryThis, PROPCODE_023_UF_LIMIT)
  Call sc_ElemAdd(aryThis, PROPCODE_024_LF_LIMIT)
  Call sc_ElemAdd(aryThis, PROPCODE_025_FLASH_POINT)
  Call sc_ElemAdd(aryThis, PROPCODE_026_AUTOIGNITION_T)
  Call sc_ElemAdd(aryThis, PROPCODE_027_COMBUSTION_HEAT)
  Call sc_ElemAdd(aryThis, PROPCODE_028_CARBON_THOD)
  Call sc_ElemAdd(aryThis, PROPCODE_029_COMBINED_THOD)
  Call sc_ElemAdd(aryThis, PROPCODE_030_COD)
  Call sc_ElemAdd(aryThis, PROPCODE_031_BCOD)
  Call sc_ElemAdd(aryThis, PROPCODE_041_FMINNOW_48H_EC50)
  Call sc_ElemAdd(aryThis, PROPCODE_042_FMINNOW_96H_EC50)
  Call sc_ElemAdd(aryThis, PROPCODE_043_FMINNOW_24H_LC50)
  Call sc_ElemAdd(aryThis, PROPCODE_044_FMINNOW_48H_LC50)
  Call sc_ElemAdd(aryThis, PROPCODE_045_FMINNOW_96H_LC50)
  Call sc_ElemAdd(aryThis, PROPCODE_046_SALMONIDAE_24H_LC50)
  Call sc_ElemAdd(aryThis, PROPCODE_047_SALMONIDAE_48H_LC50)
  Call sc_ElemAdd(aryThis, PROPCODE_048_SALMONIDAE_96H_LC50)
  Call sc_ElemAdd(aryThis, PROPCODE_049_DMAGNA_24H_EC50)
  Call sc_ElemAdd(aryThis, PROPCODE_050_DMAGNA_48H_EC50)
  Call sc_ElemAdd(aryThis, PROPCODE_051_DMAGNA_24H_LC50)
  Call sc_ElemAdd(aryThis, PROPCODE_052_DMAGNA_48H_LC50)
  Call sc_ElemAdd(aryThis, PROPCODE_053_MYSID_96H_LC50)
  Call sc_ElemAdd(aryThis, PROPCODE_054_ALTERNATE_SPECIES)
  out_List_PropCodes = aryThis
  '
  ' EXIT OUTTA HERE.
  '
exit_normally_ThisFunc:
  Get_Complete_List_of_PropCodes = True
  Exit Function
exit_err_ThisFunc:
  Get_Complete_List_of_PropCodes = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("Get_Complete_List_of_PropCodes")
  Resume exit_err_ThisFunc
End Function


Function Get_Complete_List_of_TechCodes( _
    in_PropCode As Long, _
    out_List_TechCodes() As Long) _
    As Boolean
On Error GoTo err_ThisFunc
  Select Case in_PropCode
    '
    '////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '///////  PROPERTY SHEET "GENERAL 1": DEFAULT TECHNIQUES  ///////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '
    Case PROPCODE_000_MOLEC_WEIGHT:
      ReDim out_List_TechCodes(1 To 4)
      out_List_TechCodes(1) = TECHCODE_ANY_000u_USER_INPUT
      out_List_TechCodes(2) = TECHCODE_ANY_991d_DB911
      out_List_TechCodes(3) = TECHCODE_ANY_992d_DB801
      out_List_TechCodes(4) = TECHCODE_000_002e_UNIFAC
    Case PROPCODE_001_LIQDENS_298K:
      ReDim out_List_TechCodes(1 To 4)
      out_List_TechCodes(1) = TECHCODE_ANY_000u_USER_INPUT
      out_List_TechCodes(2) = TECHCODE_ANY_991d_DB911
      out_List_TechCodes(3) = TECHCODE_001_003e_BHIRUDS_1978
      out_List_TechCodes(4) = TECHCODE_001_004e_RACKETT_1978
    Case PROPCODE_002_LIQDENS_FOFT:
      ReDim out_List_TechCodes(1 To 3)
      out_List_TechCodes(1) = TECHCODE_ANY_000u_USER_INPUT
      out_List_TechCodes(2) = TECHCODE_ANY_991d_DB911
      out_List_TechCodes(3) = TECHCODE_ANY_992d_DB801
    Case PROPCODE_003_MELTING_POINT:
      ReDim out_List_TechCodes(1 To 5)
      out_List_TechCodes(1) = TECHCODE_ANY_000u_USER_INPUT
      out_List_TechCodes(2) = TECHCODE_ANY_991d_DB911
      out_List_TechCodes(3) = TECHCODE_ANY_992d_DB801
      out_List_TechCodes(4) = TECHCODE_003_005e_TAFT_STAREK_1930
      out_List_TechCodes(5) = TECHCODE_003_006e_LORENZ_HERZ_1922
    Case PROPCODE_004_NBP:
      ReDim out_List_TechCodes(1 To 3)
      out_List_TechCodes(1) = TECHCODE_ANY_000u_USER_INPUT
      out_List_TechCodes(2) = TECHCODE_ANY_991d_DB911
      out_List_TechCodes(3) = TECHCODE_ANY_992d_DB801
    Case PROPCODE_005_VP_298K:
      ReDim out_List_TechCodes(1 To 2)
      out_List_TechCodes(1) = TECHCODE_ANY_000u_USER_INPUT
      out_List_TechCodes(2) = TECHCODE_ANY_991d_DB911
    Case PROPCODE_006_VP_FOFT:
      ReDim out_List_TechCodes(1 To 2)
      out_List_TechCodes(1) = TECHCODE_ANY_000u_USER_INPUT
      out_List_TechCodes(2) = TECHCODE_006_007d_ANTOINELIKE_EXPRESSION
    Case PROPCODE_007_HEAT_FORMATION:
      ReDim out_List_TechCodes(1 To 3)
      out_List_TechCodes(1) = TECHCODE_ANY_000u_USER_INPUT
      out_List_TechCodes(2) = TECHCODE_ANY_991d_DB911
      out_List_TechCodes(3) = TECHCODE_ANY_992d_DB801
    '
    '////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '///////  PROPERTY SHEET "GENERAL 2": DEFAULT TECHNIQUES  ///////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '
    Case PROPCODE_008_LIQUID_HEAT_CAPACITY_FOFT:
      ReDim out_List_TechCodes(1 To 3)
      out_List_TechCodes(1) = TECHCODE_ANY_000u_USER_INPUT
      out_List_TechCodes(2) = TECHCODE_ANY_991d_DB911
      out_List_TechCodes(3) = TECHCODE_ANY_992d_DB801
    Case PROPCODE_009_VAPOR_HEAT_CAPACITY_FOFT:
      ReDim out_List_TechCodes(1 To 3)
      out_List_TechCodes(1) = TECHCODE_ANY_000u_USER_INPUT
      out_List_TechCodes(2) = TECHCODE_ANY_991d_DB911
      out_List_TechCodes(3) = TECHCODE_ANY_992d_DB801
    Case PROPCODE_010_HEAT_OF_VAPORIZATION_298K:
      ReDim out_List_TechCodes(1 To 3)
      out_List_TechCodes(1) = TECHCODE_ANY_000u_USER_INPUT
      out_List_TechCodes(2) = TECHCODE_ANY_991d_DB911
      out_List_TechCodes(3) = TECHCODE_010_008e_WATSON
    Case PROPCODE_011_HEAT_OF_VAPORIZATION_NBP:
      ReDim out_List_TechCodes(1 To 4)
      out_List_TechCodes(1) = TECHCODE_ANY_000u_USER_INPUT
      out_List_TechCodes(2) = TECHCODE_ANY_991d_DB911
      out_List_TechCodes(3) = TECHCODE_011_009e_KLEIN_1949
      out_List_TechCodes(4) = TECHCODE_011_010e_CHEN_PITZER_1965
    Case PROPCODE_012_HEAT_OF_VAPORIZATION_FOFT:
      ReDim out_List_TechCodes(1 To 3)
      out_List_TechCodes(1) = TECHCODE_ANY_000u_USER_INPUT
      out_List_TechCodes(2) = TECHCODE_ANY_991d_DB911
      out_List_TechCodes(3) = TECHCODE_ANY_992d_DB801
    Case PROPCODE_013_CRITICAL_T:
      ReDim out_List_TechCodes(1 To 3)
      out_List_TechCodes(1) = TECHCODE_ANY_000u_USER_INPUT
      out_List_TechCodes(2) = TECHCODE_ANY_991d_DB911
      out_List_TechCodes(3) = TECHCODE_ANY_992d_DB801
    Case PROPCODE_014_CRITICAL_P:
      ReDim out_List_TechCodes(1 To 3)
      out_List_TechCodes(1) = TECHCODE_ANY_000u_USER_INPUT
      out_List_TechCodes(2) = TECHCODE_ANY_991d_DB911
      out_List_TechCodes(3) = TECHCODE_ANY_992d_DB801
    Case PROPCODE_038_CRITICAL_V:
      ReDim out_List_TechCodes(1 To 3)
      out_List_TechCodes(1) = TECHCODE_ANY_000u_USER_INPUT
      out_List_TechCodes(2) = TECHCODE_ANY_991d_DB911
      out_List_TechCodes(3) = TECHCODE_ANY_992d_DB801
    '
    '////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '///////  PROPERTY SHEET "TRANSPORT": DEFAULT TECHNIQUES  ///////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '
    Case PROPCODE_015_DIFFUSIVITY_H2O:
      ReDim out_List_TechCodes(1 To 5)
      out_List_TechCodes(1) = TECHCODE_ANY_000u_USER_INPUT
      out_List_TechCodes(2) = TECHCODE_ANY_991d_DB911
      out_List_TechCodes(3) = TECHCODE_015_011e_HAYDUK_MINHAS_1982
      out_List_TechCodes(4) = TECHCODE_015_012e_HAYDUK_LAUDIE_1974
      out_List_TechCodes(5) = TECHCODE_015_013e_WILKE_CHANG
    Case PROPCODE_016_DIFFUSIVITY_AIR:
      ReDim out_List_TechCodes(1 To 3)
      out_List_TechCodes(1) = TECHCODE_ANY_000u_USER_INPUT
      out_List_TechCodes(2) = TECHCODE_ANY_991d_DB911
      out_List_TechCodes(3) = TECHCODE_016_014e_WILKE_LEE_MOD
    Case PROPCODE_017_SURFACE_TENSION_298K:
      ReDim out_List_TechCodes(1 To 3)
      out_List_TechCodes(1) = TECHCODE_ANY_000u_USER_INPUT
      out_List_TechCodes(2) = TECHCODE_ANY_991d_DB911
      out_List_TechCodes(3) = TECHCODE_017_015e_BROCK_BIRD_1983
    Case PROPCODE_018_SURFACE_TENSION_FOFT:
      ReDim out_List_TechCodes(1 To 3)
      out_List_TechCodes(1) = TECHCODE_ANY_000u_USER_INPUT
      out_List_TechCodes(2) = TECHCODE_ANY_991d_DB911
      out_List_TechCodes(3) = TECHCODE_ANY_992d_DB801
    Case PROPCODE_019_VAPOR_VISCOSITY_FOFT:
      ReDim out_List_TechCodes(1 To 3)
      out_List_TechCodes(1) = TECHCODE_ANY_000u_USER_INPUT
      out_List_TechCodes(2) = TECHCODE_ANY_991d_DB911
      out_List_TechCodes(3) = TECHCODE_ANY_992d_DB801
    Case PROPCODE_020_LIQUID_VISCOSITY_FOFT:
      ReDim out_List_TechCodes(1 To 3)
      out_List_TechCodes(1) = TECHCODE_ANY_000u_USER_INPUT
      out_List_TechCodes(2) = TECHCODE_ANY_991d_DB911
      out_List_TechCodes(3) = TECHCODE_ANY_992d_DB801
    Case PROPCODE_021_LIQUID_THERMAL_CONDUC_FOFT:
      ReDim out_List_TechCodes(1 To 3)
      out_List_TechCodes(1) = TECHCODE_ANY_000u_USER_INPUT
      out_List_TechCodes(2) = TECHCODE_ANY_991d_DB911
      out_List_TechCodes(3) = TECHCODE_ANY_992d_DB801
    Case PROPCODE_022_VAPOR_THERMAL_CONDUC_FOFT:
      ReDim out_List_TechCodes(1 To 3)
      out_List_TechCodes(1) = TECHCODE_ANY_000u_USER_INPUT
      out_List_TechCodes(2) = TECHCODE_ANY_991d_DB911
      out_List_TechCodes(3) = TECHCODE_ANY_992d_DB801
    '
    '////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '///////  PROPERTY SHEET "PARTITIONING/EQUILIBRIUM": DEFAULT TECHNIQUES  ////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '
    Case PROPCODE_034_AC_CHEM_IN_H2O:
      ReDim out_List_TechCodes(1 To 4)
      out_List_TechCodes(1) = TECHCODE_ANY_000u_USER_INPUT
      out_List_TechCodes(2) = TECHCODE_ANY_991d_DB911
      out_List_TechCodes(3) = TECHCODE_034_016e_UNIFAC
      out_List_TechCodes(4) = TECHCODE_034_017e_HANSCH_1968
    Case PROPCODE_032_AC_H2O_IN_CHEM:
      ReDim out_List_TechCodes(1 To 3)
      out_List_TechCodes(1) = TECHCODE_ANY_000u_USER_INPUT
      out_List_TechCodes(2) = TECHCODE_ANY_991d_DB911
      out_List_TechCodes(3) = TECHCODE_032_018e_UNIFAC
    Case PROPCODE_033_HENRY_CONSTANT:
      ReDim out_List_TechCodes(1 To 2)
      out_List_TechCodes(1) = TECHCODE_ANY_000u_USER_INPUT
      out_List_TechCodes(2) = TECHCODE_ANY_991d_DB911
      '
      ' MORE TECHNIQUES TO COME LATER !!!
      '
    Case PROPCODE_039_SOL_LIMIT_CHEM_IN_H2O:
      ReDim out_List_TechCodes(1 To 5)
      out_List_TechCodes(1) = TECHCODE_ANY_000u_USER_INPUT
      out_List_TechCodes(2) = TECHCODE_ANY_991d_DB911
      out_List_TechCodes(3) = TECHCODE_039_020d_YAWS
      out_List_TechCodes(4) = TECHCODE_039_019e_UNIFAC
      out_List_TechCodes(5) = TECHCODE_039_021e_YALKOWSKY_1990
    Case PROPCODE_035_LOG_KOW:
      ReDim out_List_TechCodes(1 To 3)
      out_List_TechCodes(1) = TECHCODE_ANY_000u_USER_INPUT
      out_List_TechCodes(2) = TECHCODE_ANY_991d_DB911
      out_List_TechCodes(3) = TECHCODE_035_022e_KENAGA_GORING_1978
    Case PROPCODE_036_LOG_KOC:
      ReDim out_List_TechCodes(1 To 3)
      out_List_TechCodes(1) = TECHCODE_ANY_000u_USER_INPUT
      out_List_TechCodes(2) = TECHCODE_ANY_991d_DB911
      out_List_TechCodes(3) = TECHCODE_036_023e_BAKER_1994
    Case PROPCODE_037_BIOCONC_FACTOR:
      ReDim out_List_TechCodes(1 To 4)
      out_List_TechCodes(1) = TECHCODE_ANY_000u_USER_INPUT
      out_List_TechCodes(2) = TECHCODE_ANY_991d_DB911
      out_List_TechCodes(3) = TECHCODE_037_024e_KOBAYSHI_1981
      out_List_TechCodes(4) = TECHCODE_037_025e_KENAGA_GORING_1980
    '
    '////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '///////  PROPERTY SHEET "FIRE AND EXPLOSION": DEFAULT TECHNIQUES  //////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '
    Case PROPCODE_023_UF_LIMIT:
      ReDim out_List_TechCodes(1 To 7)
      out_List_TechCodes(1) = TECHCODE_ANY_000u_USER_INPUT
      out_List_TechCodes(2) = TECHCODE_ANY_991d_DB911
      out_List_TechCodes(3) = TECHCODE_ANY_992d_DB801
      out_List_TechCodes(4) = TECHCODE_023_026d_MTU_FIREEXP_DATA
      out_List_TechCodes(5) = TECHCODE_023_027d_MTU_GROUP_CONTRIB
      out_List_TechCodes(6) = TECHCODE_023_028d_MTU_COMBUSTION_RXN
      out_List_TechCodes(7) = TECHCODE_023_029d_PENN_GROUP_CONTRIB
    Case PROPCODE_024_LF_LIMIT:
      ReDim out_List_TechCodes(1 To 8)
      out_List_TechCodes(1) = TECHCODE_ANY_000u_USER_INPUT
      out_List_TechCodes(2) = TECHCODE_ANY_991d_DB911
      out_List_TechCodes(3) = TECHCODE_ANY_992d_DB801
      out_List_TechCodes(4) = TECHCODE_024_030d_MTU_FIREEXP_DATA
      out_List_TechCodes(5) = TECHCODE_024_031d_MTU_GROUP_CONTRIB
      out_List_TechCodes(6) = TECHCODE_024_032d_PENN_GROUP_CONTRIB
      out_List_TechCodes(7) = TECHCODE_024_033d_MTU_COMBUSTION_RXN
      out_List_TechCodes(8) = TECHCODE_024_034d_MTU_FLASHPOINT_METH
    Case PROPCODE_025_FLASH_POINT:
      ReDim out_List_TechCodes(1 To 8)
      out_List_TechCodes(1) = TECHCODE_ANY_000u_USER_INPUT
      out_List_TechCodes(2) = TECHCODE_ANY_991d_DB911
      out_List_TechCodes(3) = TECHCODE_ANY_992d_DB801
      out_List_TechCodes(4) = TECHCODE_025_035d_MTU_FIREEXP_DATA
      out_List_TechCodes(5) = TECHCODE_025_036d_LFL_DATA
      out_List_TechCodes(6) = TECHCODE_025_037d_MTU_LFL_GROUP_CONTRIB
      out_List_TechCodes(7) = TECHCODE_025_038d_PENN_GROUP_CONTRIB
      out_List_TechCodes(8) = TECHCODE_025_039d_MTU_LFL_COMBUSTION_RXN
    Case PROPCODE_026_AUTOIGNITION_T:
      ReDim out_List_TechCodes(1 To 6)
      out_List_TechCodes(1) = TECHCODE_ANY_000u_USER_INPUT
      out_List_TechCodes(2) = TECHCODE_ANY_991d_DB911
      out_List_TechCodes(3) = TECHCODE_ANY_992d_DB801
      out_List_TechCodes(4) = TECHCODE_026_040d_MTU_FIREEXP_DATA
      out_List_TechCodes(5) = TECHCODE_026_041d_MTU_LOG_METHOD
      out_List_TechCodes(6) = TECHCODE_026_042d_MTU_LINEAR_METHOD
    Case PROPCODE_027_COMBUSTION_HEAT:
      ReDim out_List_TechCodes(1 To 3)
      out_List_TechCodes(1) = TECHCODE_ANY_000u_USER_INPUT
      out_List_TechCodes(2) = TECHCODE_ANY_991d_DB911
      out_List_TechCodes(3) = TECHCODE_ANY_992d_DB801
    '
    '////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '///////  PROPERTY SHEET "OXYGEN DEMAND": DEFAULT TECHNIQUES  ///////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '
    Case PROPCODE_028_CARBON_THOD:
      ReDim out_List_TechCodes(1 To 2)
      out_List_TechCodes(1) = TECHCODE_ANY_000u_USER_INPUT
      out_List_TechCodes(2) = TECHCODE_ANY_991d_DB911
    Case PROPCODE_029_COMBINED_THOD:
      ReDim out_List_TechCodes(1 To 2)
      out_List_TechCodes(1) = TECHCODE_ANY_000u_USER_INPUT
      out_List_TechCodes(2) = TECHCODE_ANY_991d_DB911
    Case PROPCODE_030_COD:
      ReDim out_List_TechCodes(1 To 3)
      out_List_TechCodes(1) = TECHCODE_ANY_000u_USER_INPUT
      out_List_TechCodes(2) = TECHCODE_ANY_991d_DB911
      out_List_TechCodes(3) = TECHCODE_030_043e_MTU_DIPPR
    Case PROPCODE_031_BCOD:
      ReDim out_List_TechCodes(1 To 2)
      out_List_TechCodes(1) = TECHCODE_ANY_000u_USER_INPUT
      out_List_TechCodes(2) = TECHCODE_ANY_991d_DB911
    '
    '////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '///////  PROPERTY SHEET "AQUATIC TOXICITY 1": DEFAULT TECHNIQUES  //////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '
    Case PROPCODE_041_FMINNOW_48H_EC50:
      ReDim out_List_TechCodes(1 To 2)
      out_List_TechCodes(1) = TECHCODE_ANY_000u_USER_INPUT
      out_List_TechCodes(2) = TECHCODE_ANY_991d_DB911
    Case PROPCODE_042_FMINNOW_96H_EC50:
      ReDim out_List_TechCodes(1 To 2)
      out_List_TechCodes(1) = TECHCODE_ANY_000u_USER_INPUT
      out_List_TechCodes(2) = TECHCODE_ANY_991d_DB911
    Case PROPCODE_043_FMINNOW_24H_LC50:
      ReDim out_List_TechCodes(1 To 2)
      out_List_TechCodes(1) = TECHCODE_ANY_000u_USER_INPUT
      out_List_TechCodes(2) = TECHCODE_ANY_991d_DB911
    Case PROPCODE_044_FMINNOW_48H_LC50:
      ReDim out_List_TechCodes(1 To 2)
      out_List_TechCodes(1) = TECHCODE_ANY_000u_USER_INPUT
      out_List_TechCodes(2) = TECHCODE_ANY_991d_DB911
    Case PROPCODE_045_FMINNOW_96H_LC50:
      ReDim out_List_TechCodes(1 To 2)
      out_List_TechCodes(1) = TECHCODE_ANY_000u_USER_INPUT
      out_List_TechCodes(2) = TECHCODE_ANY_991d_DB911
    Case PROPCODE_046_SALMONIDAE_24H_LC50:
      ReDim out_List_TechCodes(1 To 2)
      out_List_TechCodes(1) = TECHCODE_ANY_000u_USER_INPUT
      out_List_TechCodes(2) = TECHCODE_ANY_991d_DB911
    Case PROPCODE_047_SALMONIDAE_48H_LC50:
      ReDim out_List_TechCodes(1 To 2)
      out_List_TechCodes(1) = TECHCODE_ANY_000u_USER_INPUT
      out_List_TechCodes(2) = TECHCODE_ANY_991d_DB911
    Case PROPCODE_048_SALMONIDAE_96H_LC50:
      ReDim out_List_TechCodes(1 To 2)
      out_List_TechCodes(1) = TECHCODE_ANY_000u_USER_INPUT
      out_List_TechCodes(2) = TECHCODE_ANY_991d_DB911
    '
    '////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '///////  PROPERTY SHEET "AQUATIC TOXICITY 2": DEFAULT TECHNIQUES  //////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////
    Case PROPCODE_049_DMAGNA_24H_EC50:
      ReDim out_List_TechCodes(1 To 2)
      out_List_TechCodes(1) = TECHCODE_ANY_000u_USER_INPUT
      out_List_TechCodes(2) = TECHCODE_ANY_991d_DB911
    Case PROPCODE_050_DMAGNA_48H_EC50:
      ReDim out_List_TechCodes(1 To 2)
      out_List_TechCodes(1) = TECHCODE_ANY_000u_USER_INPUT
      out_List_TechCodes(2) = TECHCODE_ANY_991d_DB911
    Case PROPCODE_051_DMAGNA_24H_LC50:
      ReDim out_List_TechCodes(1 To 2)
      out_List_TechCodes(1) = TECHCODE_ANY_000u_USER_INPUT
      out_List_TechCodes(2) = TECHCODE_ANY_991d_DB911
    Case PROPCODE_052_DMAGNA_48H_LC50:
      ReDim out_List_TechCodes(1 To 2)
      out_List_TechCodes(1) = TECHCODE_ANY_000u_USER_INPUT
      out_List_TechCodes(2) = TECHCODE_ANY_991d_DB911
    Case PROPCODE_053_MYSID_96H_LC50:
      ReDim out_List_TechCodes(1 To 2)
      out_List_TechCodes(1) = TECHCODE_ANY_000u_USER_INPUT
      out_List_TechCodes(2) = TECHCODE_ANY_991d_DB911
    Case PROPCODE_054_ALTERNATE_SPECIES:
      ReDim out_List_TechCodes(1 To 2)
      out_List_TechCodes(1) = TECHCODE_ANY_000u_USER_INPUT
      out_List_TechCodes(2) = TECHCODE_ANY_991d_DB911
  End Select
  '
  ' EXIT OUT OF HERE.
  '
exit_normally_ThisFunc:
  Get_Complete_List_of_TechCodes = True
  Exit Function
exit_err_ThisFunc:
  Get_Complete_List_of_TechCodes = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("Get_Complete_List_of_TechCodes")
  Resume exit_err_ThisFunc
End Function




