Attribute VB_Name = "Calc_Mod_Techniques_Const"
Option Explicit

'
' WARNING, WARNING, WARNING:
' ==========================
'
' 1.) ANY TIME A GIVEN PROPERTY CODE OR TECHNIQUE CODE IS REFERRED
'   TO IN ANY VISUAL BASIC (OR OTHER) CODE, IT SHOULD USE THE
'   CONSTANTS DISPLAYED BELOW.  IT SHOULD _NOT_ USE THE NUMERIC
'   VALUES IN THE REFERENCE!
'
' 2.) ANY TIME A NEW PROPERTY IS ADDED/DELETED/MODIFIED, MAKE SURE
'   TO UPDATE THE FOLLOWING SUBROUTINES:
'   ---- Given_PropCode_Get_Name()
'   ---- Given_PropCode_Get_UnitType_and_UnitBase()
'   ---- Given_PropCode_Get_Is_FofT()
'   ---- Project_UserHierarchy_SetDefaults()
'   ---- Get_Complete_List_of_PropCodes()
'
' 3.) ANY TIME A NEW TECHNIQUE IS ADDED/DELETED/MODIFIED, MAKE SURE
'   TO UPDATE THE FOLLOWING SUBROUTINES:
'   ---- Given_TechCode_Get_Name()
'   ---- Given_TechCode_Get_TechCategory()
'   ---- Project_UserHierarchy_SetDefaults()
'   ---- Get_Complete_List_of_TechCodes()
'




'
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
'///////  GENERAL TECHNIQUES  ///////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
'
Global Const TECHCODE_ANY_000u_USER_INPUT = 0
Global Const TECHCODE_ANY_991d_DB911 = 991
Global Const TECHCODE_ANY_992d_DB801 = 992

'
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
'///////  PROPERTY SHEET "GENERAL 1": DEFAULT TECHNIQUES  ///////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
'
Global Const PROPCODE_000_MOLEC_WEIGHT = 0
Global Const TECHCODE_000_002e_UNIFAC = 2

Global Const PROPCODE_001_LIQDENS_298K = 1
Global Const TECHCODE_001_003e_BHIRUDS_1978 = 3
Global Const TECHCODE_001_004e_RACKETT_1978 = 4

Global Const PROPCODE_002_LIQDENS_FOFT = 2

Global Const PROPCODE_003_MELTING_POINT = 3
Global Const TECHCODE_003_005e_TAFT_STAREK_1930 = 5
Global Const TECHCODE_003_006e_LORENZ_HERZ_1922 = 6

Global Const PROPCODE_004_NBP = 4

Global Const PROPCODE_005_VP_298K = 5

Global Const PROPCODE_006_VP_FOFT = 6
Global Const TECHCODE_006_007d_ANTOINELIKE_EXPRESSION = 7

Global Const PROPCODE_007_HEAT_FORMATION = 7

'
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
'///////  PROPERTY SHEET "GENERAL 2": DEFAULT TECHNIQUES  ///////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
'
Global Const PROPCODE_008_LIQUID_HEAT_CAPACITY_FOFT = 8

Global Const PROPCODE_009_VAPOR_HEAT_CAPACITY_FOFT = 9

Global Const PROPCODE_010_HEAT_OF_VAPORIZATION_298K = 10
Global Const TECHCODE_010_008e_WATSON = 8

Global Const PROPCODE_011_HEAT_OF_VAPORIZATION_NBP = 11
Global Const TECHCODE_011_009e_KLEIN_1949 = 9
Global Const TECHCODE_011_010e_CHEN_PITZER_1965 = 10

Global Const PROPCODE_012_HEAT_OF_VAPORIZATION_FOFT = 12

Global Const PROPCODE_013_CRITICAL_T = 13

Global Const PROPCODE_014_CRITICAL_P = 14

Global Const PROPCODE_038_CRITICAL_V = 38

'
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
'///////  PROPERTY SHEET "TRANSPORT": DEFAULT TECHNIQUES  ///////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
'
Global Const PROPCODE_015_DIFFUSIVITY_H2O = 15
Global Const TECHCODE_015_011e_HAYDUK_MINHAS_1982 = 11
Global Const TECHCODE_015_012e_HAYDUK_LAUDIE_1974 = 12
Global Const TECHCODE_015_013e_WILKE_CHANG = 13

Global Const PROPCODE_016_DIFFUSIVITY_AIR = 16
Global Const TECHCODE_016_014e_WILKE_LEE_MOD = 14

Global Const PROPCODE_017_SURFACE_TENSION_298K = 17
Global Const TECHCODE_017_015e_BROCK_BIRD_1983 = 15

Global Const PROPCODE_018_SURFACE_TENSION_FOFT = 18

Global Const PROPCODE_019_VAPOR_VISCOSITY_FOFT = 19

Global Const PROPCODE_020_LIQUID_VISCOSITY_FOFT = 20

Global Const PROPCODE_021_LIQUID_THERMAL_CONDUC_FOFT = 21

Global Const PROPCODE_022_VAPOR_THERMAL_CONDUC_FOFT = 22

'
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
'///////  PROPERTY SHEET "PARTITIONING/EQUILIBRIUM": DEFAULT TECHNIQUES  ////////////////
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
'
Global Const PROPCODE_034_AC_CHEM_IN_H2O = 34
Global Const TECHCODE_034_016e_UNIFAC = 16
Global Const TECHCODE_034_017e_HANSCH_1968 = 17

Global Const PROPCODE_032_AC_H2O_IN_CHEM = 32
Global Const TECHCODE_032_018e_UNIFAC = 18

Global Const PROPCODE_033_HENRY_CONSTANT = 33
          '
          ' MORE TECHNIQUES TO COME LATER !!!
          '

Global Const PROPCODE_039_SOL_LIMIT_CHEM_IN_H2O = 39
Global Const TECHCODE_039_020d_YAWS = 20
Global Const TECHCODE_039_019e_UNIFAC = 19
Global Const TECHCODE_039_021e_YALKOWSKY_1990 = 21

Global Const PROPCODE_035_LOG_KOW = 35
Global Const TECHCODE_035_022e_KENAGA_GORING_1978 = 22

Global Const PROPCODE_036_LOG_KOC = 36
Global Const TECHCODE_036_023e_BAKER_1994 = 23

Global Const PROPCODE_037_BIOCONC_FACTOR = 37
Global Const TECHCODE_037_024e_KOBAYSHI_1981 = 24
Global Const TECHCODE_037_025e_KENAGA_GORING_1980 = 25

'
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
'///////  PROPERTY SHEET "FIRE AND EXPLOSION": DEFAULT TECHNIQUES  //////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
'
Global Const PROPCODE_023_UF_LIMIT = 23
Global Const TECHCODE_023_026d_MTU_FIREEXP_DATA = 26
Global Const TECHCODE_023_027d_MTU_GROUP_CONTRIB = 27
Global Const TECHCODE_023_028d_MTU_COMBUSTION_RXN = 28
Global Const TECHCODE_023_029d_PENN_GROUP_CONTRIB = 29

Global Const PROPCODE_024_LF_LIMIT = 24
Global Const TECHCODE_024_030d_MTU_FIREEXP_DATA = 30
Global Const TECHCODE_024_031d_MTU_GROUP_CONTRIB = 31
Global Const TECHCODE_024_032d_PENN_GROUP_CONTRIB = 32
Global Const TECHCODE_024_033d_MTU_COMBUSTION_RXN = 33
Global Const TECHCODE_024_034d_MTU_FLASHPOINT_METH = 34

Global Const PROPCODE_025_FLASH_POINT = 25
Global Const TECHCODE_025_035d_MTU_FIREEXP_DATA = 35
Global Const TECHCODE_025_036d_LFL_DATA = 36
Global Const TECHCODE_025_037d_MTU_LFL_GROUP_CONTRIB = 37
Global Const TECHCODE_025_038d_PENN_GROUP_CONTRIB = 38
Global Const TECHCODE_025_039d_MTU_LFL_COMBUSTION_RXN = 39

Global Const PROPCODE_026_AUTOIGNITION_T = 26
Global Const TECHCODE_026_040d_MTU_FIREEXP_DATA = 40
Global Const TECHCODE_026_041d_MTU_LOG_METHOD = 41
Global Const TECHCODE_026_042d_MTU_LINEAR_METHOD = 42

Global Const PROPCODE_027_COMBUSTION_HEAT = 27

'
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
'///////  PROPERTY SHEET "OXYGEN DEMAND": DEFAULT TECHNIQUES  ///////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
'
Global Const PROPCODE_028_CARBON_THOD = 28

Global Const PROPCODE_029_COMBINED_THOD = 29

Global Const PROPCODE_030_COD = 30
Global Const TECHCODE_030_043e_MTU_DIPPR = 43

Global Const PROPCODE_031_BCOD = 31

'
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
'///////  PROPERTY SHEET "AQUATIC TOXICITY 1": DEFAULT TECHNIQUES  //////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
'
Global Const PROPCODE_041_FMINNOW_48H_EC50 = 41

Global Const PROPCODE_042_FMINNOW_96H_EC50 = 42

Global Const PROPCODE_043_FMINNOW_24H_LC50 = 43

Global Const PROPCODE_044_FMINNOW_48H_LC50 = 44

Global Const PROPCODE_045_FMINNOW_96H_LC50 = 45

Global Const PROPCODE_046_SALMONIDAE_24H_LC50 = 46

Global Const PROPCODE_047_SALMONIDAE_48H_LC50 = 47

Global Const PROPCODE_048_SALMONIDAE_96H_LC50 = 48

'
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
'///////  PROPERTY SHEET "AQUATIC TOXICITY 2": DEFAULT TECHNIQUES  //////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
'
Global Const PROPCODE_049_DMAGNA_24H_EC50 = 49

Global Const PROPCODE_050_DMAGNA_48H_EC50 = 50

Global Const PROPCODE_051_DMAGNA_24H_LC50 = 51

Global Const PROPCODE_052_DMAGNA_48H_LC50 = 52

Global Const PROPCODE_053_MYSID_96H_LC50 = 53

Global Const PROPCODE_054_ALTERNATE_SPECIES = 54


'
' LIST OF TECHNIQUE CATEGORY CONSTANTS.
'
Global Const TECHCATEGORY_USER = 1
Global Const TECHCATEGORY_DATA = 2
Global Const TECHCATEGORY_ESTIMATE = 3
'
' LIST OF CONSTANTS FOR TABS ON frmTechniques.
'
Global Const TECHNIQUE_TAB_01a_LIST = 10
Global Const TECHNIQUE_TAB_02a_DIPPR801 = 20
Global Const TECHNIQUE_TAB_02b_DIPPR911 = 21
Global Const TECHNIQUE_TAB_03a_NOTE = 30
'
' MISCELLANEOUS.
'
Global Const TECH_ERRORCODE_NEVER_INITED = "Value was never initialized!"







Const Calc_Mod_Techniques_Const_decl_end = True



