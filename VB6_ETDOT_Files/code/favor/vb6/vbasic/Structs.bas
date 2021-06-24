Attribute VB_Name = "Structs"
Option Explicit

Global frmPrint_DO_INPUTS As Boolean
Global frmPrint_DO_OUTPUTS As Boolean
Global frmPrint_DO_PLOTS As Boolean

Global Const Latest_DataVersion_Major = 1
Global Const Latest_DataVersion_Minor = 0
Global Current_Filename As String

Global Const WEIR_MODEL_TYPE_NAPPE = 0
Global Const WEIR_MODEL_TYPE_POOL = 1
Type TYPE_Weir
  '
  ' VALUES.
  ModelingMechanism As Integer
  Width As Double
  WaterLevelDiff As Double
  GasFlow As Double
  '
  ' UNITS.
  UnitsOfDisplay(0 To 2) As String
End Type
Global Const GRITCHAMBER_MAX_CHAMBER = 9
Type TYPE_GritChamber
  '
  ' VALUES.
  IsCovered As Boolean
  Count As Integer
  VentilationRate As Double
  Depth As Double
  Volume As Double
  GasFlow As Double
  SOTR As Double
  '
  ' UNITS.
  UnitsOfDisplay(1 To 5) As String
End Type
Global Const PRIMCLARIF_SORPTION_REMOVAL_DOBBS = 0
Global Const PRIMCLARIF_SORPTION_REMOVAL_MATTER_MULLER = 1
Global Const PRIMCLARIF_VOLATILIZATION_REMOVAL_MACKAY_YEUN = 0
Global Const PRIMCLARIF_VOLATILIZATION_REMOVAL_KLA = 1
Global Const PRIMCLARIF_MAX_CLARIFIERS = 9
Global Const SECONDCLARIF_MAX_CLARIFIERS = 9
Type TYPE_Clarifier
  '
  ' VALUES.
  IsCovered As Boolean
  Count As Integer
  SorptionRemovalMethod As Integer
  VolatilizationRemovalMechanism As Integer
  VentilationRate As Double
  Depth As Double
  Volume As Double
  WastageFlow As Double           'PRIMARY CLARIFIER ONLY.
  PercentageRemoval As Double     'PRIMARY CLARIFIER ONLY, %.
  EffluentSolidsConc As Double    'SECONDARY CLARIFIER ONLY, mg/L.
  '
  ' UNITS.
  UnitsOfDisplay(1 To 5) As String
End Type
Type TYPE_CSTRModeling
  Count As Integer
  UseStepFeed As Boolean
  UniformFeed As Boolean
  Feed(0 To 8) As Double
  UniformVolume As Boolean
  Volume(0 To 8) As Double
  UniformGasFlow As Boolean
  GasFlow(0 To 8) As Double
  UniformBioMass As Boolean
  BioMass(0 To 8) As Double
End Type
Type TYPE_BioTreatmentModeling
  '
  ' VALUES.
  MaxGrowthRate As Double
  HalfVelocityConst As Double
  BacterialDecay As Double
  YieldCoeff As Double
  BOD5Conc As Double
  '
  ' UNITS.
  UnitsOfDisplay(0 To 4) As String
End Type
Global Const AERATIONBASIN_MAX_BASIN = 9
Global Const AERATIONBASIN_MAX_CSTR = 9
Global Const BASIN_MODEL_TYPE_SURFACE = 0
Global Const BASIN_MODEL_TYPE_DIFFBUBBLE = 1
Type TYPE_AerationBasin
  '
  ' VALUES.
  IsCovered As Boolean
  Count As Integer
  ModelingMechanism As Integer
  AutoCalcBioMass As Boolean
  VentilationRate As Double
  Depth As Double
  WastageFlow As Double
  RecycleFlow As Double
  SolidsConcInRecycle As Double
  SOTR As Double
  Volume As Double
  GasFlow As Double
  BioMass As Double
  CSTR As TYPE_CSTRModeling
  BioTreat As TYPE_BioTreatmentModeling
  '
  ' UNITS.
  UnitsOfDisplay(1 To 5) As String
End Type
Global Const DATASOURCETYPE_USERINPUT = 1
Global Const DATASOURCETYPE_STEPP = 2
Global Const DATASOURCETYPE_CORR = 3
Type TYPE_DataSource
  ' NOTE, -1E20 OR LOWER INDICATES "UNAVAILABLE".
  SourceType As Integer
  Val_UserInput As Double   'DIRECT USER ENTRY VALUE.
  Val_StEPP As Double       'StEPP IMPORT VALUE.
  Val_Corr As Double        'INTERNAL CORRELATION VALUE.
End Type
Type TYPE_PhysicoChemicalData
  '
  ' VALUES.
  env_Pressure As Double
  env_Temperature As Double
  env_WindVelocity As Double
  ContaminantName As String       '* 80
  InfluentConc As Double
  BiodegredationRate As Double
  LogKow As Double
  VOC_HenrysConstant As Double
  VOC_MolecularWeight As Double
  VOC_DiffusivityInH2O As Double
  VOC_DiffusivityInGas As Double
  O2_SaturationConc As Double
  O2_HenrysConstant As Double
  O2_Diffusivity As Double
  H2O_Density As Double
  H2O_Viscosity As Double
  H2O_VaporPressure As Double
  H2O_Alpha As Double
  AIR_Density As Double
  AIR_Viscosity As Double
  '
  ' VARIOUS SOURCES OF DATA USED TO POPULATE "VALUES".
  ' NOTE, -1E20 OR LOWER INDICATES "UNAVAILABLE".
  DataSources(0 To 18) As TYPE_DataSource
  '
  ' NOTE, O2_CInfinity IS APPARENTLY UNUSED.
  O2_CInfinity As Double
  '
  ' UNITS.
  UnitsOfDisplay(0 To 18) As String
End Type
Type TYPE_PlantDiagram
  ''''IsModified As Boolean
  en_InfluentWeir As Boolean
  en_GritChamber As Boolean
  en_PrimaryWeir As Boolean
  en_SecondaryWeir As Boolean
  Flow As Double
  SolidsConc As Double
  '
  '
  InfluentWeir As TYPE_Weir
  GritChamber As TYPE_GritChamber
  PrimaryClarifier As TYPE_Clarifier
  PrimaryWeir As TYPE_Weir
  AerationBasin As TYPE_AerationBasin
  SecondaryClarifier As TYPE_Clarifier
  SecondaryWeir As TYPE_Weir
  ChemicalData As TYPE_PhysicoChemicalData
End Type


Type TYPE_OutputMechanismRecord
  EffluentConc As Double
  Stripping As Double
  Volatilization As Double
  SolidWaste As Double
  LiquidWaste As Double
  Biodegredation As Double
  pr_Stripping As Double
  pr_Volatilization As Double
  pr_SolidWaste As Double
  pr_LiquidWaste As Double
  pr_Biodegredation As Double
End Type
Type TYPE_OutputRecord
  IsDisplayed As Boolean
  TotalAmount As TYPE_OutputMechanismRecord
  pr_TotalRemoved As Double
  TotalInfluent As Double
  TotalEffluent As Double
  '
  '
  InfluentWeir As TYPE_OutputMechanismRecord
  GritChamber As TYPE_OutputMechanismRecord
  PrimaryClarifier As TYPE_OutputMechanismRecord
  PrimaryWeir As TYPE_OutputMechanismRecord
  AerationBasin As TYPE_OutputMechanismRecord
  SecondaryClarifier As TYPE_OutputMechanismRecord
  SecondaryWeir As TYPE_OutputMechanismRecord
End Type


Global Const UnitType___ENGLISH = 1
Global Const UnitType___SI = 2
Type Project_Type
  'length As Double
  'Diameter As Double
  'Mass As Double
  'FlowRate As Double
  '
  ' MAIN INPUTS.
  '
  Plant As TYPE_PlantDiagram
  '
  ' MAIN OUTPUTS.
  '
  OutputRec As TYPE_OutputRecord
  '
  ' MISCELLANEOUS.
  '
  UnitType As Integer
  KP1_OUT As Double                   'L/mg
  XVALS_OUT(1 To 7) As Double         'mg/L
End Type
Global Calculated_OK As Boolean
Global NowProj As Project_Type



