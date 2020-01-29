Attribute VB_Name = "Structs2"
Option Explicit

'//////// COMMUNICATIONS WITH frmEditAdsorber: ////////////////////////////////////////////////////////
Type rec_adsorber_db_manufacturers
  UniqueID As String    'Must be an integer in string form!
  Name As String
End Type

Type rec_adsorber_db_adsorbers
  UniqueID_Manufacturer As Integer
  Phase As Integer
  PartNumber As String * 20
  InternalArea As String * 20
  MaxCapacity As String * 20
  OutsideDiameter As String * 20
  DesignPressure As String * 20
  DesignFlowRange As String * 20
  DefaultFlowRate As String * 20
  Note As String * 100
End Type

Global adsorber_db_num_manufacturers As Integer
Global adsorber_db_manufacturers() As rec_adsorber_db_manufacturers

Global adsorber_db_num_adsorbers As Integer
Global adsorber_db_adsorbers() As rec_adsorber_db_adsorbers

Type rec_frmEditAdsorber_ReturnParameters
  D As Double
  L As Double
  M As Double
  Q As Double
End Type

Global frmEditAdsorber_ReturnParameters As rec_frmEditAdsorber_ReturnParameters


'//////// COMMUNICATIONS WITH frmEditAdsorberData: /////////////////////////////////////////////////
Global frmEditAdsorberData_Record As rec_adsorber_db_adsorbers


'//////// COMMUNICATIONS WITH frmEditCarbonData: /////////////////////////////////////////////////
Type frmEditCarbonData_Record_Type
  PhaseIsLiquid As Boolean
  Name As String
  Manufacturer As String
  AppDen As Double
  ParticleRadius As Double
  ParticlePorosity As Double
  AdsType As String
  W0 As Double
  BB As Double
  PolanyiExponent As Double
End Type
Global frmEditCarbonData_Record As frmEditCarbonData_Record_Type


'//////// COMMUNICATIONS WITH frmEditIsothermData: /////////////////////////////////////////////////
Type frmEditIsothermData_Record_Type
  PhaseIsLiquid As Boolean
  Name As String
  k As Double
  OneOverN As Double
  Cmin As Double
  Cmax As Double
  pHmin As Double
  pHmax As Double
  Source As String
  CarbonName As String
  Tmin As String
  CAS As String
  Comments As String
End Type
Global frmEditIsothermData_Record As frmEditIsothermData_Record_Type



'---- frmConcentrations variables
Global frmConcentrations_cancelled As Integer
Global frmConcentrations_Times(1 To 400) As Double
Global frmConcentrations_Concs(1 To 10, 1 To 400) As Double
Global frmConcentrations_NumPoints As Integer
Global frmConcentrations_NumConcs As Integer
Global frmConcentrations_caption As String
Global frmConcentrations_TimeOrderImportant As Integer
Global frmConcentrations_Cunits As String
Global frmConcentrations_Tunits As String

'---- frmShow_Data_And_Prediction variables
Global frmCompareData_WhichSet As Integer
Global Const frmCompareData_WhichSet_PSDM = 1
Global Const frmCompareData_WhichSet_CPHSDM = 2
Global frmCompareData_caption As String






'MISCELLANEOUS.
Global FileNote As String

