Attribute VB_Name = "Structs"
Option Explicit

Global Const Latest_DataVersion_Major = 1
Global Const Latest_DataVersion_Minor = 0
Global Current_Filename As String





Type PropertyOrder_Type
  Property_Code As Long          'Uses the PROPCODE_* constants
  Technique_Code() As Long       'Uses the TECHCODE_* constants
  '
  ' Very important note: The Customize command must be designed
  ' so that special handling occurs if the user creates more than
  ' one instance of the same property.  For this case, the
  ' technique order must be identical for both properties !!
  ' Otherwise, the calculation code (especially the subroutine
  ' Recalculate_OneProperty) is incorrect.
  '
End Type
Global Const PROPERTYSHEETNAME_BASIC_CHEMICAL_INFO = "Basic Chemical Info"
Global Const PROPERTYSHEETNAME_CHEMICAL_NOTE = "Chemical Note"
Type PropertySheetOrder_Type
  Name As String
  PropertyOrder() As PropertyOrder_Type
End Type
Type UserHierarchy_Type
  PropertySheetOrder() As PropertySheetOrder_Type
End Type
Global Const MAX_PROPERTYSHEETS = 100


''''Global Const FOFT_EQFORM_0101_something_or_another = 101
''''Global Const FOFT_EQFORM_0102_something_or_another = 102
''''Global Const FOFT_EQFORM_0103_something_or_another = 103
Type TechniqueData_Type
  '
  ' Important note: The value actually reported by the program
  ' on the main window is the first technique (ordered by
  ' NowProj.UserHierarchy) that has .IsAvail=true.
  '
  Technique_Code As Long        'Uses the TECHCODE_* constants
  IsAvail As Boolean            'Whether data is available
  Error_Code As String
      ' Error_Code could be anything ranging from "Not in database" to
      ' "The temperature is outside the valid range" to
      ' "A run-time error #134 occurred: `Whatever this error message
      ' can be translated as.`"  It's important to note that if a
      ' run-time error occurs during the calculation of a given technique,
      ' the code records this fact and continues on to attempt to
      ' calculate the next techinque.
  value As Double               'Calculation value, or 0 if "Not Available"
  IsTagged As Boolean           'Is technique tagged for calculation?
  ReferenceText As String       'Text for reference; entered by user for user-entered data
  Text_When_Blank As String     'Text to display when .IsAvail=True and .Value=0#
  '
  ' DIPPR RELATED VALUES.
  '
  DIPPR_REF As String           'DIPPR reference code (DIPPR801 non-T-dependent only)
  DIPPR_REL As String           'DIPPR reliability code (T-dependent only)
  DIPPR_R As Integer            'DIPPR rating code (non-T-dependent only)
  DIPPR_Value As Double
  DIPPR_Units As String
  DIPPR_Pressure As String      'DIPPR911-ONLY
  DIPPR_DescMethod As String    'DIPPR911-ONLY
  DIPPR_Comment As String       'DIPPR911-ONLY
  DIPPR_ArticleNumber As Long   'DIPPR911-ONLY
  '
  ' FUNCTION OF TEMPERATURE VALUES.
  '
  FofT_EqForm As Integer        'Equation form, see FOFT_EQFORM_* constants
  FofT_Coeffs() As Double       'Set of (5) coefficients for f(T) expression
  FofT_Units_F As String        'Units of "f"
  FofT_Units_T As String        'Units of "T"
  FofT_Minimum_T As Double      'Minimum T of correlation, K
  FofT_Maximum_T As Double      'Maximum T of correlation, K
End Type
Type PropertyData_Type
  '
  ' MISCELLANEOUS PROPERTY DATA.
  '
  User_Note As String
  '
  ' MAIN DATA SET.
  '
  UnitType As String            'This is the type of units, e.g. "concentration" or "molar_volume"
  UnitBase As String            '.Technique(*).Value stored in these units
  UnitDisplayed As String       'The window displays value in these units
  Property_Code As Long         'Uses the PROPCODE_* constants
  Is_FofT As Boolean            'Is technique f(T)?
  TechniqueData() As TechniqueData_Type
  IsAvail As Boolean            'Whether data is available
  idx_Technique_Used As Integer 'Index into TechniqueData() of technique used, or -1 if IsAvail=False
  Override_Technique_Code As Long   'Technique code used as override, or -1 if no override is present
End Type
Type UserChemical_Type
  '
  ' MISCELLANEOUS CHEMICAL DATA.
  '
  User_Note As String
  '
  ' BASIC CHEMICAL INFO.
  '
  Name As String
  CAS As String
  SMILES As String
  Formula As String
  Family As String
  Source As String
  '
  ' CALCULATED RESULTS.
  '
  PropertyData() As PropertyData_Type
End Type
Type Project_Type
''''  '
''''  ' TEMPORARY, SOON-TO-BE-DELETED DATA.
''''  '
''''  length As Double
''''  Diameter As Double
''''  Mass As Double
''''  FlowRate As Double
  '
  ' MISCELLANEOUS FILE DATA.
  '
  File_Note As String
  '
  ' HIERARCHY RELATED DATA.
  '
  UserHierarchy As UserHierarchy_Type
  '
  ' MAIN DATA SET.
  '
  Op_T As Double      'K
  Op_P As Double      'Pa
  Op_T_UnitDisplayed As String
  Op_P_UnitDisplayed As String
  UserChemicals() As UserChemical_Type
End Type
Global Const MAX_USERCHEMICALS = 1000

Global NowProj As Project_Type


'
'for each chemical:
'- user note for this chemical
'- for each property:
'  - property code
'  - unit of display
'  - user note for this property
'  - for each hard-coded technique:
'    - technique code
'    - value
'    - IsAvail: boolean stating whether value is available
'    - error_code:
'      - Could be anything ranging from "Not in database" to
'        "The temperature is outside the valid range" to
'        "A run-time error #134 occurred: `Whatever this error message
'        can be translated as.`"
'    - Important note: The value actually reported by the program
'      on the main window is the first technique (ordered by
'      NowProj.UserHierarchy) that has .IsAvail=true.
'


Global Const NUMFORMAT_3SIGFIG = 1
Global Const NUMFORMAT_4SIGFIG = 2
Global Const NUMFORMAT_5SIGFIG = 3
Global Const NUMFORMAT_6SIGFIG = 4
Global Const NUMFORMAT_EXP3 = 5
Global Const NUMFORMAT_EXP4 = 6
Global Const NUMFORMAT_EXP5 = 7
Global Const NUMFORMAT_3PASTDEC = 8
Global Const NUMFORMAT_4PASTDEC = 9
Global Const NUMFORMAT_5PASTDEC = 10
Type PrefEnvironment_Type
  '
  ' NUMERICAL DISPLAY FORMAT.
  '
  NumFormat_Greater1000 As Integer
  NumFormat_Less0_001 As Integer
  NumFormat_Other As Integer
  '
  ' MISCELLANEOUS.
  '
  FontSize_Lists As Integer
End Type
Global PrefEnvironment As PrefEnvironment_Type





