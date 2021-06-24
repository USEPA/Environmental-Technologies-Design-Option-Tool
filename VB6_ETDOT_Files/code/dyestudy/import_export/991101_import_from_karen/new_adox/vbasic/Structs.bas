Attribute VB_Name = "Structs"
Option Explicit

Global frmPrint_DO_INPUTS As Boolean
Global frmPrint_DO_OUTPUTS As Boolean
Global frmPrint_DO_PLOTS As Boolean

Global Const USE_FONTNAME = "arial"
Global Const USE_FONTSIZE = 8
Global Const USE_FORMAT_CURRENCYSTANDARD = "$#,##0_);[Red]($#,##0)"
Global Const USE_FORMAT_CURRENCYDIGITSPAST2 = "$#,##0.00_);[Red]($#,##0.00)"

Global Const Latest_DataVersion_Major = 1
Global Const Latest_DataVersion_Minor = 0
Global Current_Filename As String

Global Const IDREACT_CMBR = 0
Global Const IDREACT_CMFR = 1

Global Const IDCARBN_TIC = 0
Global Const IDCARBN_ALKALINITY = 1

Global Const IDUVI_EINSTEINS_L_S = 0
Global Const IDUVI_WATTS = 1
Global Const IDUVI_EFFICIENCY = 2

Type TankConcLabels_Type
  Label1 As String        'LINE 1 OF OUTPUT.
  Label2 As String        'LINE 2 OF OUTPUT (UNITS).
End Type
'Global TankConcLabels() As TankConcLabels_Type
'Global TankConcs() As Double       'GMOL/L OR MG/L
    'TankConcs(i,j,k): i=CHEMICAL #, j=TANK #, k=ROW #.
'Global Tank_Times() As Double      'MINUTES


Type Fortran_CompName_Type
  idx As Integer
  name As String
End Type

Type Fortran_Comp_Type
  comname As String
  concini As Double
  val As Double
  mw As Double
End Type
Type Fortran_IrrRxn_Type
  compa As Fortran_CompName_Type
  compb As Fortran_CompName_Type
  compc As Fortran_CompName_Type
  compd As Fortran_CompName_Type
  xk As Double
End Type
Type Fortran_RevRxn_Type
  compe As Fortran_CompName_Type
  compf As Fortran_CompName_Type
  xke As Double
End Type
Type Fortran_PhotRxn_Type
  compg As Fortran_CompName_Type
  comph As Fortran_CompName_Type
  stocphot As Double
  extcoef() As Double
  quatyd() As Double
End Type
Type Fortran_Comp2_Type
  ncarbn As Double
  nsubstt As Double
End Type
Type Wavelength_Type
  lwave As Double           'wavelength, nanometers
  uvi As Double             'UV light intensity, units described by iduvi
End Type
'''Type DyeStudy_Type
'''  time As Double
'''  concentration As Double
'''End Type
Type TargetCompound_Type
  comname As String           'compound name
  
  'PROPERTIES OF THIS COMPOUND.
  concini As Double           'initial/influent concentration, gmol/L
      'note: concini units are mg/L for NOM.
  val As Double               'valence, dim'less (N)
  mw As Double                'molecular weight, g/gmol
  ncarbn As Integer           'number of carbon atoms per molecule (N)
  nsubstt As Integer          'number of hydrogen substituted atoms per molecule (e.g. Cl, Br, etc.) (N)
  xk As Double                'second order rate constant for irreversible OH* reaction
  
  'PROPERTIES OF THE DEPROTONATED COMPOUND (NOT APPLICABLE TO NOM).
  dep_comname As String       'deprotonated compound name
  'dep_concini as double       'forced to zero
  dep_val As Double           'deprotonated compound valence, dim'less
  dep_mw As Double            'deprotonated compound molecular weight, g/gmol
  dep_xk As Double            'second order rate constant for irreversible OH* reaction
  dep_xke As Double           'equilibrium constant for reversible deprotonation reaction
  
  'ADDITIONAL REACTIONS.
  xk_co3XM As Double          'reaction with CO3*-
  xk_hpo4XM As Double         'reaction with HPO4*-
  xk_o2XM As Double           'reaction with O2*-
  xk_ho2X As Double           'reaction with HO2*
  
  'NOTES:
  'N = NOT APPLICABLE TO NOM.
End Type

Type Project_Type
  Filename As String
  dirty As Integer          'has any data changed?
  
  'REACTOR PROPERTIES.
  idreact As Integer        'reactor type (see IDREACT_*)
  volume As Double          'reactor volume, liters
  unitsofdisplay(1 To 5) As String
  tau As Double             'hydraulic retention time, min (F)
  
  
  num_tanks As Integer      'number of tanks if idreact = IDREACT_CMFR
  
  'NUMERICAL SIMULATION PARAMETERS.
  ssize As Double           'simulation time step, sec
  ttotal As Double          'total simulation time, min (B)
  opsize As Double          'time interval for output data, min
                            'note: this is assumed equal to ssize!
  xntimes As Double         'number of hydraulic retention times to simulation, dim'less (F)

  'WATER QUALITY PROPERTIES.
  ph0 As Double             'initial/influent value of pH
  phosph As Double          'initial/influent value of total inorganic phosphate ion concentration, gmol/L
  idcarbn As Integer        'how total inorganic carbon is input (see IDCARBN_*)
  alk As Double             'initial/influent alkalinity, mg/L as CaCO3 (used only if idcarbn = IDCARBN_ALKALINITY)
  ticarbn As Double         'initial/influent total inorganic carbon conc., gmol/L (used only if idcarbn = IDCARBN_TIC)
  inf_h2o2 As Double        'initial/influent h2o2 concentration, gmol/L
  
  'TARGET COMPOUNDS.
  'NOTE: TargetCompounds(1) IS THE NOM COMPOUND.
  TargetCompounds_Count As Integer
  TargetCompounds() As TargetCompound_Type
  
  'PHOTOCHEMICAL PARAMETERS.
  iduvi As Integer          'units of .uvi: 0=light intensity (Einsteins/L-s), 1=light intensity (watts), 2=efficiency (dimensionless)
  Wavelength_Count As Integer
  Wavelengths() As Wavelength_Type
  extcoef() As Double       'extinction coefficients for each wavelength
  quatyd() As Double        'quantum yields for each wavelength
  extcoef_h2o2() As Double  'for h2o2, extinction coefficients for each wavelength
  quatyd_h2o2() As Double   'for h2o2, quantum yields for each wavelength
  uvpathl As Double         'optical path length of UV light, cm
  ''''lamp_eff As Double        'lamp efficiency, percentage (0-100%)
  lamp_power As Double      'used only if IDUVI=2
  lamp_name As String       'lamp name
  
'''  'DYE STUDY PARAMETERS.
'''  dyestudy_output As String
'''  dyestudy_count As Integer
'''  DyeStudy() As DyeStudy_Type
'''  dyestudy_calcdate As String
    
  'TEMPORARY VARIABLES; NOT SAVED OR LOADED.
  ntarget As Integer        'number of target compounds
  nmultiacid As Integer     'number of multiprotic acids
  nphot As Integer          'number of photolysis reactions
  nirrev As Integer         'number of irreversible reactions
  ncomp As Integer          'number of compounds
  nwvlen As Integer         'number of wavelengths
  NUM_REV As Integer        'number of reversible reactions
  Fortran_Comp() As Fortran_Comp_Type
  Fortran_IrrRxn() As Fortran_IrrRxn_Type
  Fortran_RevRxn() As Fortran_RevRxn_Type
  Fortran_PhotRxn() As Fortran_PhotRxn_Type
  Fortran_Comp2() As Fortran_Comp2_Type
  
  'NOTES:
  'F = FOR CMFR ONLY.
  'B = FOR CMBR ONLY.
End Type

'GLOBAL VARIABLES -- PROJECT RELATED.
Global NowProj As Project_Type



