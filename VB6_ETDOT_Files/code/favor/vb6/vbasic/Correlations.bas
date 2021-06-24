Attribute VB_Name = "Correlations"
Option Explicit




Const Correlations_declarations_end = True


'C------ CALCULATE WATER DENSITY (kg/m^3) VIA MTU CORRELATION.
'INPUTS:
'    - TEMPERATURE as Input_Temp, degK.
'RETURNS:
'    - WATER DENSITY, g/cm^3.
Function Corr_WaterDensity( _
    Input_Temp) As Double
Dim TA As Double
Dim RetVal As Double
Dim A1 As Double
Dim A2 As Double
Dim A3 As Double
Dim A4 As Double
Dim A5 As Double
Dim XAVG As Double
Dim FAVG As Double
  A1 = -1.4176800403
  A2 = 8.976651524
  A3 = -12.275501969
  A4 = 7.4584410413
  A5 = -1.738491605
  XAVG = 324.65
  FAVG = 0.98396
  'NOTE: Input_Temp IS IN UNITS OF DEGK.
  TA = Input_Temp / XAVG
  RetVal = (A1 + A2 * TA + _
           A3 * TA ^ 2# + _
           A4 * TA ^ 3# + _
           A5 * TA ^ 4#) * FAVG * 1000#
  '(NOTE 1000# FACTOR IS TO CONVERT g/cm^3 TO kg/m^3.)
  Corr_WaterDensity = RetVal
End Function
'C------ CALCULATE WATER VISCOSITY (g/cm-s) VIA YAWS CORRELATION.
'INPUTS:
'    - TEMPERATURE as Input_Temp, degK.
'RETURNS:
'    - WATER VISCOSITY, kg/m-s
'      (THIS WAS INCORRECTLY LABELLED AS g/cm-s BEFORE 1/6/99).
Function Corr_WaterViscosity( _
    Input_Temp As Double) As Double
Dim TB As Double
Dim RetVal As Double
  'NOTE: Input_Temp IS IN UNITS OF DEGK.
  TB = Input_Temp
  RetVal = Exp((-24.71) + (4209#) / TB + _
           (0.04527) * TB - (0.00003376) * TB ^ 2#)
  RetVal = RetVal / 1000#
  Corr_WaterViscosity = RetVal
End Function
'C------ CALCULATE AIR VISCOSITY (g/cm-s) VIA HOKANSON CODE.
'INPUTS:
'    - TEMPERATURE as Input_Temp, degK.
'RETURNS:
'    - AIR VISCOSITY, kg/m-s
'      (THIS WAS INCORRECTLY LABELLED AS g/cm-s BEFORE 1/6/99).
Function Corr_AirViscosity( _
    Input_Temp As Double) As Double
  'NOTE: Input_Temp IS IN UNITS OF DEGK.
  Corr_AirViscosity = _
      0.00000017 * (Input_Temp ^ 0.818)
End Function
'C------ CALCULATE AIR DENSITY (kg/m^3) VIA HOKANSON CODE.
'INPUTS:
'    - TEMPERATURE as Input_Temp, degK.
'    - PRESSURE as Input_Pres, atm.
'RETURNS:
'    - AIR DENSITY, kg/m^3.
Function Corr_AirDensity( _
    Input_Temp As Double, _
    Input_Pres As Double) As Double
Dim MWAVG As Double
Dim r As Double
Dim ThisVal As Double
  'NOTE: Input_Temp IS IN UNITS OF DEGK.
  'NOTE: Input_Pres IS IN UNITS OF ATM.
  MWAVG = 28.95   'UNITS OF G/GMOL.
  r = 0.08205     'UNITS OF (L-ATM)/(GMOL-K).
  '
  ' ON THE NEXT LINE, THE UNITS ARE gmol/L =
  ' (atm/((L-atm)/(gmol-K))/(K)) =
  ' (atm*gmol-K)/(L-atm-K) = gmol/L (CHECKS).
  ThisVal = (Input_Pres / r / Input_Temp)
  '
  ' ON THE NEXT LINE, THE UNITS ARE g/L = (gmol/L)*(g/gmol) (CHECKS).
  ThisVal = ThisVal * MWAVG
  '
  ' ON THE NEXT LINE, THE UNITS ARE kg/m^3 = g/L (CHECKS).
  ThisVal = ThisVal * 1#
  '
  ' RETURN THE VALUE.
  Corr_AirDensity = ThisVal
  'Corr_AirDensity = _
      (1# / 1000#) * ((MWAVG) * (Input_Pres)) / ((r) * (Input_Temp))
End Function
'
' PURPOSE:
'     - CALCULATE WATER VAPOR PRESSURE (Pa).
' SOURCE:
'     - Temperature range of 0-40 degC from Handbook
'       of Chemistry and Physics (72nd Edition).
' INPUTS:
'     - TEMPERATURE as Input_Temp, degK.
' RETURNS:
'     - WATER VAPOR PRESSURE, Pa.
'
Function Corr_WaterVaporPressure( _
    Input_Temp As Double)
Dim TempInDegC As Double
Dim P_kPa As Double
Dim P_Pa As Double
  TempInDegC = Input_Temp - 273.15
  P_kPa = _
      0.776 - 0.0047 * (TempInDegC) + _
      0.0041 * (TempInDegC) ^ 2
  P_Pa = P_kPa * 1000#
  '
  ' RETURN THE VALUE.
  Corr_WaterVaporPressure = P_Pa
End Function
'
' PURPOSE:
'     - CALCULATE HENRY'S CONSTANT FOR OXYGEN (dimensionless).
' SOURCE:
'     - Montgomery, J.M. Water Treatment Principles and Design,
'       John Wiley and Sons, NY, 1985.
' INPUTS:
'     - TEMPERATURE as Input_Temp, degK.
' RETURNS:
'     - HENRY'S CONSTANT FOR OXYGEN, dimensionless.
'
Function Corr_OxygenHenrysConst( _
    Input_Temp As Double)
Dim T_K As Double
Dim RetVal As Double
  T_K = Input_Temp
  RetVal = _
      (10# ^ (7.11 + (-1450) / (1.987 * T_K))) / _
      (0.08205 * T_K * 55.5)
  Corr_OxygenHenrysConst = RetVal
End Function
'
' PURPOSE:
'     - CALCULATE OXYGEN SATURATION CONCENTRATION (mg/L).
' SOURCE:
'     - UNKNOWN (DR. MIHELCIC, 1996).
' INPUTS:
'     - TEMPERATURE as Input_Temp, degK.
'     - PRESSURE as Input_Pres, atm.
' RETURNS:
'     - OXYGEN SATURATION CONCENTRATION, mg/L.
'
'
' DERIVATION:
' -----------
'   q'_O2 = (0.21)*(P)/(R*T)
'   q_O2 = (C'_O2)*(32 g/gmol)*(1000 mg/g)
'   C_O2 = (q_O2)/(H)
'   C_O2 = (0.21)*(P)/(R*T)*(32 g/gmol)*(1000 mg/g)/(H)
'   Note that R = 0.08205 L-atm/(gmol-K)
'
' UNITS IN DERIVATION:
' --------------------
'   VARIABLE     UNITS
'   ========     =============
'   q'_O2        gmol/L
'   q_O2         mg/L
'   C_O2         mg/L
'   P            atm
'   R            L-atm/(gmol-K)
'   T            K
'   n/V          gmol/L
'   H            (dimensionless)
'
'
Function Corr_OxygenSatConc( _
    Input_Temp As Double, _
    Input_Pres As Double)
Dim T_K As Double
Dim RetVal As Double
Dim HenrysO2 As Double
  T_K = Input_Temp
  HenrysO2 = Corr_OxygenHenrysConst(T_K)
''''''''''''''''  RetVal = _
''''''''''''''''      (0.21 * HenrysO2 * 32# * 1000#) / _
''''''''''''''''      (0.08205 * T_K)
''''  RetVal = _
''''      (0.21 * 32# * 1000#) / _
''''      (0.08205 * T_K * HenrysO2)
  RetVal = _
      (0.21 * Input_Pres * 32# * 1000#) / _
      (0.08205 * T_K * HenrysO2)
  Corr_OxygenSatConc = RetVal
End Function
'
' PURPOSE:
'     - CALCULATE OXYGEN DIFFUSIVITY IN WATER (cm^2/s).
' SOURCE:
'     - Holmen, K. and P. Liss, "Models for Air-Water Gas
'       Transfer: An Experimental Investigation," Tellus,
'       36B:92-100, (1984).
' INPUTS:
'     - TEMPERATURE as Input_Temp, degK.
' RETURNS:
'     - OXYGEN DIFFUSIVITY IN WATER, cm^2/s.
'
Function Corr_OxygenAqDiffusivity( _
    Input_Temp As Double)
Dim T_K As Double
Dim RetVal As Double
  T_K = Input_Temp
  RetVal = _
      (10# ^ (3.15 + (-8.31) / (T_K))) * (10# ^ -5)
  Corr_OxygenAqDiffusivity = RetVal
End Function


Sub Corr_SetWaterAndAirAndOxygen(Pp As TYPE_PlantDiagram)
Dim ThisVal As Double
Dim Temp_in_C As Double
Dim Temp_in_K As Double
Dim Pres_in_kPa As Double
Dim Pres_in_atm As Double
  With Pp.ChemicalData
    '
    '///////////////////////////////////////////////////////////////////////////////////////////////////
    ' SET CURRENT TEMPERATURE AND PRESSURE.
    '///////////////////////////////////////////////////////////////////////////////////////////////////
    '
    ' NOTE, THIS CODE CONVERTS TEMPERATURE FROM degC TO degK.
    Temp_in_C = .DataSources(1).Val_UserInput  '.env_Temperature
    Temp_in_K = Temp_in_C + 273.15
    ' NOTE, THIS CODE CONVERTS PRESSURE FROM kPa TO atm.
    Pres_in_kPa = .DataSources(0).Val_UserInput  '.env_Pressure
    Pres_in_atm = Pres_in_kPa * 1000# / 101325#
    '
    '///////////////////////////////////////////////////////////////////////////////////////////////////
    ' SET WATER PROPERTIES.
    '///////////////////////////////////////////////////////////////////////////////////////////////////
    '
    ' SET .H2O_Density TO CORRECT CORRELATION SETTING.
    '
    ThisVal = Corr_WaterDensity(Temp_in_K)
    .DataSources(13).Val_Corr = ThisVal
    '
    ' SET .H2O_Viscosity TO CORRECT CORRELATION SETTING.
    '
    ThisVal = Corr_WaterViscosity(Temp_in_K)
    .DataSources(14).Val_Corr = ThisVal
    '
    ' SET .H2O_VaporPressure TO CORRECT CORRELATION SETTING.
    '
    ThisVal = Corr_WaterVaporPressure(Temp_in_K)
    '
    ' THE NEXT LINE CONVERTS Pa TO kPa.
    ThisVal = ThisVal / 1000#
    .DataSources(15).Val_Corr = ThisVal
    '
    '///////////////////////////////////////////////////////////////////////////////////////////////////
    ' SET AIR PROPERTIES.
    '///////////////////////////////////////////////////////////////////////////////////////////////////
    '
    ' SET .AIR_Density TO CORRECT CORRELATION SETTING.
    '
    ThisVal = Corr_AirDensity( _
        Temp_in_K, _
        Pres_in_atm)
    .DataSources(17).Val_Corr = ThisVal
    '
    ' SET .AIR_Viscosity TO CORRECT CORRELATION SETTING.
    '
    ThisVal = Corr_AirViscosity( _
        Temp_in_K)
    .DataSources(18).Val_Corr = ThisVal
    '
    '///////////////////////////////////////////////////////////////////////////////////////////////////
    ' SET OXYGEN PROPERTIES.
    '///////////////////////////////////////////////////////////////////////////////////////////////////
    '
    ' SET .O2_HenrysConstant TO CORRECT CORRELATION SETTING.
    '
    ThisVal = Corr_OxygenHenrysConst(Temp_in_K)
    .DataSources(11).Val_Corr = ThisVal
    '
    ' SET .O2_SaturationConc TO CORRECT CORRELATION SETTING.
    '
    ThisVal = Corr_OxygenSatConc(Temp_in_K, Pres_in_atm)
    .DataSources(10).Val_Corr = ThisVal
    '
    ' SET .O2_Diffusivity TO CORRECT CORRELATION SETTING.
    '
    ThisVal = Corr_OxygenAqDiffusivity(Temp_in_K)
    .DataSources(12).Val_Corr = ThisVal
  End With
End Sub


