Attribute VB_Name = "UnitConvMod"
Option Explicit
Global Const PRESSURE_PA = 0
Global Const PRESSURE_KPA = 1
Global Const PRESSURE_BARS = 2
Global Const PRESSURE_ATM = 3
Global Const PRESSURE_PSI = 4
Global Const PRESSURE_MMHG = 5
Global Const PRESSURE_MH2O = 6
Global Const PRESSURE_FTH2O = 7
Global Const PRESSURE_INHG = 8

Global Const TEMPERATURE_K = 0
Global Const TEMPERATURE_C = 1
Global Const TEMPERATURE_R = 2
Global Const TEMPERATURE_F = 3

Global Const LENGTH_M = 0
Global Const LENGTH_CM = 1
Global Const LENGTH_FT = 2
Global Const LENGTH_IN = 3

Global Const MASS_KG = 0
Global Const MASS_G = 1
Global Const MASS_LB = 2

Global Const FLOW_M3_per_S = 0
Global Const FLOW_M3_per_D = 1
Global Const FLOW_CM3_per_S = 2
Global Const FLOW_ML_per_MIN = 3
Global Const FLOW_FT3_per_S = 4
Global Const FLOW__FT3_per_D = 5
Global Const FLOW_GPM = 6
Global Const FLOW_GPD = 7
Global Const FLOW_MGD = 8
    
Global Const TIME_S = 0
Global Const TIME_MIN = 1
Global Const TIME_HR = 2
Global Const TIME_D = 3

Global Const APPARENT_DENSITY_G_per_ML = 0
Global Const APPARENT_DENSITY_KG_per_M3 = 1
Global Const APPARENT_DENSITY_LB_per_FT3 = 2
Global Const APPARENT_DENSITY_LB_per_GAL = 3

Global Const RESIN_CAPACITY_MEQ_per_G = 0
Global Const RESIN_CAPACITY_MEQ_per_MLbed = 1
Global Const RESIN_CAPACITY_MEQ_per_MLresin = 2

Global Const MOLECULAR_WEIGHT_MG_per_MMOL = 0
Global Const MOLECULAR_WEIGHT_UG_per_UMOL = 1
Global Const MOLECULAR_WEIGHT_G_per_GMOL = 2
Global Const MOLECULAR_WEIGHT_KG_per_KMOL = 3

Global Const CONCENTRATION_UG_per_L = 0
Global Const CONCENTRATION_MG_per_L = 1
Global Const CONCENTRATION_G_per_L = 2
'Global Const CONCENTRATION_MEQ_per_L = 3
'Global Const CONCENTRATION_EQ_per_L = 4
'Global Const CONCENTRATION_MMOL_per_L = 5
'Global Const CONCENTRATION_UMOL_per_L = 6
'Global Const CONCENTRATION_GMOL_per_L = 7

Global Const DIFFUSIVITY_M2_per_S = 0
Global Const DIFFUSIVITY_M2_per_MIN = 1
Global Const DIFFUSIVITY_M2_per_HR = 2
Global Const DIFFUSIVITY_M2_per_D = 3
Global Const DIFFUSIVITY_CM2_per_S = 4
Global Const DIFFUSIVITY_CM2_per_MIN = 5
Global Const DIFFUSIVITY_FT2_per_S = 6
Global Const DIFFUSIVITY_FT2_per_MIN = 7
Global Const DIFFUSIVITY_FT2_per_HR = 8
Global Const DIFFUSIVITY_FT2_per_D = 9

Global Const VELOCITY_CM_per_S = 0
Global Const VELOCITY_CM_per_MIN = 1
Global Const VELOCITY_M_per_S = 2
Global Const VELOCITY_M_per_MIN = 3
Global Const VELOCITY_M_per_HR = 4
Global Const VELOCITY_M_per_D = 5
Global Const VELOCITY_FT_per_S = 6
Global Const VELOCITY_FT_per_MIN = 7
Global Const VELOCITY_FT_per_HR = 8
Global Const VELOCITY_FT_per_D = 9

Global Const MOLAR_VOLUME_M3_per_KMOL = 0
Global Const MOLAR_VOLUME_M3_per_GMOL = 1
Global Const MOLAR_VOLUME_L_per_GMOL = 2
Global Const MOLAR_VOLUME_ML_per_GMOL = 3

Global Const K_FREUNDLICH_MG_G_L = 0
Global Const K_FREUNDLICH_MMOL_G_L = 1
Global Const K_FREUNDLICH_UG_G_L = 2
Global Const K_FREUNDLICH_UMOL_G_L = 3

Global Const PRESSUREPERLENGTH_PA_per_M = 0
Global Const PRESSUREPERLENGTH_PSI_per_FT = 1
Global Const PRESSUREPERLENGTH_ATM_per_FT = 2

Global Const MASSLOADINGRATE_KG_M2_S = 0
Global Const MASSLOADINGRATE_G_M2_D = 1
Global Const MASSLOADINGRATE_LBM_FT2_S = 2

Global Const AREA_M2 = 0
Global Const AREA_CM2 = 1
Global Const AREA_FT2 = 2

Global Const VOLUME_M3 = 0
Global Const VOLUME_CM3 = 1
Global Const VOLUME_LITER = 2
Global Const VOLUME_FT3 = 3
Global Const VOLUME_GAL = 4

Global Const INVERSETIME_S = 0
Global Const INVERSETIME_MIN = 1
Global Const INVERSETIME_HR = 2
Global Const INVERSETIME_D = 3

Global Const POWER_KW = 0
Global Const POWER_W = 1
Global Const POWER_HP = 2
Global Const POWER_FTLB_per_S = 3

Global Const POWERPERVOLUME_W_per_M3 = 0
Global Const POWERPERVOLUME_KW_per_M3 = 1
Global Const POWERPERVOLUME_HP_per_FT3 = 2
Global Const POWERPERVOLUME_FTLB_per_S_FT3 = 3


'--------------------------------------------------------
'---  Unit Types:
'--------------------------------------------------------
Global Const UNITS_LENGTH = 1
Global Const UNITS_TIME = 2
Global Const UNITS_MASS = 3
Global Const UNITS_PRESSURE = 4
Global Const UNITS_TEMPERATURE = 5
Global Const UNITS_FLOW = 6
Global Const UNITS_PRESSUREPERLENGTH = 7
Global Const UNITS_MASSLOADINGRATE = 8
Global Const UNITS_INVERSETIME = 9
Global Const UNITS_AREA = 10
Global Const UNITS_VOLUME = 11
Global Const UNITS_DIFFUSIVITY = 12
Global Const UNITS_CONCENTRATION = 13
Global Const UNITS_POWER = 14
Global Const UNITS_POWERPERVOLUME = 15
Global Const UNITS_MW = 16
Global Const UNITS_MOLARVOLUME = 17


'--------------------------------------------------------
'---  SI or English indicator
'--------------------------------------------------------
Global Const UNITSTYPE_SI = 1
Global Const UNITSTYPE_ENGLISH = 2

Function AreaConversionFactor(UnitsToConvertTo As Integer) As Double
'This function will convert from standard area units (m^2) to
'the units specified by the user.

  Select Case UnitsToConvertTo
    Case AREA_M2
      AreaConversionFactor = 1#
    Case AREA_CM2
      AreaConversionFactor = 10000#
    Case AREA_FT2
      AreaConversionFactor = 10.7639
  End Select

End Function

Function ConcentrationConversionFactor(UnitsToConvertTo As Integer) As Double
'This function will convert from standard concentration units (ug/L) to
'the units specified by the user.

   Select Case UnitsToConvertTo
      Case CONCENTRATION_UG_per_L
         ConcentrationConversionFactor = 1#
      Case CONCENTRATION_MG_per_L   'Convert to mg/L
         ConcentrationConversionFactor = 1# / 1000#
      Case CONCENTRATION_G_per_L   'Convert to g/L
         ConcentrationConversionFactor = 1# / 1000# / 1000#

   End Select

End Function

Function DensityConversionFactor(UnitsToConvertTo As Integer) As Double
   'This function will convert from standard density units (g/ml) to
   'the units specified by the user.

   Select Case UnitsToConvertTo
      Case APPARENT_DENSITY_G_per_ML
         DensityConversionFactor = 1#
      Case APPARENT_DENSITY_KG_per_M3   'Convert to kg/m3
         DensityConversionFactor = 1000#
      Case APPARENT_DENSITY_LB_per_FT3   'Convert to lb/ft3
         DensityConversionFactor = (2.20462 / 1000#) / (35.3145 / 1000#)
      Case APPARENT_DENSITY_LB_per_GAL   'Convert to lb/gal
         DensityConversionFactor = (2.20462 / 1000#) / (7.4805 / 28317)
   End Select

End Function

Function DiffusivityConversionFactor(UnitsToConvertTo As Integer) As Double
   'This function will convert from standard diffusivity units (m^2/s) to
   'the units specified by the user.

   Select Case UnitsToConvertTo
      Case DIFFUSIVITY_M2_per_S
         DiffusivityConversionFactor = 1
      Case DIFFUSIVITY_M2_per_MIN   'Convert to m2/min
         DiffusivityConversionFactor = 60#
      Case DIFFUSIVITY_M2_per_HR   'Convert to m2/hr
         DiffusivityConversionFactor = 60# * 60#
      Case DIFFUSIVITY_M2_per_D   'Convert to m2/d
         DiffusivityConversionFactor = (60# * 60# * 24#)
      Case DIFFUSIVITY_CM2_per_S  'Convert to cm2/s
         DiffusivityConversionFactor = 100# * 100#
      Case DIFFUSIVITY_CM2_per_MIN   'Convert to cm2/min
         DiffusivityConversionFactor = (100# * 100#) * (60#)
      Case DIFFUSIVITY_FT2_per_S   'Convert to ft2/s
         DiffusivityConversionFactor = 10.7639
      Case DIFFUSIVITY_FT2_per_MIN   'Convert to ft2/min
         DiffusivityConversionFactor = 10.7639 * 60#
      Case DIFFUSIVITY_FT2_per_HR   'Convert to ft2/hr
         DiffusivityConversionFactor = 10.7639 * (60# * 60#)
      Case DIFFUSIVITY_FT2_per_D   'Convert to ft2/d
         DiffusivityConversionFactor = 10.7639 * (60# * 60# * 24#)
   End Select

End Function

Function FlowConversionFactor(UnitsToConvertTo As Integer) As Double
   'This function will convert from standard flow units (m3/s) to
   'the units specified by the user.

   Select Case UnitsToConvertTo
      Case FLOW_M3_per_S
         FlowConversionFactor = 1#
      Case FLOW_M3_per_D   'Convert to m3/d
         FlowConversionFactor = 86400#
      Case FLOW_CM3_per_S   'Convert to cm3/s
         FlowConversionFactor = 100# * 100# * 100#
      Case FLOW_ML_per_MIN   'Convert to ml/min
         FlowConversionFactor = 1000000# * 60#
      Case FLOW_FT3_per_S   'Convert to ft3/s
         FlowConversionFactor = 35.3145
      Case FLOW__FT3_per_D   'Convert to ft3/d
         FlowConversionFactor = 35.3145 * 86400#
      Case FLOW_GPM   'Convert to gpm
         FlowConversionFactor = 264.17 * 60#
      Case FLOW_GPD   'Convert to gpd
         FlowConversionFactor = 264.17 * 86400#
      Case FLOW_MGD   'Convert to MGD
         FlowConversionFactor = (264.17 / 1000000#) * 86400

   End Select

End Function

Function GetConversionFactor(WhatUnits As Integer, ListIndex As Integer)

  Select Case WhatUnits
    Case UNITS_LENGTH
      GetConversionFactor = LengthConversionFactor(ListIndex)
    Case UNITS_TIME
      GetConversionFactor = TimeConversionFactor(ListIndex)
    Case UNITS_MASS
      GetConversionFactor = MassConversionFactor(ListIndex)
    Case UNITS_PRESSURE
      GetConversionFactor = PressureConversionFactor(ListIndex)
    'Case UNITS_TEMPERATURE
    '  GetConversionFactor = TemperatureConversionFactor(ListIndex)
    Case UNITS_FLOW
      GetConversionFactor = FlowConversionFactor(ListIndex)
    Case UNITS_PRESSUREPERLENGTH
      GetConversionFactor = PressurePerLengthConversionFactor(ListIndex)
    Case UNITS_MASSLOADINGRATE
      GetConversionFactor = MassLoadingRateConversionFactor(ListIndex)
    Case UNITS_INVERSETIME
      GetConversionFactor = InverseTimeConversionFactor(ListIndex)
    Case UNITS_AREA
      GetConversionFactor = AreaConversionFactor(ListIndex)
    Case UNITS_VOLUME
      GetConversionFactor = VolumeConversionFactor(ListIndex)
    Case UNITS_DIFFUSIVITY
      GetConversionFactor = DiffusivityConversionFactor(ListIndex)
    Case UNITS_CONCENTRATION
      GetConversionFactor = ConcentrationConversionFactor(ListIndex)
    Case UNITS_POWER
      GetConversionFactor = PowerConversionFactor(ListIndex)
    Case UNITS_POWERPERVOLUME
      GetConversionFactor = PowerPerVolumeConversionFactor(ListIndex)
    Case UNITS_MW
      GetConversionFactor = MolecularWeightConversionFactor(ListIndex)
    Case UNITS_MOLARVOLUME
      GetConversionFactor = MolarVolumeConversionFactor(ListIndex)
      
      
    'more to come....

  End Select

End Function

Function GetUnits(UnitList As ComboBox) As String

  GetUnits = Trim$(UnitList.List(UnitList.ListIndex))

End Function

Function InverseTimeConversionFactor(UnitsToConvertTo As Integer) As Double
   'This function will convert from standard inverse time units (1/sec) to
   'the units specified by the user.

   Select Case UnitsToConvertTo
      Case INVERSETIME_S
         InverseTimeConversionFactor = 1#
      Case INVERSETIME_MIN    'Convert to 1/min
         InverseTimeConversionFactor = 60#
      Case INVERSETIME_HR     'Convert to 1/hr
         InverseTimeConversionFactor = 60# * 60#
      Case INVERSETIME_D      'Convert to 1/d
         InverseTimeConversionFactor = 60# * 60# * 24#
   End Select

End Function

Function KFreundlichConversionFactor(UnitsToConvertTo As Integer, OneOverN As Double, MW As Double) As Double
   'This function will convert from standard Freundlich K units
   '((mg/g)*(L/mg)^(1/n)) to the units specified by the user.
Dim factor1 As Double
Dim factor2 As Double
   
   factor1 = MW ^ (OneOverN - 1)
   factor2 = 1000 ^ (1 - OneOverN)
   
   Select Case UnitsToConvertTo
      Case K_FREUNDLICH_MG_G_L
         'No conversion.
         KFreundlichConversionFactor = 1#
      Case K_FREUNDLICH_MMOL_G_L
         'Convert to (mmol/g)*(L/mmol)^(1/n)
         KFreundlichConversionFactor = factor1
      Case K_FREUNDLICH_UG_G_L
         'Convert to (ug/g)*(L/ug)^(1/n)
         KFreundlichConversionFactor = factor2
      Case K_FREUNDLICH_UMOL_G_L
         'Convert to (umol/g)*(L/umol)^(1/n)
         KFreundlichConversionFactor = factor1 * factor2
   End Select

End Function

Function LengthConversionFactor(UnitsToConvertTo As Integer) As Double
   'This function will convert from standard length units (m) to
   'the units specified by the user.

   Select Case UnitsToConvertTo
      Case LENGTH_M
         LengthConversionFactor = 1#
      Case LENGTH_CM   'Convert to cm
         LengthConversionFactor = 100#
      Case LENGTH_FT   'Convert to feet
         LengthConversionFactor = 3.2808
      Case LENGTH_IN   'Convert to inches
         LengthConversionFactor = 39.37
   End Select

End Function

Function MassConversionFactor(UnitsToConvertTo As Integer) As Double
   'This function will convert from standard mass units (kg) to
   'the units specified by the user.

   Select Case UnitsToConvertTo
      Case MASS_KG
         MassConversionFactor = 1#
      Case MASS_G   'Convert to g
         MassConversionFactor = 1000#
      Case MASS_LB   'Convert to lb
         MassConversionFactor = 2.20462
   End Select

End Function

Function MassLoadingRateConversionFactor(UnitsToConvertTo As Integer) As Double
'This function will convert from standard mass loading rate units (kg/m^2/s) to
'the units specified by the user.

  Select Case UnitsToConvertTo
    Case MASSLOADINGRATE_KG_M2_S
      MassLoadingRateConversionFactor = 1#
    Case MASSLOADINGRATE_G_M2_D
      MassLoadingRateConversionFactor = 86400# * 1000#
    Case MASSLOADINGRATE_LBM_FT2_S
      MassLoadingRateConversionFactor = 2.20462 / 10.7639
  End Select

End Function

Function MolarVolumeConversionFactor(UnitsToConvertTo As Integer) As Double
   'This function will convert from standard molar volume units (m^3/kmol) to
   'the units specified by the user.

   Select Case UnitsToConvertTo
      Case MOLAR_VOLUME_M3_per_KMOL
         MolarVolumeConversionFactor = 1#
      Case MOLAR_VOLUME_M3_per_GMOL   'Convert to m^3/gmol
         MolarVolumeConversionFactor = 1# / 1000#
      Case MOLAR_VOLUME_L_per_GMOL    'Convert to L/gmol
         MolarVolumeConversionFactor = 1#
      Case MOLAR_VOLUME_ML_per_GMOL   'Convert to mL/gmol
         MolarVolumeConversionFactor = 1000#
   End Select

End Function

Function MolecularWeightConversionFactor(UnitsToConvertTo As Integer) As Double
   'This function always returns 1#
   '(all MW's are the same in this version of UNITCONV.BAS).

   MolecularWeightConversionFactor = 1#

End Function

Sub Populate_Area_Units(dest As ComboBox, Default_Unit As Integer)
  
  dest.Clear
  dest.AddItem "m" & Chr$(178)
  dest.AddItem "cm" & Chr$(178)
  dest.AddItem "ft" & Chr$(178)
  dest.ListIndex = Default_Unit

End Sub

Sub Populate_Concentration_Units(dest As ComboBox, Default_Unit As Integer)

  dest.Clear
  dest.AddItem Chr$(181) & "g/L"
  dest.AddItem "mg/L"
  dest.AddItem "g/L"
  dest.ListIndex = Default_Unit

End Sub

Sub Populate_Density_Units(dest As ComboBox, Default_Unit As Integer)

  dest.Clear
  dest.AddItem "g/mL"
  dest.AddItem "kg/m" & Chr$(179)
  dest.AddItem "lb/ft" & Chr$(179)
  dest.AddItem "lb/gal"
  dest.ListIndex = Default_Unit

End Sub

Sub Populate_Diffusivity_Units(dest As ComboBox, Default_Unit As Integer)

  dest.Clear
  dest.AddItem "m" & Chr$(178) & "/s"
  dest.AddItem "m" & Chr$(178) & "/min"
  dest.AddItem "m" & Chr$(178) & "/hr"
  dest.AddItem "m" & Chr$(178) & "/d"
  dest.AddItem "cm" & Chr$(178) & "/s"
  dest.AddItem "cm" & Chr$(178) & "/min"
  dest.AddItem "ft" & Chr$(178) & "/s"
  dest.AddItem "ft" & Chr$(178) & "/min"
  dest.AddItem "ft" & Chr$(178) & "/hr"
  dest.AddItem "ft" & Chr$(178) & "/d"
  dest.ListIndex = Default_Unit

End Sub

Sub Populate_FlowRate_Units(dest As ComboBox, Default_Unit As Integer)
  
  dest.Clear
  dest.AddItem "m" & Chr$(179) & "/s"
  dest.AddItem "m" & Chr$(179) & "/d"
  dest.AddItem "cm" & Chr$(179) & "/s"
  dest.AddItem "mL/min"
  dest.AddItem "ft" & Chr$(179) & "/s"
  dest.AddItem "ft" & Chr$(179) & "/d"
  dest.AddItem "gpm"
  dest.AddItem "gpd"
  dest.AddItem "MGD"
  dest.ListIndex = Default_Unit

End Sub

Sub Populate_InverseTime_Units(dest As ComboBox, Default_Unit As Integer)

  dest.Clear
  dest.AddItem "1/s"
  dest.AddItem "1/min"
  dest.AddItem "1/hr"
  dest.AddItem "1/d"
  dest.ListIndex = Default_Unit

End Sub

Sub Populate_KFreundlich_Units(dest As ComboBox, Default_Unit As Integer)
'
'  dest.Clear
'  dest.AddItem "(mg/g)*(L/mg)^(1/n)"
'  dest.AddItem "(mmol/g)*(L/mmol)^(1/n)"
'  dest.AddItem "(" & Chr$(181) & "g/g)*(L/" & Chr$(181) & "g)^(1/n)"
'  dest.AddItem "(" & Chr$(181) & "mol/g)*(L/" & Chr$(181) & "mol)^(1/n)"
'  dest.ListIndex = Default_Unit
'
End Sub

Sub Populate_Length_Units(dest As ComboBox, Default_Unit As Integer)

  dest.Clear
  dest.AddItem "m"
  dest.AddItem "cm"
  dest.AddItem "ft"
  dest.AddItem "in"
  dest.ListIndex = Default_Unit

End Sub

Sub Populate_Mass_Units(dest As ComboBox, Default_Unit As Integer)

  dest.Clear
  dest.AddItem "kg"
  dest.AddItem "g"
  dest.AddItem "lb"
  dest.ListIndex = Default_Unit

End Sub

Sub Populate_MassLoadingRate_Units(dest As ComboBox, Default_Unit As Integer)

  dest.Clear
  dest.AddItem "kg/m" & Chr$(178) & "-s"
  dest.AddItem "g/m" & Chr$(178) & "-d"
  dest.AddItem "lbm/ft" & Chr$(178) & "-s"
  dest.ListIndex = Default_Unit

End Sub

Sub Populate_MolarVolume_Units(dest As ComboBox, Default_Unit As Integer)

  dest.Clear
  dest.AddItem "m" & Chr$(179) & "/kmol"
  dest.AddItem "m" & Chr$(179) & "/gmol"
  dest.AddItem "L/gmol"
  dest.AddItem "mL/gmol"
  dest.ListIndex = Default_Unit

End Sub

Sub Populate_MolecularWeight_Units(dest As ComboBox, Default_Unit As Integer)

  dest.Clear
  dest.AddItem "mg/mmol"
  dest.AddItem Chr$(181) & "g/" & Chr$(181) & "mol"
  dest.AddItem "g/gmol"
  dest.AddItem "kg/kmol"
  dest.ListIndex = Default_Unit

End Sub

Sub Populate_Power_Units(dest As ComboBox, Default_Unit As Integer)

  dest.Clear
  dest.AddItem "kW"
  dest.AddItem "W"
  dest.AddItem "hp"
  dest.AddItem "ft-lb/s"
  dest.ListIndex = Default_Unit

End Sub

Sub Populate_PowerPerVolume_Units(dest As ComboBox, Default_Unit As Integer)
  
  dest.Clear
  dest.AddItem "W/m" & Chr$(179)
  dest.AddItem "kW/m" & Chr$(179)
  dest.AddItem "hp/ft" & Chr$(179)
  dest.AddItem "ft-lb/ft" & Chr$(179) & "-s"
  dest.ListIndex = Default_Unit

End Sub

Sub Populate_Pressure_Units(dest As ComboBox, Default_Unit As Integer)

  dest.Clear
  dest.AddItem "Pa"
  dest.AddItem "kPa"
  dest.AddItem "bars"
  dest.AddItem "atm"
  dest.AddItem "psi"
  dest.AddItem "mmHg"
  dest.AddItem "mH20"
  dest.AddItem "ftH20"
  dest.AddItem "inHg"
  dest.ListIndex = Default_Unit

End Sub

Sub Populate_PressurePerLength_Units(dest As ComboBox, Default_Unit As Integer)

  dest.Clear
  dest.AddItem "Pa/m"
  dest.AddItem "psi/ft"
  dest.AddItem "atm/ft"
  dest.ListIndex = Default_Unit

End Sub

Sub Populate_Temperature_Units(dest As ComboBox, Default_Unit As Integer)

  dest.Clear
  dest.AddItem "K"
  dest.AddItem "C"
  dest.AddItem "R"
  dest.AddItem "F"
  dest.ListIndex = Default_Unit

End Sub

Sub Populate_Time_Units(dest As ComboBox, Default_Unit As Integer)

  dest.Clear
  dest.AddItem "s"
  dest.AddItem "min"
  dest.AddItem "hr"
  dest.AddItem "d"
  dest.ListIndex = Default_Unit

End Sub

Sub Populate_Volume_Units(dest As ComboBox, Default_Unit As Integer)

  dest.Clear
  dest.AddItem "m" & Chr$(179)
  dest.AddItem "cm" & Chr$(179)
  dest.AddItem "liter"
  dest.AddItem "ft" & Chr$(179)
  dest.AddItem "gal"
  dest.ListIndex = Default_Unit

End Sub

Function PowerConversionFactor(UnitsToConvertTo As Integer) As Double
'This function will convert from standard power units (kilowatts) to
'the units specified by the user.

  Select Case UnitsToConvertTo
    Case POWER_KW
      PowerConversionFactor = 1#
    Case POWER_W
      PowerConversionFactor = 1000#
    Case POWER_HP
      PowerConversionFactor = 1000# / 745.6999
    Case POWER_FTLB_per_S
      PowerConversionFactor = 1000# / 1.35582
  End Select

End Function

Function PowerPerVolumeConversionFactor(UnitsToConvertTo As Integer) As Double
'This function will convert from standard power per volume units (watts/m^3) to
'the units specified by the user.

  Select Case UnitsToConvertTo
    Case POWERPERVOLUME_W_per_M3
      PowerPerVolumeConversionFactor = 1#
    Case POWERPERVOLUME_KW_per_M3
      PowerPerVolumeConversionFactor = 1# / 1000#
    Case POWERPERVOLUME_HP_per_FT3
      PowerPerVolumeConversionFactor = 1# / 26334.14
    Case POWERPERVOLUME_FTLB_per_S_FT3
      PowerPerVolumeConversionFactor = 1# / 47.88026
  End Select

End Function

Function PressureConversionFactor(UnitsToConvertTo As Integer) As Double

   'This function will convert from standard pressure units (Pa) to
   'the units specified by the user.

   Select Case UnitsToConvertTo
      Case PRESSURE_PA
         PressureConversionFactor = 1#
      Case PRESSURE_KPA   'Convert to kPa
         PressureConversionFactor = 1# / 1000#
      Case PRESSURE_BARS   'Convert to bars
         PressureConversionFactor = 1# / 100000#
      Case PRESSURE_ATM   'Convert to atm
         PressureConversionFactor = 1# / 101325#
      Case PRESSURE_PSI   'Convert to psi
         PressureConversionFactor = 14.696 / 101325#
      Case PRESSURE_MMHG   'Convert to mm Hg
         PressureConversionFactor = 760# / 101325#
      Case PRESSURE_MH2O   'Convert to m H2O
         PressureConversionFactor = 10.333 / 101325#
      Case PRESSURE_FTH2O   'Convert to ft H2O
         PressureConversionFactor = 33.9 / 101325#
      Case PRESSURE_INHG   'Convert to in. Hg
         PressureConversionFactor = 29.921 / 101325#

   End Select

End Function

Function PressurePerLengthConversionFactor(UnitsToConvertTo As Integer) As Double
'This function will convert from standard pressure/length units (Pa/m) to
'the units specified by the user.

  Select Case UnitsToConvertTo
    Case PRESSUREPERLENGTH_PA_per_M
      PressurePerLengthConversionFactor = 1#
    Case PRESSUREPERLENGTH_PSI_per_FT
      PressurePerLengthConversionFactor = 0.3048 / 6894.76
    Case PRESSUREPERLENGTH_ATM_per_FT
      PressurePerLengthConversionFactor = 0.3048 / 101325
  End Select

End Function

Function ResinCapacityConversionFactor(UnitsToConvertTo As Integer) As Double
''   'This function will convert from standard resin capacity units (meq/g) to
'   'the units specified by the user.
'
'   Dim BedVolume As Double, ColumnArea As Double
'
'   ColumnArea = Pi * (Bed.Diameter * 100#) ^ 2 / 4
'   Select Case UnitsToConvertTo
'      Case RESIN_CAPACITY_MEQ_per_MLbed   'Convert to meq/ml bed
'         ResinCapacityConversionFactor = Bed.Weight * 1000# / ColumnArea / (Bed.Length * 100#)
'      Case RESIN_CAPACITY_MEQ_per_MLresin   'Convert to meq/ml resin
'         ResinCapacityConversionFactor = Resin.ApparentDensity
'   End Select
'
End Function

Function ReverseTemperatureConversion(UnitsToConvertFrom As Integer, Temperature_NonKelvin As Double) As Double
   'This function will convert from non-standard temperature units (C, R, or F) to
   'standard temperature units (K)

   Select Case UnitsToConvertFrom
      Case TEMPERATURE_K   'Convert from Deg. K
         ReverseTemperatureConversion = Temperature_NonKelvin
      Case TEMPERATURE_C   'Convert from Deg. C
         ReverseTemperatureConversion = Temperature_NonKelvin + 273.15
      Case TEMPERATURE_R   'Convert from Deg. R
         ReverseTemperatureConversion = Temperature_NonKelvin / 1.8
      Case TEMPERATURE_F   'Convert from Deg. F
         ReverseTemperatureConversion = (Temperature_NonKelvin + 459.67) / 1.8
   End Select

End Function

Sub Units_DoRefresh(UnitList As ComboBox)
Dim OldUnit As Integer
  OldUnit = UnitList.ListIndex
  UnitList.ListIndex = -1
  UnitList.ListIndex = OldUnit
End Sub

Sub SetUnits(UnitList As ComboBox, NewUnitStr As String)
Dim i As Integer
Dim NewUnit As Integer

  NewUnit = -1

  For i = 0 To UnitList.ListCount - 1
    If (Trim$(UnitList.List(i)) = Trim$(NewUnitStr)) Then
      NewUnit = i
      Exit For
    End If
  Next i

  If (NewUnit = -1) Then
    'Error!  Unit not found!
    NewUnit = 0
  End If

  UnitList.ListIndex = -1
  UnitList.ListIndex = NewUnit
  
End Sub

Function TemperatureConversion(UnitsToConvertTo As Integer, Temperature_Kelvin As Double) As Double
   'This function will convert from standard temperature units (K) to
   'the units specified by the user.

   Select Case UnitsToConvertTo
      Case TEMPERATURE_K
         TemperatureConversion = Temperature_Kelvin
      Case TEMPERATURE_C   'Convert to Deg. C
         TemperatureConversion = Temperature_Kelvin - 273.15
      Case TEMPERATURE_R   'Convert to Deg. R
         TemperatureConversion = 1.8 * Temperature_Kelvin
      Case TEMPERATURE_F   'Convert to Deg. F
         TemperatureConversion = 1.8 * Temperature_Kelvin - 459.67
   End Select

End Function

Function TimeConversionFactor(UnitsToConvertTo As Integer) As Double
   'This function will convert from standard time units (sec) to
   'the units specified by the user.

   Select Case UnitsToConvertTo
      Case TIME_S
         TimeConversionFactor = 1#
      Case TIME_MIN   'Convert to min
         TimeConversionFactor = 1 / 60#
      Case TIME_HR    'Convert to hr
         TimeConversionFactor = 1# / 60# / 60#
      Case TIME_D     'Convert to d
         TimeConversionFactor = 1# / 60# / 60# / 24#
   End Select

End Function

Function VelocityConversionFactor(UnitsToConvertTo As Integer) As Double

   'This function will convert from standard velocity units (cm/s) to
   'the units specified by the user.

   Select Case UnitsToConvertTo
      Case VELOCITY_CM_per_S
         VelocityConversionFactor = 1#
      Case VELOCITY_CM_per_MIN   'Convert to cm/min
         VelocityConversionFactor = 60#
      Case VELOCITY_M_per_S   'Convert to m/s
         VelocityConversionFactor = 1 / (100#)
      Case VELOCITY_M_per_MIN   'Convert to m/min
         VelocityConversionFactor = 60# / (100#)
      Case VELOCITY_M_per_HR   'Convert to m/hr
         VelocityConversionFactor = (60# * 60#) / (100#)
      Case VELOCITY_M_per_D   'Convert to m/d
         VelocityConversionFactor = (60# * 60# * 24#) / (100#)
      Case VELOCITY_FT_per_S   'Convert to ft/s
         VelocityConversionFactor = (3.2808 / 100#)
      Case VELOCITY_FT_per_MIN   'Convert to ft/min
         VelocityConversionFactor = 60# * (3.2808 / 100#)
      Case VELOCITY_FT_per_HR   'Convert to ft/hr
         VelocityConversionFactor = 60# * 60# * (3.2808 / 100#)
      Case VELOCITY_FT_per_D   'Convert to ft/d
         VelocityConversionFactor = 60# * 60# * 24# * (3.2808 / 100#)
   End Select

End Function

Function VolumeConversionFactor(UnitsToConvertTo As Integer) As Double
'This function will convert from standard volume units (m^3) to
'the units specified by the user.

  Select Case UnitsToConvertTo
    Case VOLUME_M3
      VolumeConversionFactor = 1#
    Case VOLUME_CM3
      VolumeConversionFactor = 1000# * 1000#
    Case VOLUME_LITER
      VolumeConversionFactor = 1000#
    Case VOLUME_FT3
      VolumeConversionFactor = 35.3147
    Case VOLUME_GAL
      VolumeConversionFactor = 264.172
  End Select

End Function

