Attribute VB_Name = "UnitCnv0"
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
    
Global Const TIME_MIN = 0
Global Const TIME_S = 1
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

Global Const CONCENTRATION_MG_per_L = 0
Global Const CONCENTRATION_UG_per_L = 1
Global Const CONCENTRATION_G_per_L = 2
Global Const CONCENTRATION_MEQ_per_L = 3
Global Const CONCENTRATION_EQ_per_L = 4
Global Const CONCENTRATION_MMOL_per_L = 5
Global Const CONCENTRATION_UMOL_per_L = 6
Global Const CONCENTRATION_GMOL_per_L = 7

Global Const DIFFUSIVITY_CM2_per_S = 0
Global Const DIFFUSIVITY_CM2_per_MIN = 1
Global Const DIFFUSIVITY_M2_per_S = 2
Global Const DIFFUSIVITY_M2_per_MIN = 3
Global Const DIFFUSIVITY_M2_per_HR = 4
Global Const DIFFUSIVITY_M2_per_D = 5
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

Function ConcentrationConversionFactor(UnitsToConvertTo As Integer, Valence As Double, MolecularWeight As Double) As Double
   'Valence in eq/mol; Molecular Weight in mg/mmol

   'This function will convert from standard concentration units (mg/L) to
   'the units specified by the user.

   Select Case UnitsToConvertTo
      Case CONCENTRATION_UG_per_L   'Convert to ug/L
         ConcentrationConversionFactor = 1000#
      Case CONCENTRATION_G_per_L   'Convert to g/L
         ConcentrationConversionFactor = 1# / 1000#
      Case CONCENTRATION_MEQ_per_L   'Convert to meq/L
         ConcentrationConversionFactor = Valence / MolecularWeight
      Case CONCENTRATION_EQ_per_L   'Convert to eq/L
         ConcentrationConversionFactor = Valence / MolecularWeight / 1000
      Case CONCENTRATION_MMOL_per_L   'Convert to mmol/L
         ConcentrationConversionFactor = 1# / MolecularWeight
      Case CONCENTRATION_UMOL_per_L   'Convert to umol/L
         ConcentrationConversionFactor = 1000# / MolecularWeight
      Case CONCENTRATION_GMOL_per_L   'Convert to gmol/L
         ConcentrationConversionFactor = 1 / 1000# / MolecularWeight

   End Select

End Function

Function DensityConversionFactor(UnitsToConvertTo As Integer) As Double
   'This function will convert from standard density units (g/ml) to
   'the units specified by the user.

   Select Case UnitsToConvertTo
      Case APPARENT_DENSITY_KG_per_M3   'Convert to kg/m3
         DensityConversionFactor = 1000#
      Case APPARENT_DENSITY_LB_per_FT3   'Convert to lb/ft3
         DensityConversionFactor = (2.20462 / 1000#) / (35.3145 / 1000#)
      Case APPARENT_DENSITY_LB_per_GAL   'Convert to lb/gal
         DensityConversionFactor = (2.20462 / 1000#) / (7.4805 / 28317)
   End Select

End Function

Function DiffusivityConversionFactor(UnitsToConvertTo As Integer) As Double
   'This function will convert from standard diffusivity units (cm2/s) to
   'the units specified by the user.

   Select Case UnitsToConvertTo
      Case DIFFUSIVITY_CM2_per_MIN   'Convert to cm2/min
         DiffusivityConversionFactor = 60#
      Case DIFFUSIVITY_M2_per_S   'Convert to m2/s
         DiffusivityConversionFactor = 1 / (100# ^ 2)
      Case DIFFUSIVITY_M2_per_MIN   'Convert to m2/min
         DiffusivityConversionFactor = 60# / (100# ^ 2)
      Case DIFFUSIVITY_M2_per_HR   'Convert to m2/hr
         DiffusivityConversionFactor = (60# * 60#) / (100# ^ 2)
      Case DIFFUSIVITY_M2_per_D   'Convert to m2/d
         DiffusivityConversionFactor = (60# * 60# * 24#) / (100# ^ 2)
      Case DIFFUSIVITY_FT2_per_S   'Convert to ft2/s
         DiffusivityConversionFactor = (3.2808 ^ 2 / 100# ^ 2)
      Case DIFFUSIVITY_FT2_per_MIN   'Convert to ft2/min
         DiffusivityConversionFactor = 60# * (3.2808 ^ 2 / 100# ^ 2)
      Case DIFFUSIVITY_FT2_per_HR   'Convert to ft2/hr
         DiffusivityConversionFactor = 60# * 60# * (3.2808 ^ 2 / 100# ^ 2)
      Case DIFFUSIVITY_FT2_per_D   'Convert to ft2/d
         DiffusivityConversionFactor = 60# * 60# * 24# * (3.2808 ^ 2 / 100# ^ 2)
   End Select

End Function

Function FlowConversionFactor(UnitsToConvertTo As Integer) As Double
   'This function will convert from standard flow units (m3/s) to
   'the units specified by the user.

   Select Case UnitsToConvertTo
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

Function LengthConversionFactor(UnitsToConvertTo As Integer) As Double
   'This function will convert from standard length units (m) to
   'the units specified by the user.

   Select Case UnitsToConvertTo
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
      Case MASS_G   'Convert to g
         MassConversionFactor = 1000#
      Case MASS_LB   'Convert to lb
         MassConversionFactor = 2.20462
   End Select

End Function

Function PressureConversionFactor(UnitsToConvertTo As Integer) As Double

   'This function will convert from standard pressure units (Pa) to
   'the units specified by the user.

   Select Case UnitsToConvertTo
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

Function ResinCapacityConversionFactor(UnitsToConvertTo As Integer) As Double
   'This function will convert from standard resin capacity units (meq/g) to
   'the units specified by the user.

   Dim BedVolume As Double, ColumnArea As Double

   ColumnArea = PI * (Bed.Diameter * 100#) ^ 2 / 4
   Select Case UnitsToConvertTo
      Case RESIN_CAPACITY_MEQ_per_MLbed   'Convert to meq/ml bed
         ResinCapacityConversionFactor = Bed.Weight * 1000# / ColumnArea / (Bed.Length * 100#)
      Case RESIN_CAPACITY_MEQ_per_MLresin   'Convert to meq/ml resin
         ResinCapacityConversionFactor = Resin.ApparentDensity
   End Select

End Function

Function ReverseTemperatureConversion(UnitsToConvertFrom As Integer, Temperature_NonKelvin As Double) As Double
   'This function will convert from non-standard temperature units (C, R, or F) to
   'standard temperature units (K)

   Select Case UnitsToConvertFrom
      Case TEMPERATURE_C   'Convert from Deg. C
         ReverseTemperatureConversion = Temperature_NonKelvin + 273.15
      Case TEMPERATURE_R   'Convert from Deg. R
         ReverseTemperatureConversion = Temperature_NonKelvin / 1.8
      Case TEMPERATURE_F   'Convert from Deg. F
         ReverseTemperatureConversion = ((Temperature_NonKelvin - 32#) / 1.8) + 273.15
   End Select

End Function

Function TemperatureConversion(UnitsToConvertTo As Integer, Temperature_Kelvin As Double) As Double
   'This function will convert from standard temperature units (K) to
   'the units specified by the user.

   Select Case UnitsToConvertTo
      Case TEMPERATURE_C   'Convert to Deg. C
         TemperatureConversion = Temperature_Kelvin - 273.15
      Case TEMPERATURE_R   'Convert to Deg. R
         TemperatureConversion = 1.8 * Temperature_Kelvin
      Case TEMPERATURE_F   'Convert to Deg. F
         TemperatureConversion = 1.8 * (Temperature_Kelvin - 273.15) + 32#
   End Select

End Function

Function TimeConversionFactor(UnitsToConvertTo As Integer) As Double
   'This function will convert from standard time units (min) to
   'the units specified by the user.

   Select Case UnitsToConvertTo
      Case TIME_S   'Convert to s
         TimeConversionFactor = 60#
      Case TIME_HR   'Convert to hr
         TimeConversionFactor = 1# / 60#
      Case TIME_D   'Convert to d
         TimeConversionFactor = 1# / 1440#

   End Select

End Function

Function VelocityConversionFactor(UnitsToConvertTo As Integer) As Double

   'This function will convert from standard velocity units (cm/s) to
   'the units specified by the user.

   Select Case UnitsToConvertTo
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

