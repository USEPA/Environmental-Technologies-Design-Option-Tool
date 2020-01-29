Attribute VB_Name = "Correlations_Old"
Option Explicit







Const Correlations_Old_declarations_end = True


'
' PURPOSE:
'     - CALCULATE WATER VAPOR PRESSURE (Pa).
' INPUTS:
'     - TEMPERATURE as Input_Temp, degK.
' RETURNS:
'     - WATER VAPOR PRESSURE, Pa.
'
Function Corr_WaterVaporPressure_OLD( _
    Input_Temp As Double)
Dim A As Double
Dim B As Double
Dim C As Double
Dim D As Double
Dim E As Double
Dim ThisVal As Double
Dim T As Double
  A = 58.773
  B = -5900#
  C = -5.659
  D = 0.005
  E = 1.029
  T = Input_Temp
  '
  ' NOTE: Log(X) IS THE NATURAL LOGARITHM OF X.
  ThisVal = Exp(A + B / T + C * (Log(T)) + D * (T ^ E))
  '
  ' RETURN THE VALUE.
  Corr_WaterVaporPressure_OLD = ThisVal
End Function




