Attribute VB_Name = "Calc_Mod_FofT"
Option Explicit






Const Calc_Mod_FofT_decl_end = True


Function Ln(in_X As Double) As Double
  Ln = Log(in_X)
End Function
Function sinh(in_X As Double) As Double
  sinh = (Exp(in_X) - Exp(-in_X)) / 2#
End Function
Function cosh(in_X As Double) As Double
  cosh = (Exp(in_X) + Exp(-in_X)) / 2#
End Function


Function Calc_FofT_Equation( _
    in_idx_Chem As Integer, _
    TechDat As TechniqueData_Type) _
    As Boolean
On Error GoTo err_ThisFunc
Dim A As Double
Dim B As Double
Dim C As Double
Dim D As Double
Dim E As Double
Dim T As Double
Dim This_FofT_EqForm As Integer
Dim CalcFofT As Double
Dim Tc As Double
Dim Tr As Double
Dim Attempted_to_use_Tr As Boolean
Dim Tr_Is_Unavailable As Boolean
  T = NowProj.Op_T
  With TechDat
    A = .FofT_Coeffs(1)
    B = .FofT_Coeffs(2)
    C = .FofT_Coeffs(3)
    D = .FofT_Coeffs(4)
    E = .FofT_Coeffs(5)
    If (T < .FofT_Minimum_T) Or (T > .FofT_Maximum_T) Then
      .IsAvail = False
      .Error_Code = "Operating temperature specified by user (" & _
          Format_Numerical_Value(T) & "K) is outside the " & _
          "acceptable correlation range of " & _
          Format_Numerical_Value(.FofT_Minimum_T) & "K to " & _
          Format_Numerical_Value(.FofT_Maximum_T) & "K."
      .Value = 0#
      .IsTagged = False
      .ReferenceText = ""
      GoTo exit_err_ThisFunc
    End If
    This_FofT_EqForm = .FofT_EqForm
  End With
  Tr_Is_Unavailable = False
  If (False = TechValue_Get( _
      in_idx_Chem, _
      PROPCODE_013_CRITICAL_T, _
      Tc)) Then
    Tr_Is_Unavailable = True
    Tr = 1#
  Else
    Tr = T / Tc
  End If
  ''''Tr = 1#           ' TEMPORARY VALUE SETTING!
  Attempted_to_use_Tr = False
  Select Case This_FofT_EqForm
    Case 100:
      CalcFofT = A + B * T + C * T ^ 2# + D * T ^ 3# + E * T ^ 4#
    Case 101:
      CalcFofT = Exp(A + (B / T) + (C * Ln(T)) + D * T ^ E)
    Case 102:
      CalcFofT = A * T ^ B / (1# + C / T + D / T ^ 2#)
    Case 105:
      CalcFofT = A / (B ^ (1# + (1# - (T / C)) ^ D))    ' LD
    Case 106:
      CalcFofT = A * (1# - Tr) ^ (B + C * Tr + D * Tr ^ 2# + E * Tr ^ 3#)
      Attempted_to_use_Tr = True
    Case 107:
      CalcFofT = A + B * ((C / T) / sinh(C / T)) ^ 2# + D * ((E / T) / cosh(E / T)) ^ 2#
    Case 114:
      CalcFofT = (A ^ 2# / (1# - Tr)) + B - (2# * A * C * (1# - Tr)) - (A * D * (1# - Tr) ^ 2#) _
                  - ((C ^ 2# * (1# - Tr) ^ 3#) / 3#) _
                  - ((C * D * (1# - Tr) ^ 4#) / 2#) _
                  - ((D ^ 2# * (1# - Tr) ^ 5#) / 5#)
      Attempted_to_use_Tr = True
    Case 115:
      CalcFofT = 2.718282 ^ (A + (B / T) + (C * Ln(T)) + (D * T ^ 2#) + (E / T ^ 2#))
    Case 116:
      CalcFofT = A + (B * (1# - Tr) ^ 0.35) + (C * (1# - Tr) ^ (2# / 3#)) + (D * (1# - Tr)) + (E * (1# - Tr) ^ (4# / 3#))
      Attempted_to_use_Tr = True
    Case 200:
      CalcFofT = A + B * T + C * T ^ 2# * Ln(T) + D * T ^ 2.5 + E * T ^ 3#
    Case 201:
      CalcFofT = A + B * T ^ 2# * Ln(T) + C * T ^ 2.5 + D * T ^ 3#
    Case 202:
      T = T - 273.15
      CalcFofT = (Exp(A + B / (T + C))) * 133.33
  End Select
  With TechDat
    .Value = CalcFofT
  End With
  If (Attempted_to_use_Tr = True) And (Tr_Is_Unavailable = True) Then
    With TechDat
      .IsAvail = False
      .Error_Code = "The reduced temperature, Tr=T/Tc, could not be calculated " & _
          "due to the unavailability of the critical temperature, Tc."
      .Value = 0#
      .IsTagged = False
      ''''.ReferenceText = ""
    End With
  End If
exit_normally_ThisFunc:
  Calc_FofT_Equation = True
  Exit Function
exit_err_ThisFunc:
  Calc_FofT_Equation = False
  Exit Function
err_ThisFunc:
  ''''Call Show_Trapped_Error("Calc_FofT_Equation")
  With TechDat
    .IsAvail = False
    .Error_Code = Get_Trapped_Error_String( _
        "Calc_FofT_Equation")
    .Value = 0#
    .IsTagged = False
    .ReferenceText = ""
  End With
  Resume exit_err_ThisFunc
End Function

