Attribute VB_Name = "UnitsMod"
Option Explicit

Sub Combobox_KeyPress(KeyAscii As Integer)

  If KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys "{Tab}"
    Exit Sub
  End If

End Sub

Function FormatSciNumber(Dummy As Double) As String

  Select Case Abs(Dummy)
    Case Is < 0.1
      FormatSciNumber = Format$(Dummy, "0.000e+00")
    Case Is > 100#
      FormatSciNumber = Format$(Dummy, "0.000e+00")
    Case Else
      FormatSciNumber = Format$(Dummy, "0.000")
  End Select

End Function

Function NoUnits_LostFocus(txt As TextBox, NewVal As Double, Temp_Text As String)
Dim NoUnitBox As ComboBox

  NoUnits_LostFocus = Unitted_LostFocus(0, txt, NoUnitBox, NewVal, Temp_Text)

End Function

Sub TextBoxNumber_KeyPress(KeyAscii As Integer)

  If KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys "{Tab}"
    Exit Sub
  End If
  
  If ((KeyAscii > Asc("9") Or KeyAscii < Asc("0")) And KeyAscii <> Asc(".") And KeyAscii <> 8 And KeyAscii <> Asc("E") And (KeyAscii <> Asc("e")) And (KeyAscii <> Asc("-"))) Then
    KeyAscii = 0
    Beep
  End If

End Sub

Sub TextBoxString_KeyPress(KeyAscii As Integer)
    
  If (KeyAscii = 13) Then
    KeyAscii = 0
    SendKeys "{Tab}"
    Exit Sub
  End If

End Sub

Sub TextHandleError(IsError As Integer, txt As TextBox, Temp_Text As String)
Dim Dummy As Double
Dim i As Integer

  IsError = False
  On Error GoTo ErrorHandler
  If StrComp("Not", Left$(Trim$(txt.Text), 3)) Then
    Dummy = CDbl(txt.Text)
'    If Dummy < 0# Then GoTo NegativeNumberError
    If IsError Then txt.SetFocus
  Else
    txt.Text = Temp_Text
  End If
  GoTo ContinueSub

ErrorHandler:
  IsError = True
  'frmAirWaterProperties.Print "Error Occurred"
  'MsgBox "Incorrect Value Will Be Replaced By Previous Value", , "Invalid Data Error"
  txt.Text = Temp_Text
  Resume

NegativeNumberError:
  IsError = True
  txt.Text = Temp_Text
  txt.SetFocus

ContinueSub:

End Sub

Sub TextNumberChanged_WithUnits(ConversionFactor As Double, NewVal As Double, ValueChanged As Integer, txt As TextBox, Temp_Text As String)
Dim Dummy1 As Double, Dummy2 As Double

  If (Temp_Text = "") Then
    ValueChanged = True
    Exit Sub
  End If

  If (ConversionFactor < 0) Then
    Dummy1 = ReverseTemperatureConversion(-ConversionFactor - 1, CDbl(txt.Text))
  Else
    Dummy1 = CDbl(txt.Text) / ConversionFactor
  End If
  Dummy2 = CDbl(Temp_Text)

  ValueChanged = True
  If (txt.Text = Temp_Text) Then ValueChanged = False
  If (Abs(Dummy1 - Dummy2) < NUMBER_CHANGING_CRITERIA) Then ValueChanged = False

  If (ValueChanged) Then
    NewVal = Dummy1
  End If

End Sub

Function Unitted_LostFocus(WhatUnits As Integer, txt As TextBox, UnitBox As ComboBox, NewVal As Double, Temp_Text As String)
Dim Dummy1 As Double
Dim Dummy2 As Double
Dim IsNew As Integer
Dim ConversionFactor As Double

  IsNew = False

  Call TextHandleError(IsError, txt, Temp_Text)
  If (IsError) Then GoTo EntryError
  If (Not HaveValue(CDbl(txt))) Then GoTo EntryError
  
  Select Case WhatUnits
    Case 0
      'No units.
      Dummy1 = CDbl(txt.Text)
    Case UNITS_TEMPERATURE
      'Handle special case of temperature conversions.
      Dummy1 = ReverseTemperatureConversion(CInt(UnitBox.ListIndex), CDbl(txt.Text))
    Case Else
      'Handle non-temperature units.
      ConversionFactor = GetConversionFactor(WhatUnits, CInt(UnitBox.ListIndex))
      Dummy1 = CDbl(txt.Text) / ConversionFactor
  End Select
  
  If (Temp_Text = "") Then
    IsNew = True
    NewVal = Dummy1
  Else
    Dummy2 = CDbl(Temp_Text)
    IsNew = True
    If (txt.Text = Temp_Text) Then IsNew = False
    If (Abs(Dummy1 - Dummy2) < NUMBER_CHANGING_CRITERIA) Then
      IsNew = False
    End If
  End If
  
  If (IsNew) Then NewVal = Dummy1
  
  Unitted_LostFocus = IsNew
  Exit Function

EntryError:
  IsNew = False
  txt.Text = Temp_Text
  txt.SetFocus
  Unitted_LostFocus = False

End Function

Sub Unitted_NumberUpdate(UnitBox As ComboBox)
'Generate a click event on this unitbox, thus resulting
'in an update of the associated number into its correct
'units (if it has changed and/or been badly written to).
Dim SaveListIndex As Integer

  SaveListIndex = UnitBox.ListIndex
  UnitBox.ListIndex = -1
  UnitBox.ListIndex = SaveListIndex

End Sub

Sub Unitted_UnitChange(WhatUnits As Integer, MemVal As Double, UnitBox As ComboBox, txt As Control)
Dim ConversionFactor As Double
Dim Dummy1 As Double

  If (WhatUnits = UNITS_TEMPERATURE) Then
    'Handle special case of temperature conversions.
    Dummy1 = TemperatureConversion(CInt(UnitBox.ListIndex), MemVal)
  Else
    'Handle non-temperature units.
    ConversionFactor = GetConversionFactor(WhatUnits, CInt(UnitBox.ListIndex))
    Dummy1 = MemVal * ConversionFactor
  End If

  txt = FormatSciNumber(Dummy1)

End Sub

