Attribute VB_Name = "UNITSYS0"
Option Explicit

Type rec_unitsys
  deleted As Integer
  formx As Form
  lblx As Control
  TxtX As Control
  CboX As Control
  UnitType As String
  baseunit As String
  format_entry As String
  format_display As String
  current_value As Double
  has_units As Integer
  dirty As Integer
End Type

Global unitsys() As rec_unitsys

Global Const WHICHFORMAT_FORMAT_ENTRY = 1
Global Const WHICHFORMAT_FORMAT_DISPLAY = 2

Global Most_Recent_GotFocus As Integer




Const UNITSYS0_declarations_end = 0


Sub unitsys_clear_dirty_flag_on_form(formx As Form)
Dim i As Integer

  For i = 1 To UBound(unitsys)
    If (Not unitsys(i).deleted) Then
      If (formx.hwnd = unitsys(i).formx.hwnd) Then
        unitsys(i).dirty = False
      End If
    End If
  Next i

End Sub


Sub unitsys_control_cbox_click(CboX As Control)
Dim H As Integer
Dim NewUnits As String

  H = unitsys_lookup_cbox(CboX)
  If (H = -1) Then Exit Sub

  NewUnits = CboX.List(CboX.ListIndex)
  Call unitsys_display_a_number(H, NewUnits, WHICHFORMAT_FORMAT_DISPLAY)

End Sub


Sub unitsys_control_txtx_gotfocus(TxtX As Control)
Dim H As Integer
Dim nowunit As String
  'GET HANDLE.
  H = unitsys_lookup_txtx(TxtX)     'txtx.name
  If (H = -1) Then Exit Sub
  'DISPLAY NUMBER USING DATA-ENTRY FORMAT
  If (unitsys(H).has_units) Then
    nowunit = unitsys(H).CboX.List(unitsys(H).CboX.ListIndex)
    Call unitsys_display_a_number(H, nowunit, WHICHFORMAT_FORMAT_ENTRY)
  Else
    Call unitsys_display_a_number(H, "", WHICHFORMAT_FORMAT_ENTRY)
  End If
  'HIGHLIGHT THE SELECTION.
  'unitsys(h).txtx.SelStart = 0
  'unitsys(h).txtx.SelLength = Len(unitsys(h).txtx)
  Call Global_GotFocus(unitsys(H).TxtX)
  'SET MOST RECENT.
  Most_Recent_GotFocus = H
End Sub
'PARAMETERS:
'- txtx: the text control
'- newvalue: new numerical value in BASE units!
Sub unitsys_control_txtx_lostfocus(TxtX As Control, NewValue As Double)
Dim H As Integer
Dim nowunit As String
  'GET HANDLE.
  H = unitsys_lookup_txtx(TxtX)
  If (H = -1) Then Exit Sub
  'STORE NEW DATA INTO VARIABLE.
  unitsys(H).current_value = NewValue
  'FORMAT FOR DISPLAY.
  If (unitsys(H).has_units) Then
    nowunit = unitsys(H).CboX.List(unitsys(H).CboX.ListIndex)
    Call unitsys_display_a_number(H, nowunit, WHICHFORMAT_FORMAT_DISPLAY)
  Else
    Call unitsys_display_a_number(H, "", WHICHFORMAT_FORMAT_DISPLAY)
  End If
  'DE-HIGHLIGHT THE SELECTION.
  Call Global_LostFocus(TxtX)
  'SET MOST RECENT.
  Most_Recent_GotFocus = -1
End Sub
Sub unitsys_control_MostRecent_Force_lostfocus()
Dim H As Integer
Dim NewValue As Double
Dim TxtX As Control
  'GET HANDLE.
  H = Most_Recent_GotFocus
  If (H <= 0) Then Exit Sub
  'FORCE MOST RECENT TEXTBOX TO LOSE FOCUS (WITHOUT CHANGES).
  NewValue = unitsys(H).current_value
  Set TxtX = unitsys(H).TxtX
  Call unitsys_control_txtx_lostfocus(TxtX, NewValue)
End Sub


'RETURNS:
'- FALSE = Failed to validate this new data
'- TRUE = New data validated OK
'PARAMETERS:
'- txtx = Text control (INPUT)
'- val_low = Low value in the base unit (double) (INPUT)
'- val_high = High value in the base unit (double) (INPUT)
'- newvalue = The new value output to the caller (double) (RETURNED)
'- raise_dirty_flag = Whether the data has changed (boolean) (RETURNED)
Function unitsys_control_txtx_lostfocus_validate( _
    TxtX As Control, _
    Val_Low As Double, _
    Val_High As Double, _
    NewValue As Double, _
    Raise_Dirty_Flag As Boolean)
Dim H As Integer
Dim dbl As Double
Dim msg As String

Dim newvalue_oldunits As Double
Dim newvalue_baseunits As Double
Dim nowunits As String

Dim OldValue As Double

  H = unitsys_lookup_txtx(TxtX)
  If (H = -1) Then
    unitsys_control_txtx_lostfocus_validate = False
    Exit Function
  End If

  'DEFAULT: DATA HAS NOT CHANGED.
  Raise_Dirty_Flag = False

  'READ IN THE OLD VALUE FROM UNIT SYSTEM MEMORY TO LOCAL MEMORY.
  OldValue = unitsys(H).current_value
  NewValue = OldValue   'SET NEWVALUE IN CASE AN ERROR OCCURS!
  
  'IS THIS THE CONTROL THAT MOST RECENTLY HAD FOCUS ACCORDING
  'TO MY INTERNAL RECORDS?  IF NOT, WE HAVE TO INVALIDATE
  'THE LOSTFOCUS.
  If (H = Most_Recent_GotFocus) Then
    'THIS IS OKAY; DO NOTHING.
  Else
    unitsys_control_txtx_lostfocus_validate = False
    NewValue = OldValue
    Exit Function
  End If

  'DETERMINE UNITS; IF AN ERROR OCCURS, THIS WILL ALLOW THE
  'OLD VALUE TO BE REDISPLAYED IN ITS PROPER UNITS.
  If (unitsys(H).has_units) Then
    nowunits = unitsys(H).CboX.List(unitsys(H).CboX.ListIndex)
  Else
    nowunits = ""
  End If
  
  'CONVERT STRING TO DOUBLE.
  On Error GoTo err_unitsys_control_txtx_lostfocus_validate
  NewValue = CDbl(TxtX.Text)

  'CONVERT VALUE TO BASE UNITS (IF APPLICABLE).
  If (unitsys(H).has_units) Then
    ''''nowunits = unitsys(h).cbox.List(unitsys(h).cbox.ListIndex)
    newvalue_oldunits = NewValue
    Call unitsys_convert(unitsys(H).UnitType, nowunits, unitsys(H).baseunit, newvalue_oldunits, newvalue_baseunits)
    NewValue = newvalue_baseunits
  Else
    ''''nowunits = ""
  End If

  'PERFORM RANGE CHECK.
  If ((NewValue < Val_Low) Or (NewValue > Val_High)) Then
    unitsys_control_txtx_lostfocus_validate = False
    If (unitsys(H).has_units) Then
      msg = "The data entered (" & Chr$(34) & TxtX.Text & Chr$(34) & _
          " " & nowunits & ") is outside of the range " & Trim$(Str$(Val_Low)) & _
          " to " & Trim$(Str$(Val_High)) & " " & unitsys(H).baseunit & _
          ".  This data has been replaced with its previous value."
      Call unitsys_display_a_number(H, nowunits, WHICHFORMAT_FORMAT_DISPLAY)
      'ADDED ON 8/27/98.
      NewValue = OldValue
    Else
      msg = "The data entered (" & Chr$(34) & TxtX.Text & Chr$(34) & _
          ") is outside of the range " & Trim$(Str$(Val_Low)) & " to " _
          & Trim$(Str$(Val_High)) & ".  This data has been replaced with its previous value."
      'xaxaxa
      ''''unitsys(h).txtx = Format$(unitsys(h).current_value, unitsys(h).format_entry)
      Call unitsys_display_a_number(H, "", WHICHFORMAT_FORMAT_DISPLAY)
      'ADDED ON 8/27/98.
      NewValue = OldValue
    End If
    ''''MsgBox msg
    'NOTE: DISPLAYING A MESSAGE BOX INTERFERES
    'WITH THE LOSTFOCUS-GOTFOCUS; SKIPPING THIS STEP.
    Exit Function
  End If
  
  'TELL CALLER THIS DATA IS VALID!!
  unitsys_control_txtx_lostfocus_validate = True
  
  'UPDATE DIRTY FLAG IF NEEDED.
  If (Abs(((OldValue + 0.00000001) / (NewValue + 0.00000001)) - 1) > 0.00001) Then
    unitsys(H).dirty = True
    'MsgBox "Dirty flag raised!"
    Raise_Dirty_Flag = True
  End If
  
exit_err_unitsys_control_txtx_lostfocus_validate:
  Exit Function

err_unitsys_control_txtx_lostfocus_validate:
  unitsys_control_txtx_lostfocus_validate = False
  msg = "The data entered is invalid (" & Chr$(34) & TxtX.Text & Chr$(34) & _
      ").  This data has been replaced with its previous value."
  'xaxaxa
  ''''unitsys(h).txtx = Format$(unitsys(h).current_value, unitsys(h).format_display)
  Call unitsys_display_a_number(H, nowunits, WHICHFORMAT_FORMAT_DISPLAY)
  ''''MsgBox msg
  'NOTE: DISPLAYING A MESSAGE BOX INTERFERES
  'WITH THE LOSTFOCUS-GOTFOCUS; SKIPPING THIS STEP.
  GoTo exit_err_unitsys_control_txtx_lostfocus_validate

End Function


'PARAMETERS:
'- unittype = type of unit
'- unit_from = unit to convert from
'- unit_to = unit to convert to
'- val_from = value to convert from
'- val_to = results (OUTPUT)
Sub unitsys_convert( _
    UnitType As String, _
    unit_from As String, _
    unit_to As String, _
    val_from As Double, _
    val_to As Double)
Dim Found As Integer

Dim uf As String
Dim ut As String
Dim temp_degK As Double
Dim factor_from As Double
Dim factor_to As Double

  Found = False
  If (UCase$(UnitType) = "TEMPERATURE") Then
    uf = UCase$(unit_from)
    If (uf = "K") Then
      temp_degK = val_from
    ElseIf (uf = "C") Then
      temp_degK = val_from + 273.15
    ElseIf (uf = "R") Then
      temp_degK = val_from * 5# / 9#
    ElseIf (uf = "F") Then
      temp_degK = 273.15 + (val_from - 32#) * (5# / 9#)
    End If
    ut = UCase$(unit_to)
    If (ut = "K") Then
      val_to = temp_degK
    ElseIf (ut = "C") Then
      val_to = temp_degK - 273.15
    ElseIf (ut = "R") Then
      val_to = temp_degK * 9# / 5#
    ElseIf (ut = "F") Then
      val_to = 32 + (temp_degK - 273.15) * (9# / 5#)
    End If
    Found = True
  End If
  If (Not Found) Then
    factor_from = unitsys_convert_getfactor(UnitType, unit_from)
    factor_to = unitsys_convert_getfactor(UnitType, unit_to)
    If (factor_from = 0#) Or (factor_to = 0#) Then
      'CONVERT USING APPLICATION-SPECIFIC UNIT CODE.
      Call local_unitsys_convert(UnitType, unit_from, unit_to, val_from, val_to)
      Exit Sub
    End If
    
    'CONVERT USING SIMPLE CONVERSION FACTORS.
    val_to = val_from / factor_to * factor_from
  End If
  
End Sub


Function unitsys_convert_getfactor(UnitType As String, unitname As String)
Dim X As Double    'RETURN VALUE
Dim ut As String
Dim un As String

  X = 0#
  ut = UCase$(Trim$(UnitType))
  un = UCase$(Trim$(unitname))
  If (ut = "LENGTH") Then
    If (un = "M") Then X = 1#
    If (un = "CM") Then X = 0.01
    If (un = "FT") Then X = 0.3048
    If (un = "IN") Then X = 0.0254
  End If
  If (ut = "MASS") Then
    If (un = "KG") Then X = 1#
    If (un = "G") Then X = 1# / 1000#
    If (un = "LB") Then X = 0.45359237
  End If
  If (ut = "TIME") Then
    If (un = "S") Then X = 1#
    If (un = "MIN") Then X = 1# * 60#
    If (un = "HR") Then X = 1# * 60# * 60#
    If (un = "D") Then X = 1# * 60# * 60# * 24#
    If (un = "YEAR") Then X = 1# * 60# * 60# * 24# * 365.25
  End If
  If (ut = "INVERSE_TIME") Then
    If (un = "1/S") Then X = 1#
    If (un = "1/MIN") Then X = 1# / 60#
    If (un = "1/HR") Then X = 1# / 60# / 60#
    If (un = "1/DAY") Then X = 1# / 60# / 60# / 24#
    If (un = "1/YEAR") Then X = 1# / 60# / 60# / 24# / 365.25
  End If
  If (ut = "REACTION_SOLIDPHASE") Then
    If (un = "1/S") Then X = 1#
    If (un = "1/MIN") Then X = 1# / 60#
    If (un = "1/HR") Then X = 1# / 60# / 60#
    If (un = "1/DAY") Then X = 1# / 60# / 60# / 24#
    If (un = "1/YEAR") Then X = 1# / 60# / 60# / 24# / 365.25
  End If
  If (ut = "REACTION_LIQUIDPHASE") Then
    If (un = "L/µMOL-S") Then X = 1#
    If (un = "CM³/µMOL-S") Then X = 1# / 1000#
  End If
  If (ut = "REACTION_GASPHASE") Then
    If (un = "L/µMOL-S") Then X = 1#
    If (un = "CM³/µMOL-S") Then X = 1# / 1000#
  End If
  If (ut = "LANGMUIR_QM") Then
    If (un = "µMOL/G") Then X = 1#
  End If
  If (ut = "LANGMUIR_B") Then
    If (un = "L/µMOL") Then X = 1#
  End If
  If (ut = "FLOW_VOLUMETRIC") Then
    If (un = "M³/S") Then X = 1#
    If (un = "M³/D") Then X = 1# / (60# * 60# * 24#)
    If (un = "CM³/S") Then X = 1# / (100# * 100# * 100#)
    If (un = "ML/MIN") Then X = 1# / (100# * 100# * 100#) / (60#)
    If (un = "FT³/S") Then X = 1# / (35.31466672)
    If (un = "FT³/D") Then X = 1# / (35.31466672) / (60# * 60# * 24#)
    If (un = "GPM") Then X = 1# / (264.1720524) / (60#)
    If (un = "GPD") Then X = 1# / (264.1720524) / (60# * 60# * 24#)
    If (un = "MGD") Then X = (1000# * 1000#) / (264.1720524) / (60# * 60# * 24#)
    ''''If (un = "FT³/MIN") Then X = 1# / (35.31466672) / (60# * 24#)
    If (un = "FT³/MIN") Then X = 1# / (35.31466672) / 60#
  End If
  If (ut = "DENSITY") Then
    If (un = "G/ML") Then X = 1# * 1000#
    If (un = "KG/M³") Then X = 1#
    If (un = "LB/FT³") Then X = 1# * (0.45359237) / (0.028316847)
    If (un = "LB/GAL") Then X = 1# * (0.45359237) / (0.00378541178)
  End If
  If (ut = "CONCENTRATION") Then
    If (un = "G/L") Then X = 1# * 1000# * 1000#
    If (un = "MG/L") Then X = 1# * 1000#
    If (un = "µG/L") Then X = 1#
  End If
  If (ut = "PRESSURE") Then
    ''If (un = "N/M²") Then X = 1#
    'If (un = "PA") Then x = 1#
    'If (un = "LBF/IN²") Then x = 1# * 6894.75729
    'If (un = "ATM") Then x = 1# * 101325#
    If (un = "PA") Then X = 1#
    If (un = "KPA") Then X = 1# / (1# / 1000#)
    If (un = "BARS") Then X = 1# / (1# / 100000#)
    If (un = "ATM") Then X = 1# / (1# / 101325#)
    If (un = "PSI") Then X = 1# / (14.696 / 101325#)
    If (un = "MMHG") Then X = 1# / (760# / 101325#)
    If (un = "MH20") Then X = 1# / (10.333 / 101325#)
    If (un = "FTH20") Then X = 1# / (33.9 / 101325#)
    If (un = "INHG") Then X = 1# / (29.921 / 101325#)
  End If
  If (ut = "VELOCITY") Then
    If (un = "M/S") Then X = 1#
    If (un = "M/HR") Then X = 1# * 0.0002777777
    If (un = "FT/S") Then X = 1# * 0.3048
    If (un = "FT/HR") Then X = 1# * 0.3048 * 0.0002777777
  End If
  If (ut = "MOLAR_VOLUME") Then
    If (un = "M³/KMOL") Then X = 1# * 0.001
    If (un = "M³/GMOL") Then X = 1#
    If (un = "L/GMOL") Then X = 1# * 0.001
    If (un = "ML/GMOL") Then X = 1# * 0.000001
  End If
  If (ut = "VISCOSITY") Then
    If (un = "KG/M-S") Then X = 1#
    If (un = "G/CM-S") Then X = 1# * 0.1
    If (un = "CP") Then X = 1# * 0.001
  End If
  If (ut = "MOLECULAR_WEIGHT") Then
    If (un = "MG/MMOL") Then X = 1#
    If (un = "µG/µMOL") Then X = 1#
    If (un = "G/GMOL") Then X = 1#
    If (un = "KG/KMOL") Then X = 1#
  End If
  If (ut = "VOLUME") Then
    If (un = "M³") Then X = 1#
    If (un = "CM³") Then X = 0.000001
    If (un = "LITER") Then X = 0.001
    If (un = "FT³") Then X = 0.028316846592
    If (un = "GAL") Then X = 0.003785411784
  End If
  If (ut = "MASS_EMISSION_RATE") Then
    If (un = "µG/S") Then X = 1#
    If (un = "µG/MIN") Then X = 1# / 60#
    If (un = "MG/S") Then X = 1# * 1000#
    If (un = "MG/MIN") Then X = 1000# / 60#
    
    'If (un = "µG/S") Then X = 1#
    'If (un = "µG/MIN") Then X = 1# * 1000#
    'If (un = "MG/S") Then X = 1# / 60#
    'If (un = "MG/MIN") Then X = 1000# / 60#
  End If
  unitsys_convert_getfactor = X

End Function


'PARAMETERS:
'- h: handle to unit control
'- nowunit: current units of this control (if any)
'- which_format: 1=format_entry, 2=format_display
Sub unitsys_display_a_number(H As Integer, nowunit As String, which_format As Integer)
Dim num_in_properunits As Double
Dim use_format As String
Dim Number_To_Display As Double
Dim ForceToAFormat As Boolean
  'DETERMINE ACTUAL NUMBER TO DISPLAY ON THE SCREEN.
  If (Not unitsys(H).has_units) Then
    Number_To_Display = unitsys(H).current_value
  Else
    Call unitsys_convert(unitsys(H).UnitType, _
        unitsys(H).baseunit, nowunit, _
        unitsys(H).current_value, num_in_properunits)
    Number_To_Display = num_in_properunits
  End If
  'DETERMINE APPROPRIATE NUMERIC FORMAT TO USE.
  Select Case which_format
    Case WHICHFORMAT_FORMAT_ENTRY:
      ForceToAFormat = Not (Trim$(unitsys(H).format_entry) = "")
      If (ForceToAFormat) Then use_format = unitsys(H).format_entry
      If (Not ForceToAFormat) Then use_format = GetDoubleFormatLonger(Number_To_Display)
    Case WHICHFORMAT_FORMAT_DISPLAY:
      ForceToAFormat = Not (Trim$(unitsys(H).format_display) = "")
      If (ForceToAFormat) Then use_format = unitsys(H).format_display
      If (Not ForceToAFormat) Then use_format = GetDoubleFormat(Number_To_Display)
  End Select
  'DISPLAY THE NUMBER.
  unitsys(H).TxtX = Format$(Number_To_Display, use_format)
End Sub


Function unitsys_get_numerical_value(TxtX As Control) As Double
Dim H As Integer
  H = unitsys_lookup_txtx(TxtX)
  unitsys_get_numerical_value = unitsys(H).current_value
End Function


Sub unitsys_initialize()
  ReDim Preserve unitsys(1 To 1)
  unitsys(1).deleted = True
End Sub


Function unitsys_is_any_data_dirty(formx As Form)
Dim RetVal As Integer
Dim i As Integer

  RetVal = False
  For i = 1 To UBound(unitsys)
    If (Not unitsys(i).deleted) Then
      If (formx.hwnd = unitsys(i).formx.hwnd) Then
        If (unitsys(i).dirty) Then
          RetVal = True
          Exit For
        End If
      End If
    End If
  Next i
  unitsys_is_any_data_dirty = RetVal
End Function


Function unitsys_lookup_cbox(CboX As Control)
Dim i As Integer
Dim Found As Integer
Dim H As Integer

  Found = False
  For i = 1 To UBound(unitsys)
    If (Not unitsys(i).deleted) Then
      If (unitsys(i).has_units) Then
        If (CboX.hwnd = unitsys(i).CboX.hwnd) Then
          Found = True
          H = i
        End If
      End If
    End If
  Next i
  If (Not Found) Then
    unitsys_lookup_cbox = -1
  Else
    unitsys_lookup_cbox = H
  End If

End Function


Function unitsys_lookup_txtx(TxtX As Control)
Dim i As Integer
Dim Found As Integer
Dim H As Integer
  Found = False
  For i = 1 To UBound(unitsys)
    If (Not unitsys(i).deleted) Then
'txtx.hwnd
'unitsys(i).txtx.hWnd
      If (TxtX.hwnd = unitsys(i).TxtX.hwnd) Then
        Found = True
        H = i
      End If
    End If
  Next i
  If (Not Found) Then
    unitsys_lookup_txtx = -1
  Else
    unitsys_lookup_txtx = H
  End If
End Function


''''Sub unitsys_populate_units(h As Integer, UnitType As String, initunit As String)
Sub unitsys_populate_units0( _
    Cbc As Control, _
    UnitType As String, _
    initunit As String)
Dim u As String
''''Dim Cbc As Control
Dim i As Integer
Dim Found As Integer
  u = LCase$(UnitType)
  ''''Set Cbc = unitsys(h).CboX
  Cbc.Clear
  Found = False
  If (u = "length") Then
    Found = True
    Cbc.AddItem "m"
    Cbc.AddItem "cm"
    Cbc.AddItem "ft"
    Cbc.AddItem "in"
  End If
  If (u = "mass") Then
    Found = True
    Cbc.AddItem "kg"
    Cbc.AddItem "g"
    Cbc.AddItem "lb"
  End If
  If (u = "time") Then
    Found = True
    Cbc.AddItem "s"
    Cbc.AddItem "min"
    Cbc.AddItem "hr"
    Cbc.AddItem "d"
    Cbc.AddItem "year"
  End If
  If (u = "inverse_time") Then
    Found = True
    Cbc.AddItem "1/s"
    Cbc.AddItem "1/min"
    Cbc.AddItem "1/hr"
    Cbc.AddItem "1/day"
    Cbc.AddItem "1/year"
  End If
  If (u = "reaction_solidphase") Then
    Found = True
    Cbc.AddItem "1/s"
    Cbc.AddItem "1/min"
    Cbc.AddItem "1/hr"
    Cbc.AddItem "1/day"
    Cbc.AddItem "1/year"
  End If
  If (u = "reaction_liquidphase") Then
    Found = True
    Cbc.AddItem "L/µmol-s"
    Cbc.AddItem "cm³/µmol-s"
  End If
  If (u = "reaction_gasphase") Then
    Found = True
    Cbc.AddItem "L/µmol-s"
    Cbc.AddItem "cm³/µmol-s"
  End If
  If (u = "langmuir_qm") Then
    Found = True
    Cbc.AddItem "µmol/g"
  End If
  If (u = "langmuir_b") Then
    Found = True
    Cbc.AddItem "L/µmol"
  End If
  If (u = "flow_volumetric") Then
    Found = True
    Cbc.AddItem "m³/s"
    Cbc.AddItem "m³/d"
    Cbc.AddItem "cm³/s"
    Cbc.AddItem "mL/min"
    Cbc.AddItem "ft³/s"
    Cbc.AddItem "ft³/d"
    Cbc.AddItem "gpm"
    Cbc.AddItem "gpd"
    Cbc.AddItem "MGD"
    Cbc.AddItem "ft³/min"
  End If
  If (u = "density") Then
    Found = True
    Cbc.AddItem "g/mL"
    Cbc.AddItem "kg/m³"
    Cbc.AddItem "lb/ft³"
    Cbc.AddItem "lb/gal"
  End If
  If (u = "temperature") Then
    Found = True
    Cbc.AddItem "K"
    Cbc.AddItem "C"
    Cbc.AddItem "R"
    Cbc.AddItem "F"
  End If
  If (u = "concentration") Then
    Found = True
    Cbc.AddItem "µg/L"
    Cbc.AddItem "mg/L"
    Cbc.AddItem "g/L"
  End If
  If (u = "pressure") Then
    Found = True
    Cbc.AddItem "Pa"      '"N/m²"
    Cbc.AddItem "kPa"
    Cbc.AddItem "bars"
    Cbc.AddItem "atm"
    Cbc.AddItem "psi"
    Cbc.AddItem "mmHg"
    Cbc.AddItem "mH20"
    Cbc.AddItem "ftH20"
    Cbc.AddItem "inHg"
  End If
  If (u = "velocity") Then
    Found = True
    Cbc.AddItem "m/s"
    Cbc.AddItem "m/hr"
    Cbc.AddItem "ft/s"
    Cbc.AddItem "ft/hr"
  End If
  If (u = "molar_volume") Then
    Found = True
    Cbc.AddItem "m³/kmol"
    Cbc.AddItem "m³/gmol"
    Cbc.AddItem "L/gmol"
    Cbc.AddItem "mL/gmol"
  End If
  If (u = "viscosity") Then
    Found = True
    Cbc.AddItem "kg/m-s"
    Cbc.AddItem "g/cm-s"
    Cbc.AddItem "cP"
  End If
  If (u = "molecular_weight") Then
    Found = True
    Cbc.AddItem "mg/mmol"
    Cbc.AddItem "µg/µmol"
    Cbc.AddItem "g/gmol"
    Cbc.AddItem "kg/kmol"
  End If
  If (u = "volume") Then
    Found = True
    Cbc.AddItem "m³"
    Cbc.AddItem "cm³"
    Cbc.AddItem "liter"
    Cbc.AddItem "ft³"
    Cbc.AddItem "gal"
  End If
  If (u = "mass_emission_rate") Then
    Found = True
    Cbc.AddItem "µg/s"
    Cbc.AddItem "µg/min"
    Cbc.AddItem "mg/s"
    Cbc.AddItem "mg/min"
  End If
  If (Not Found) Then
    ''''Call local_unitsys_populate_units(H, UnitType)
    Call local_unitsys_populate_units(Cbc, UnitType)
  End If
  For i = 0 To Cbc.ListCount - 1
    If (UCase$(initunit) = UCase$(Cbc.List(i))) Then
      Cbc.ListIndex = i
      Exit For
    End If
  Next i
  
End Sub
Sub unitsys_populate_units(H As Integer, UnitType As String, initunit As String)
Dim Cbc As Control
  Set Cbc = unitsys(H).CboX
  Call unitsys_populate_units0(Cbc, UnitType, initunit)
End Sub


'PURPOSE: Register a new unit control.
'Parameters:
'- formx: pointer to form that contains unit control
'- lblx: description label
'- txtx: data-entry textbox
'- cbox: unit checkbox
'- unittype: unit type (string)
'- initunit: initial unit (string)
'- baseunit: base unit, used to store numerical value (string)
'- format_entry: numerical format for data-entry
'- format_display: numerical format for display
'- initnum: initial number in _base_ units! (double)
'- has_units: if unit-conversion features enabled, this is TRUE
Sub unitsys_register( _
    formx As Form, _
    lblx As Control, _
    TxtX As Control, _
    CboX As Control, _
    UnitType As String, _
    initunit As String, _
    baseunit As String, _
    format_entry As String, _
    format_display As String, _
    initnum As Double, _
    has_units As Integer)
Dim H As Integer
Dim i As Integer
Dim Found As Integer
Dim ub As Integer
  
Dim num_in_properunits As Double

  Found = 0
  ub = UBound(unitsys)
  For i = 1 To ub
    If (unitsys(i).deleted) Then
      Found = i
      Exit For
    End If
  Next i
  If (Found = 0) Then
    ReDim Preserve unitsys(1 To ub + 1)
    H = ub + 1
  Else
    H = Found
  End If
  
  '==== INSTALL NEW UNIT CONTROL
  'TELL ANY PROCESSES THIS CONTROL DOES NOT EXIST YET!
  unitsys(H).dirty = False
  unitsys(H).deleted = True
  Set unitsys(H).formx = formx
  Set unitsys(H).lblx = lblx
  Set unitsys(H).TxtX = TxtX
  Set unitsys(H).CboX = CboX
'txtx.hwnd
  unitsys(H).UnitType = UnitType
  unitsys(H).baseunit = baseunit
  unitsys(H).format_entry = format_entry
  unitsys(H).format_display = format_display
  unitsys(H).has_units = has_units
  unitsys(H).current_value = initnum
  Call unitsys_display_a_number(H, initunit, WHICHFORMAT_FORMAT_DISPLAY)
  If (unitsys(H).has_units) Then
    Call unitsys_populate_units(H, UnitType, initunit)
  End If
  
  'TELL ANY PROCESSES THIS CONTROL IS READY FOR OPERATION.
  unitsys(H).deleted = False

End Sub


Sub unitsys_set_number_in_base_units(TxtX As Control, new_value As Double)
Dim H As Integer
Dim nowunit As String
  H = unitsys_lookup_txtx(TxtX)
  If (H = -1) Then Exit Sub
  unitsys(H).current_value = new_value
  'Call unitsys_display_a_number(h, unitsys(h).baseunit, WHICHFORMAT_FORMAT_DISPLAY)
  If (unitsys(H).has_units) Then
    nowunit = unitsys(H).CboX.List(unitsys(H).CboX.ListIndex)
    Call unitsys_display_a_number(H, nowunit, WHICHFORMAT_FORMAT_DISPLAY)
  Else
    Call unitsys_display_a_number(H, "", WHICHFORMAT_FORMAT_DISPLAY)
  End If
End Sub


Function unitsys_get_units(CboX As Control) As String
  unitsys_get_units = CboX.List(CboX.ListIndex)
End Function
Sub unitsys_set_units(TxtX As Control, new_units As String)
Dim H As Integer
Dim i As Integer
Dim Found As Integer
Dim max As Integer
  H = unitsys_lookup_txtx(TxtX)
  If (H <= 0) Then Exit Sub
  Found = False
  max = unitsys(H).CboX.ListCount - 1
  For i = 0 To max
    If (Trim$(UCase$(unitsys(H).CboX.List(i))) = Trim$(UCase$(new_units))) Then
      Found = True
      Exit For
    End If
  Next i
  If (Found) Then
    unitsys(H).CboX.ListIndex = i
  End If
End Sub


Sub unitsys_unregister_one_control(H As Integer)
  unitsys(H).deleted = True
  Set unitsys(H).formx = Nothing
  Set unitsys(H).lblx = Nothing
  Set unitsys(H).TxtX = Nothing
  Set unitsys(H).CboX = Nothing
End Sub


Sub unitsys_unregister_all_on_form(formx As Form)
Dim i As Integer
  For i = 1 To UBound(unitsys)
    If (Not unitsys(i).deleted) Then
      If (formx.hwnd = unitsys(i).formx.hwnd) Then
        'formx.name
        Call unitsys_unregister_one_control(i)
      End If
    End If
  Next i
End Sub


