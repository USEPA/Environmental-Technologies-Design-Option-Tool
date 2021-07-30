Option Strict Off
Option Explicit On
Module UNITSYS0
	
	Structure rec_unitsys
		Dim deleted As Short
		Dim formx As System.Windows.Forms.Form
		Dim lblx As System.Windows.Forms.Control
		Dim TxtX As System.Windows.Forms.Control
		Dim CboX As ComboBox
		Dim UnitType As String
		Dim baseunit As String
		Dim format_entry As String
		Dim format_display As String
		Dim current_value As Double
		Dim has_units As Short
		Dim dirty As Short
	End Structure
	
	Public unitsys() As rec_unitsys
	
	Public Const WHICHFORMAT_FORMAT_ENTRY As Short = 1
	Public Const WHICHFORMAT_FORMAT_DISPLAY As Short = 2
	
	Public Most_Recent_GotFocus As Short
	
	
	
	
	Const UNITSYS0_declarations_end As Short = 0
	
	
	Sub unitsys_clear_dirty_flag_on_form(ByRef formx As System.Windows.Forms.Form)
		Dim i As Short
		
		For i = 1 To UBound(unitsys)
			If (Not unitsys(i).deleted) Then
				If (formx.Handle.ToInt32 = unitsys(i).formx.Handle.ToInt32) Then
					unitsys(i).dirty = False
				End If
			End If
		Next i
		
	End Sub


	Sub unitsys_control_cbox_click(ByRef CboX As ComboBox)
		Dim H As Short
		Dim NewUnits As String

		'UPGRADE_WARNING: Couldn't resolve default property of object unitsys_lookup_cbox(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		H = unitsys_lookup_cbox(CboX)
		If (H = -1) Then Exit Sub

		'UPGRADE_WARNING: Couldn't resolve default property of object CboX.ListIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CboX.List. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		NewUnits = CboX.Items(CboX.SelectedIndex)
		Call unitsys_display_a_number(H, NewUnits, WHICHFORMAT_FORMAT_DISPLAY)

	End Sub


	Sub unitsys_control_txtx_gotfocus(ByRef TxtX As System.Windows.Forms.Control)
		Dim H As Short
		Dim nowunit As String
		Dim Sindex As Integer
		'GET HANDLE.
		'UPGRADE_WARNING: Couldn't resolve default property of object unitsys_lookup_txtx(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		H = unitsys_lookup_txtx(TxtX) 'txtx.name
		If (H = -1) Then Exit Sub
		'DISPLAY NUMBER USING DATA-ENTRY FORMAT
		If (unitsys(H).has_units) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object unitsys(H).CboX.ListIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object unitsys().CboX.List. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Sindex = unitsys(H).CboX.SelectedIndex
			If Sindex = -1 Then Sindex = 0    ' Do not know why it can be -1 further investigation is needed Shang
			nowunit = unitsys(H).CboX.Items(Sindex)
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
	Sub unitsys_control_txtx_lostfocus(ByRef TxtX As System.Windows.Forms.Control, ByRef NewValue As Double)
		Dim H As Short
		Dim nowunit As String
		Dim Sindex As Integer
		'GET HANDLE.
		'UPGRADE_WARNING: Couldn't resolve default property of object unitsys_lookup_txtx(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		H = unitsys_lookup_txtx(TxtX)
		If (H = -1) Then Exit Sub
		'STORE NEW DATA INTO VARIABLE.
		unitsys(H).current_value = NewValue
		'FORMAT FOR DISPLAY.
		If (unitsys(H).has_units) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object unitsys(H).CboX.ListIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object unitsys().CboX.List. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Sindex = unitsys(H).CboX.SelectedIndex
			If Sindex = -1 Then Sindex = 0    ' Do not know why it can be -1 further investigation is needed Shang

			nowunit = unitsys(H).CboX.Items(unitsys(H).CboX.SelectedIndex)
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
		Dim H As Short
		Dim NewValue As Double
		Dim TxtX As System.Windows.Forms.Control
		'GET HANDLE.
		H = Most_Recent_GotFocus
		If (H <= 0) Then Exit Sub
		'FORCE MOST RECENT TEXTBOX TO LOSE FOCUS (WITHOUT CHANGES).
		NewValue = unitsys(H).current_value
		TxtX = unitsys(H).TxtX
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
	Function unitsys_control_txtx_lostfocus_validate(ByRef TxtX As System.Windows.Forms.Control, ByRef Val_Low As Double, ByRef Val_High As Double, ByRef NewValue As Double, ByRef Raise_Dirty_Flag As Boolean) As Object
		Dim H As Short
		Dim dbl As Double
		Dim msg As String
		Dim Sindex As Integer
		Dim newvalue_oldunits As Double
		Dim newvalue_baseunits As Double
		Dim nowunits As String
		
		Dim OldValue As Double
		
		'UPGRADE_WARNING: Couldn't resolve default property of object unitsys_lookup_txtx(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		H = unitsys_lookup_txtx(TxtX)
		If (H = -1) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object unitsys_control_txtx_lostfocus_validate. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			unitsys_control_txtx_lostfocus_validate = False
			Exit Function
		End If
		
		'DEFAULT: DATA HAS NOT CHANGED.
		Raise_Dirty_Flag = False
		
		'READ IN THE OLD VALUE FROM UNIT SYSTEM MEMORY TO LOCAL MEMORY.
		OldValue = unitsys(H).current_value
		NewValue = OldValue 'SET NEWVALUE IN CASE AN ERROR OCCURS!
		
		'IS THIS THE CONTROL THAT MOST RECENTLY HAD FOCUS ACCORDING
		'TO MY INTERNAL RECORDS?  IF NOT, WE HAVE TO INVALIDATE
		'THE LOSTFOCUS.
		If (H = Most_Recent_GotFocus) Then
			'THIS IS OKAY; DO NOTHING.
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object unitsys_control_txtx_lostfocus_validate. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			unitsys_control_txtx_lostfocus_validate = False
			NewValue = OldValue
			Exit Function
		End If
		
		'DETERMINE UNITS; IF AN ERROR OCCURS, THIS WILL ALLOW THE
		'OLD VALUE TO BE REDISPLAYED IN ITS PROPER UNITS.
		If (unitsys(H).has_units) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object unitsys(H).CboX.ListIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object unitsys().CboX.List. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Sindex = unitsys(H).CboX.SelectedIndex
			If Sindex = -1 Then Sindex = 0    ' Do not know why it can be -1 further investigation is needed Shang

			nowunits = unitsys(H).CboX.Items(Sindex)
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
			'UPGRADE_WARNING: Couldn't resolve default property of object unitsys_control_txtx_lostfocus_validate. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			unitsys_control_txtx_lostfocus_validate = False
			If (unitsys(H).has_units) Then
				msg = "The data entered (" & Chr(34) & TxtX.Text & Chr(34) & " " & nowunits & ") is outside of the range " & Trim(Str(Val_Low)) & " to " & Trim(Str(Val_High)) & " " & unitsys(H).baseunit & ".  This data has been replaced with its previous value."
				Call unitsys_display_a_number(H, nowunits, WHICHFORMAT_FORMAT_DISPLAY)
				'ADDED ON 8/27/98.
				NewValue = OldValue
			Else
				msg = "The data entered (" & Chr(34) & TxtX.Text & Chr(34) & ") is outside of the range " & Trim(Str(Val_Low)) & " to " & Trim(Str(Val_High)) & ".  This data has been replaced with its previous value."
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
		'UPGRADE_WARNING: Couldn't resolve default property of object unitsys_control_txtx_lostfocus_validate. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		unitsys_control_txtx_lostfocus_validate = True
		
		'UPDATE DIRTY FLAG IF NEEDED.
		If (System.Math.Abs(((OldValue + 0.00000001) / (NewValue + 0.00000001)) - 1) > 0.00001) Then
			unitsys(H).dirty = True
			'MsgBox "Dirty flag raised!"
			Raise_Dirty_Flag = True
		End If
		
exit_err_unitsys_control_txtx_lostfocus_validate: 
		Exit Function
		
err_unitsys_control_txtx_lostfocus_validate: 
		'UPGRADE_WARNING: Couldn't resolve default property of object unitsys_control_txtx_lostfocus_validate. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		unitsys_control_txtx_lostfocus_validate = False
		msg = "The data entered is invalid (" & Chr(34) & TxtX.Text & Chr(34) & ").  This data has been replaced with its previous value."
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
	Sub unitsys_convert(ByRef UnitType As String, ByRef unit_from As String, ByRef unit_to As String, ByRef val_from As Double, ByRef val_to As Double)
		Dim Found As Short
		
		Dim uf As String
		Dim ut As String
		Dim temp_degK As Double
		Dim factor_from As Double
		Dim factor_to As Double
		
		Found = False
		If (UCase(UnitType) = "TEMPERATURE") Then
			uf = UCase(unit_from)
			If (uf = "K") Then
				temp_degK = val_from
			ElseIf (uf = "C") Then 
				temp_degK = val_from + 273.15
			ElseIf (uf = "R") Then 
				temp_degK = val_from * 5# / 9#
			ElseIf (uf = "F") Then 
				temp_degK = 273.15 + (val_from - 32#) * (5# / 9#)
			End If
			ut = UCase(unit_to)
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
			'UPGRADE_WARNING: Couldn't resolve default property of object unitsys_convert_getfactor(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			factor_from = unitsys_convert_getfactor(UnitType, unit_from)
			'UPGRADE_WARNING: Couldn't resolve default property of object unitsys_convert_getfactor(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
	
	
	Function unitsys_convert_getfactor(ByRef UnitType As String, ByRef unitname As String) As Object
		Dim X As Double 'RETURN VALUE
		Dim ut As String
		Dim un As String
		
		X = 0#
		ut = UCase(Trim(UnitType))
		un = LCase(Trim(unitname))
		'changed all from UCase to LCase bc of the nasty µ

		If (ut = "LENGTH") Then
			If (un = "m") Then X = 1.0#
			If (un = "cm") Then X = 0.01
			If (un = "ft") Then X = 0.3048
			If (un = "in") Then X = 0.0254
		End If
		If (ut = "MASS") Then
			If (un = "kg") Then X = 1.0#
			If (un = "g") Then X = 1.0# / 1000.0#
			If (un = "lb") Then X = 0.45359237
		End If
		If (ut = "TIME") Then
			If (un = "s") Then X = 1.0#
			If (un = "min") Then X = 1.0# * 60.0#
			If (un = "hr") Then X = 1.0# * 60.0# * 60.0#
			If (un = "d") Then X = 1.0# * 60.0# * 60.0# * 24.0#
			If (un = "year") Then X = 1.0# * 60.0# * 60.0# * 24.0# * 365.25
		End If
		If (ut = "INVERSE_TIME") Then
			If (un = "1/s") Then X = 1.0#
			If (un = "1/min") Then X = 1.0# / 60.0#
			If (un = "1/hr") Then X = 1.0# / 60.0# / 60.0#
			If (un = "1/day") Then X = 1.0# / 60.0# / 60.0# / 24.0#
			If (un = "1/year") Then X = 1.0# / 60.0# / 60.0# / 24.0# / 365.25
		End If
		If (ut = "REACTION_SOLIDPHASE") Then
			If (un = "1/s") Then X = 1.0#
			If (un = "1/min") Then X = 1.0# / 60.0#
			If (un = "1/hr") Then X = 1.0# / 60.0# / 60.0#
			If (un = "1/day") Then X = 1.0# / 60.0# / 60.0# / 24.0#
			If (un = "1/year") Then X = 1.0# / 60.0# / 60.0# / 24.0# / 365.25
		End If
		If (ut = "REACTION_LIQUIDPHASE") Then
			If (un = "l/µmol-s") Then X = 1.0#
			If (un = "cm³/µmol-s") Then X = 1.0# / 1000.0#
		End If
		If (ut = "REACTION_GASPHASE") Then
			If (un = "l/µmol-s") Then X = 1.0#
			If (un = "cm³/µmol-s") Then X = 1.0# / 1000.0#
		End If
		If (ut = "LANGMUIR_QM") Then
			If (un = "µmol/g") Then X = 1.0#
		End If
		If (ut = "LANGMUIR_B") Then
			If (un = "l/µmol") Then X = 1.0#
		End If
		If (ut = "FLOW_VOLUMETRIC") Then
			If (un = "m³/s") Then X = 1.0#
			If (un = "m³/d") Then X = 1.0# / (60.0# * 60.0# * 24.0#)
			If (un = "cm³/s") Then X = 1.0# / (100.0# * 100.0# * 100.0#)
			If (un = "ml/min") Then X = 1.0# / (100.0# * 100.0# * 100.0#) / (60.0#)
			If (un = "ft³/s") Then X = 1.0# / (35.31466672)
			If (un = "ft³/d") Then X = 1.0# / (35.31466672) / (60.0# * 60.0# * 24.0#)
			If (un = "gpm") Then X = 1.0# / (264.1720524) / (60.0#)
			If (un = "gpd") Then X = 1.0# / (264.1720524) / (60.0# * 60.0# * 24.0#)
			If (un = "mgd") Then X = (1000.0# * 1000.0#) / (264.1720524) / (60.0# * 60.0# * 24.0#)
			''''If (un = "FT³/MIN") Then X = 1# / (35.31466672) / (60# * 24#)
			If (un = "ft³/min") Then X = 1.0# / (35.31466672) / 60.0#
		End If
		If (ut = "DENSITY") Then
			If (un = "g/ml") Then X = 1.0# * 1000.0#
			If (un = "kg/m³") Then X = 1.0#
			'UPGRADE_ISSUE: The preceding line couldn't be parsed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="82EBB1AE-1FCB-4FEF-9E6C-8736A316F8A7"'
			If (un = "lb/ft³") Then X = 1.0# * (0.45359237) / (0.028316847)
			If (un = "lb/gal") Then X = 1.0# * (0.45359237) / (0.00378541178)
		End If
		If (ut = "CONCENTRATION") Then
			If (un = "g/l") Then X = 1.0# * 1000.0# * 1000.0#
			If (un = "mg/l") Then X = 1.0# * 1000.0#
			If (un = "µg/l") Then X = 1.0#
			If (un = "ng/l") Then X = 1.0# / 1000
		End If
			If (ut = "PRESSURE") Then
			''If (un = "N/M²") Then X = 1#
			'If (un = "PA") Then x = 1#
			'If (un = "LBF/IN²") Then x = 1# * 6894.75729
			'If (un = "ATM") Then x = 1# * 101325#
			If (un = "pa") Then X = 1.0#
			If (un = "kpa") Then X = 1.0# / (1.0# / 1000.0#)
			If (un = "bars") Then X = 1.0# / (1.0# / 100000.0#)
			If (un = "atm") Then X = 1.0# / (1.0# / 101325.0#)
			If (un = "psi") Then X = 1.0# / (14.696 / 101325.0#)
			If (un = "mmhg") Then X = 1.0# / (760.0# / 101325.0#)
			If (un = "mh20") Then X = 1.0# / (10.333 / 101325.0#)
			If (un = "fth20") Then X = 1.0# / (33.9 / 101325.0#)
			If (un = "inhg") Then X = 1.0# / (29.921 / 101325.0#)
		End If
		If (ut = "VELOCITY") Then
			If (un = "m/s") Then X = 1.0#
			If (un = "m/hr") Then X = 1.0# * 0.0002777777
			If (un = "ft/s") Then X = 1.0# * 0.3048
			If (un = "ft/hr") Then X = 1.0# * 0.3048 * 0.0002777777
		End If
		If (ut = "MOLAR_VOLUME") Then
			If (un = "m³/kmol") Then X = 1.0# * 0.001
			If (un = "m³/gmol") Then X = 1.0#
			If (un = "l/gmol") Then X = 1.0# * 0.001
			If (un = "ml/gmol") Then X = 1.0# * 0.000001
		End If
		If (ut = "VISCOSITY") Then
			If (un = "kg/m-s") Then X = 1.0#
			If (un = "g/cm-s") Then X = 1.0# * 0.1
			If (un = "cp") Then X = 1.0# * 0.001
		End If
		If (ut = "MOLECULAR_WEIGHT") Then
			If (un = "mg/mmol") Then X = 1.0#
			If (un = "µg/µmol") Then X = 1.0#
			If (un = "g/gmol") Then X = 1.0#
			If (un = "kg/kmol") Then X = 1.0#
		End If
		If (ut = "VOLUME") Then
			If (un = "m") Then X = 1.0#
			If (un = "cm") Then X = 0.000001
			If (un = "liter") Then X = 0.001
			If (un = "ft") Then X = 0.028316846592
			If (un = "gal") Then X = 0.003785411784
		End If
		If (ut = "MASS_EMISSION_RATE") Then
			If (un = "µg/s") Then X = 1.0#
			If (un = "µg/min") Then X = 1.0# / 60.0#
			If (un = "mg/s") Then X = 1.0# * 1000.0#
			If (un = "mg/min") Then X = 1000.0# / 60.0#

			'If (un = "µG/S") Then X = 1#
			'If (un = "µG/MIN") Then X = 1# * 1000#
			'If (un = "MG/S") Then X = 1# / 60#
			'If (un = "MG/MIN") Then X = 1000# / 60#
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object unitsys_convert_getfactor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		unitsys_convert_getfactor = X
		
	End Function
	
	
	'PARAMETERS:
	'- h: handle to unit control
	'- nowunit: current units of this control (if any)
	'- which_format: 1=format_entry, 2=format_display
	Sub unitsys_display_a_number(ByRef H As Short, ByRef nowunit As String, ByRef which_format As Short)
		Dim num_in_properunits As Double
		Dim use_format As String
		Dim Number_To_Display As Double
		Dim ForceToAFormat As Boolean
		'DETERMINE ACTUAL NUMBER TO DISPLAY ON THE SCREEN.
		If (Not unitsys(H).has_units) Then
			Number_To_Display = unitsys(H).current_value
		Else
			Call unitsys_convert(unitsys(H).UnitType, unitsys(H).baseunit, nowunit, unitsys(H).current_value, num_in_properunits)
			Number_To_Display = num_in_properunits
		End If
		'DETERMINE APPROPRIATE NUMERIC FORMAT TO USE.
		Select Case which_format
			Case WHICHFORMAT_FORMAT_ENTRY
				ForceToAFormat = Not (Trim(unitsys(H).format_entry) = "")
				If (ForceToAFormat) Then use_format = unitsys(H).format_entry
				If (Not ForceToAFormat) Then use_format = GetDoubleFormatLonger(Number_To_Display)
			Case WHICHFORMAT_FORMAT_DISPLAY
				ForceToAFormat = Not (Trim(unitsys(H).format_display) = "")
				If (ForceToAFormat) Then use_format = unitsys(H).format_display
				If (Not ForceToAFormat) Then use_format = GetDoubleFormat(Number_To_Display)
		End Select
		'DISPLAY THE NUMBER.
		'UPGRADE_WARNING: Couldn't resolve default property of object unitsys().TxtX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        unitsys(H).TxtX.Text = VB6.Format(Number_To_Display, use_format)
	End Sub
	
	
	Function unitsys_get_numerical_value(ByRef TxtX As System.Windows.Forms.Control) As Double
		Dim H As Short
		'UPGRADE_WARNING: Couldn't resolve default property of object unitsys_lookup_txtx(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		H = unitsys_lookup_txtx(TxtX)
		unitsys_get_numerical_value = unitsys(H).current_value
	End Function
	
	
	Sub unitsys_initialize()
		'UPGRADE_WARNING: Lower bound of array unitsys was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim Preserve unitsys(1)
		unitsys(1).deleted = True
	End Sub
	
	
	Function unitsys_is_any_data_dirty(ByRef formx As System.Windows.Forms.Form) As Object
		Dim RetVal As Short
		Dim i As Short
		
		RetVal = False
		For i = 1 To UBound(unitsys)
			If (Not unitsys(i).deleted) Then
				If (formx.Handle.ToInt32 = unitsys(i).formx.Handle.ToInt32) Then
					If (unitsys(i).dirty) Then
						RetVal = True
						Exit For
					End If
				End If
			End If
		Next i
		'UPGRADE_WARNING: Couldn't resolve default property of object unitsys_is_any_data_dirty. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		unitsys_is_any_data_dirty = RetVal
	End Function
	
	
	Function unitsys_lookup_cbox(ByRef CboX As System.Windows.Forms.Control) As Object
		Dim i As Short
		Dim Found As Short
		Dim H As Short
		
		Found = False
		For i = 1 To UBound(unitsys)
			If (Not unitsys(i).deleted) Then
				If (unitsys(i).has_units) Then
					If (CboX.Handle.ToInt32 = unitsys(i).CboX.Handle.ToInt32) Then
						Found = True
						H = i
					End If
				End If
			End If
		Next i
		If (Not Found) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object unitsys_lookup_cbox. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			unitsys_lookup_cbox = -1
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object unitsys_lookup_cbox. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			unitsys_lookup_cbox = H
		End If
		
	End Function
	
	
	Function unitsys_lookup_txtx(ByRef TxtX As System.Windows.Forms.Control) As Object
		Dim i As Short
		Dim Found As Short
		Dim H As Short
		Found = False
		For i = 1 To UBound(unitsys)
			If (Not unitsys(i).deleted) Then
				'txtx.hwnd
				'unitsys(i).txtx.hWnd
				If (TxtX.Handle.ToInt32 = unitsys(i).TxtX.Handle.ToInt32) Then
					Found = True
					H = i
				End If
			End If
		Next i
		If (Not Found) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object unitsys_lookup_txtx. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			unitsys_lookup_txtx = -1
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object unitsys_lookup_txtx. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			unitsys_lookup_txtx = H
		End If
	End Function


	''''Sub unitsys_populate_units(h As Integer, UnitType As String, initunit As String)
	Sub unitsys_populate_units0(ByRef Cbc As ComboBox, ByRef UnitType As String, ByRef initunit As String)
		Dim u As String
		''''Dim Cbc As Control
		Dim i As Short
		Dim Found As Short
		u = LCase(UnitType)
		''''Set Cbc = unitsys(h).CboX
		'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.Clear. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Cbc.Items.Clear()
		Found = False
		If (u = "length") Then
			Found = True
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("m")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("cm")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("ft")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("in")
		End If
		If (u = "mass") Then
			Found = True
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("kg")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("g")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("lb")
		End If
		If (u = "time") Then
			Found = True
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("s")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("min")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("hr")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("d")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("year")
		End If
		If (u = "inverse_time") Then
			Found = True
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("1/s")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("1/min")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("1/hr")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("1/day")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("1/year")
		End If
		If (u = "reaction_solidphase") Then
			Found = True
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("1/s")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("1/min")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("1/hr")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("1/day")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("1/year")
		End If
		If (u = "reaction_liquidphase") Then
			Found = True
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("L/µmol-s")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("cm³/µmol-s")
		End If
		If (u = "reaction_gasphase") Then
			Found = True
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("L/µmol-s")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("cm³/µmol-s")
		End If
		If (u = "langmuir_qm") Then
			Found = True
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("µmol/g")
		End If
		If (u = "langmuir_b") Then
			Found = True
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("L/µmol")
		End If
		If (u = "flow_volumetric") Then
			Found = True
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("m³/s")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("m³/d")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("cm³/s")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("mL/min")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("ft³/s")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("ft³/d")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("gpm")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("gpd")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("MGD")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("ft³/min")
		End If
		If (u = "density") Then
			Found = True
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("g/mL")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("kg/m³")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("lb/ft³")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("lb/gal")
		End If
		If (u = "temperature") Then
			Found = True
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("K")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("C")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("R")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("F")
		End If
		If (u = "concentration") Then
			Found = True
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("µg/L")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("mg/L")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("g/L")

			Cbc.Items.Add("ng/L")
		End If
		If (u = "pressure") Then
			Found = True
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("Pa") '"N/m²"
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("kPa")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("bars")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("atm")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("psi")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("mmHg")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("mH20")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("ftH20")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("inHg")
		End If
		If (u = "velocity") Then
			Found = True
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("m/s")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("m/hr")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("ft/s")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("ft/hr")
		End If
		If (u = "molar_volume") Then
			Found = True
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("m³/kmol")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("m³/gmol")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("L/gmol")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("mL/gmol")
		End If
		If (u = "viscosity") Then
			Found = True
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("kg/m-s")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("g/cm-s")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("cP")
		End If
		If (u = "molecular_weight") Then
			Found = True
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("mg/mmol")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("µg/µmol")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("g/gmol")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("kg/kmol")
		End If
		If (u = "volume") Then
			Found = True
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("m³""")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("cm³""")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("liter")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("ft³""")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("gal")
		End If
		If (u = "mass_emission_rate") Then
			Found = True
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("µg/s")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("µg/min")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("mg/s")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("mg/min")
		End If
		If (Not Found) Then
			''''Call local_unitsys_populate_units(H, UnitType)
			Call local_unitsys_populate_units(Cbc, UnitType)
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.ListCount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For i = 0 To Cbc.Items.Count - 1
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.List. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If (UCase(initunit) = UCase(Cbc.Items(i))) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.ListIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Cbc.SelectedIndex = i
				Exit For
			End If
		Next i

	End Sub
	Sub unitsys_populate_units(ByRef H As Short, ByRef UnitType As String, ByRef initunit As String)
		Dim Cbc As System.Windows.Forms.Control
		Cbc = unitsys(H).CboX
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
	Sub unitsys_register(ByRef formx As System.Windows.Forms.Form, ByRef lblx As System.Windows.Forms.Control, ByRef TxtX As System.Windows.Forms.Control, ByRef CboX As System.Windows.Forms.Control, ByRef UnitType As String, ByRef initunit As String, ByRef baseunit As String, ByRef format_entry As String, ByRef format_display As String, ByRef initnum As Double, ByRef has_units As Short)
		Dim H As Short
		Dim i As Short
		Dim Found As Short
		Dim ub As Short
		
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
			'UPGRADE_WARNING: Lower bound of array unitsys was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim Preserve unitsys(ub + 1)
			H = ub + 1
		Else
			H = Found
		End If
		
		'==== INSTALL NEW UNIT CONTROL
		'TELL ANY PROCESSES THIS CONTROL DOES NOT EXIST YET!
		unitsys(H).dirty = False
		unitsys(H).deleted = True
		unitsys(H).formx = formx
		unitsys(H).lblx = lblx
		unitsys(H).TxtX = TxtX
		unitsys(H).CboX = CboX
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
	
	
	Sub unitsys_set_number_in_base_units(ByRef TxtX As System.Windows.Forms.Control, ByRef new_value As Double)
		Dim H As Short
		Dim nowunit As String
		Dim sindex As Integer
		'UPGRADE_WARNING: Couldn't resolve default property of object unitsys_lookup_txtx(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		H = unitsys_lookup_txtx(TxtX)
		If (H = -1) Then Exit Sub
		unitsys(H).current_value = new_value
		'Call unitsys_display_a_number(h, unitsys(h).baseunit, WHICHFORMAT_FORMAT_DISPLAY)
		If (unitsys(H).has_units) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object unitsys(H).CboX.ListIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object unitsys().CboX.List. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sindex = unitsys(H).CboX.SelectedIndex   'not sure why SelectedIndex is not changed from -1, investigation needed Shang
			If (sindex = -1) Then
				sindex = 0
			End If
			nowunit = unitsys(H).CboX.Items(sindex)
			Call unitsys_display_a_number(H, nowunit, WHICHFORMAT_FORMAT_DISPLAY)
		Else
			Call unitsys_display_a_number(H, "", WHICHFORMAT_FORMAT_DISPLAY)
		End If
	End Sub


	Function unitsys_get_units(ByRef CboX As ComboBox) As String
		'UPGRADE_WARNING: Couldn't resolve default property of object CboX.ListIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CboX.List. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Dim Sindex As Integer
		Sindex = CboX.SelectedIndex
		If Sindex = -1 Then Sindex = 0          'Do not know why it can be -1 should be initialized correctly Shang
		unitsys_get_units = CboX.Items(Sindex)
	End Function
	Sub unitsys_set_units(ByRef TxtX As System.Windows.Forms.Control, ByRef new_units As String)
		Dim H As Short
		Dim i As Short
		Dim Found As Short
		Dim max As Short
		'UPGRADE_WARNING: Couldn't resolve default property of object unitsys_lookup_txtx(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		H = unitsys_lookup_txtx(TxtX)
		If (H <= 0) Then Exit Sub
		Found = False
		'UPGRADE_WARNING: Couldn't resolve default property of object unitsys().CboX.ListCount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		max = unitsys(H).CboX.Items.Count - 1
		For i = 0 To max
			'UPGRADE_WARNING: Couldn't resolve default property of object unitsys().CboX.List. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If (Trim(UCase(unitsys(H).CboX.Items(i))) = Trim(UCase(new_units))) Then
				Found = True
				Exit For
			End If
		Next i
		If (Found) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object unitsys().CboX.ListIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			unitsys(H).CboX.SelectedIndex = i
		End If
	End Sub
	
	
	Sub unitsys_unregister_one_control(ByRef H As Short)
		unitsys(H).deleted = True
		'UPGRADE_NOTE: Object unitsys().formx may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		unitsys(H).formx = Nothing
		'UPGRADE_NOTE: Object unitsys().lblx may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		unitsys(H).lblx = Nothing
		'UPGRADE_NOTE: Object unitsys().TxtX may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		unitsys(H).TxtX = Nothing
		'UPGRADE_NOTE: Object unitsys().CboX may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		unitsys(H).CboX = Nothing
	End Sub
	
	
	Sub unitsys_unregister_all_on_form(ByRef formx As System.Windows.Forms.Form)
		Dim i As Short
		For i = 1 To UBound(unitsys)
			If (Not unitsys(i).deleted) Then
				'If (formx.Handle.ToInt32 = unitsys(i).formx.Handle.ToInt32) Then
				'formx.name
				If (formx.Name = unitsys(i).formx.Name) Then   'Shang
					Call unitsys_unregister_one_control(i)
				End If
			End If
		Next i
	End Sub
End Module