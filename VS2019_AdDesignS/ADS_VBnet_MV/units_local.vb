Option Strict Off
Option Explicit On
Module UNITS_LOCAL
	
	''COMMUNICATIONS WITH frmBed.
	'Global frmBed_Copy_Of_CURRENT_BEDCOMPONENT As Integer
	'Global frmBed_Copy_Of_TempBedDef As BedDefinition_Type
	'Global frmBed_Copy_Of_TempProj As Project_Type
	'      'NOTE: THIS RECORD IS USED TO OBTAIN THE VALUES
	'      'OF THE FOLLOWING VARIABLES (CURRENTLY DISPLAYED COMPONENT):
	'      '    - molecular weight
	'      '    - Freundlich 1/n
	
	
	
	
	
	Const UNITS_LOCAL_declarations_end As Short = 0


	Function local_unitsys_convert_getfactor_FreundlichK(ByRef UnitType As String, ByRef factor1 As Double, ByRef factor2 As Double) As Double
		Dim X As Double
		Dim stt As String
		stt = (Trim(UnitType)).ToUpperInvariant
		Dim b1 As Boolean = Strings.StrComp(stt, "(MG/G)*(L/MG)^(1/N)")
		Dim b2 As Boolean = Strings.StrComp(stt, "(MMOL/G)*(L/MMOL)^(1/N)")
		Dim b3 As Boolean = Strings.StrComp(stt, "(µG/G)*(L/µG)^(1/N)")
		Dim b4 As Boolean = Strings.StrComp(stt, "(µMOL/G)*(L/µMOL)^(1/N)")

		'	Select Case Trim(UCase(UnitType))
		Select Case stt
				Case "(MG/G)*(L/MG)^(1/N)"
					X = 1.0#
				Case "(MMOL/G)*(L/MMOL)^(1/N)"
					X = 1.0# / factor1
				Case "(µG/G)*(L/µG)^(1/N)"
					X = 1.0# / factor2
				Case "(µMOL/G)*(L/µMOL)^(1/N)"
					X = 1.0# / factor1 / factor2
			End Select
			local_unitsys_convert_getfactor_FreundlichK = X
	End Function


	Sub local_unitsys_convert(ByRef UnitType As String, ByRef unit_from As String, ByRef unit_to As String, ByRef val_from As Double, ByRef val_to As Double)
		Dim now_MW As Double
		Dim now_OneOverN As Double
		Dim factor1 As Double
		Dim factor2 As Double
		Dim factor_from As Double
		Dim factor_to As Double
		UnitType = Trim(UCase(UnitType))
		If (UnitType = Trim(UCase("freundlich_k"))) Then
			'now_MW = frmBed_Copy_Of_TempProj.Components( _
			''    frmBed_Copy_Of_CURRENT_BEDCOMPONENT).xwt
			'now_OneOverN = frmBed_Copy_Of_TempBedDef.BedComponents( _
			''    frmBed_Copy_Of_CURRENT_BEDCOMPONENT).XN
			now_MW = Component(0).MW
			now_OneOverN = Component(0).Use_OneOverN
			factor1 = (now_MW) ^ (now_OneOverN - 1#)
			factor2 = (1000#) ^ (1# - now_OneOverN)
			factor_from = local_unitsys_convert_getfactor_FreundlichK(unit_from, factor1, factor2)
			factor_to = local_unitsys_convert_getfactor_FreundlichK(unit_to, factor1, factor2)
			'PERFORM THE CONVERSION.
			val_to = val_from / factor_to * factor_from
		End If
	End Sub
	''''Sub local_unitsys_populate_units(H As Integer, UnitType As String)
	Sub local_unitsys_populate_units(ByRef Cbc As ComboBox, ByRef UnitType As String)
		UnitType = Trim(UCase(UnitType))
		If (UnitType = Trim(UCase("freundlich_k"))) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("(mg/g)*(L/mg)^(1/n)")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("(mmol/g)*(L/mmol)^(1/n)")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("(µg/g)*(L/µg)^(1/n)")
			'UPGRADE_WARNING: Couldn't resolve default property of object Cbc.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Cbc.Items.Add("(µmol/g)*(L/µmol)^(1/n)")
		End If
		''''  If (UnitType = Trim$(UCase$("freundlich_k"))) Then
		''''    unitsys(H).CboX.AddItem "(mg/g)*(L/mg)^(1/n)"
		''''    unitsys(H).CboX.AddItem "(mmol/g)*(L/mmol)^(1/n)"
		''''    unitsys(H).CboX.AddItem "(µg/g)*(L/µg)^(1/n)"
		''''    unitsys(H).CboX.AddItem "(µmol/g)*(L/µmol)^(1/n)"
		''''  End If
	End Sub
End Module