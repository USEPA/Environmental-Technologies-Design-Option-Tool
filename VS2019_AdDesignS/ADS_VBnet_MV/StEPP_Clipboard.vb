Option Strict Off
Option Explicit On
Module StEPP_Clipboard
	
	
	Sub Do_ImportClipboard(ByRef Was_Aborted As Boolean)
		Dim num_lines As Short
		Dim cliptext As String
		Dim line_in As String
		Dim r As Short
		Dim link_pressure As Double
		Dim link_temperature As Double
		Dim link_ChemCount As Short
		Const CHEMPROP_MIN As Short = 0
		Const CHEMPROP_MAX As Short = 12
		'UPGRADE_WARNING: Lower bound of array link_ChemProp was changed from CHEMPROP_MIN,1 to 0,0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		Dim link_ChemProp(CHEMPROP_MAX, 1) As Double
		'UPGRADE_WARNING: Lower bound of array link_ChemName was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		Dim link_ChemName(1) As String
		'UPGRADE_WARNING: Lower bound of array link_ChemCAS was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		Dim link_ChemCAS(1) As String
		'UPGRADE_WARNING: Lower bound of array link_ChemPropAvailable was changed from CHEMPROP_MIN,1 to 0,0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		Dim link_ChemPropAvailable(CHEMPROP_MAX, 1) As Short
		'UPGRADE_WARNING: Lower bound of array link_IsImportable was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		Dim link_IsImportable(1) As Short
		Dim i As Short
		Dim j As Short
		Const PROP_VAPORPRESSURE As Short = 0
		Const PROP_ACTIVITYCOEFFICIENT As Short = 1
		Const PROP_HENRYSCONSTANT As Short = 2
		Const PROP_MOLECULARWEIGHT As Short = 3
		Const PROP_NORMALBOILINGPOINT As Short = 4
		Const PROP_LIQUIDDENSITY As Short = 5
		Const PROP_MOLARVOLUMEATOPT As Short = 6
		Const PROP_MOLARVOLUMEATNBP As Short = 7
		Const PROP_REFRACTIVEINDEX As Short = 8
		Const PROP_AQUEOUSSOLUBILITY As Short = 9
		Const PROP_LOGKOW As Short = 10
		Const PROP_LIQUIDDIFFUSIVITY As Short = 11
		Const PROP_GASDIFFUSIVITY As Short = 12
		Dim Num_Imported As Short
		'UPGRADE_WARNING: Arrays in structure ThisComp may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim ThisComp As ComponentPropertyType
		Dim msg As String
		Dim vb3CrLf As String
		Dim Num_Failed As Short
		
		Was_Aborted = True
		On Error GoTo err_Do_ImportClipboard
		'UPGRADE_ISSUE: Clipboard method Clipboard.GetText was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		cliptext = My.Computer.Clipboard.GetText()
		cliptext = Parser_RemoveCharacters(Chr(10), cliptext)
		num_lines = Parser_GetNumArgs(Chr(13), cliptext)
		r = 1
		Call Parser_GetArg(Chr(13), cliptext, r, line_in)
		If (Trim(UCase(line_in)) <> Trim(UCase("1234567890:START_OF_STEPP_CLIPBOARD_EXPORT"))) Then
			GoTo err_nonfatal_err_Do_ImportClipboard
		End If
		r = r + 2
		Call Parser_GetArg(Chr(13), cliptext, r, line_in)
		link_pressure = CDbl(Val(line_in))
		If (link_pressure <= 0#) Then GoTo err_nonfatal_err_Do_ImportClipboard
		r = r + 2
		Call Parser_GetArg(Chr(13), cliptext, r, line_in)
		link_temperature = CDbl(Val(line_in))
		If (link_temperature <= 0#) Then GoTo err_nonfatal_err_Do_ImportClipboard
		r = r + 2
		Call Parser_GetArg(Chr(13), cliptext, r, line_in)
		link_ChemCount = CShort(Val(line_in))
		If (link_ChemCount <= 0) Then GoTo err_nonfatal_err_Do_ImportClipboard
		'UPGRADE_WARNING: Lower bound of array link_ChemProp was changed from CHEMPROP_MIN,1 to 0,0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim link_ChemProp(CHEMPROP_MAX, link_ChemCount)
		'UPGRADE_WARNING: Lower bound of array link_ChemName was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim link_ChemName(link_ChemCount)
		'UPGRADE_WARNING: Lower bound of array link_ChemCAS was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim link_ChemCAS(link_ChemCount)
		'UPGRADE_WARNING: Lower bound of array link_ChemPropAvailable was changed from CHEMPROP_MIN,1 to 0,0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim link_ChemPropAvailable(CHEMPROP_MAX, link_ChemCount)
		'UPGRADE_WARNING: Lower bound of array link_IsImportable was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim link_IsImportable(link_ChemCount)
		For i = 1 To link_ChemCount
			For j = CHEMPROP_MIN To CHEMPROP_MAX
				link_ChemPropAvailable(j, i) = True
			Next j
			r = r + 2
			Call Parser_GetArg(Chr(13), cliptext, r, line_in)
			link_ChemName(i) = Trim(UCase(line_in))
			If (link_ChemName(i) = "") Then GoTo err_nonfatal_err_Do_ImportClipboard
			r = r + 2
			Call Parser_GetArg(Chr(13), cliptext, r, line_in)
			link_ChemCAS(i) = Trim(UCase(line_in))
			'If (link_ChemCAS(i) = "") Then GoTo err_nonfatal_err_Do_ImportClipboard
			For j = CHEMPROP_MIN To CHEMPROP_MAX
				r = r + 2
				Call Parser_GetArg(Chr(13), cliptext, r, line_in)
				line_in = Trim(UCase(line_in))
				If (Trim(UCase("UNAVAILABLE")) = line_in) Then
					link_ChemPropAvailable(j, i) = False
				Else
					link_ChemProp(j, i) = CDbl(Val(line_in))
				End If
			Next j
		Next i
		r = r + 1
		Call Parser_GetArg(Chr(13), cliptext, r, line_in)
		If (Trim(UCase(line_in)) <> Trim(UCase("1234567890:END_OF_STEPP_CLIPBOARD_EXPORT"))) Then
			GoTo err_nonfatal_err_Do_ImportClipboard
		End If
		
		'ARE THERE ENOUGH EMPTY COMPONENT SLOTS REMAINING?
		If (Number_Component + link_ChemCount > Number_Compo_Max) Then
			Call Show_Error("Unable to import all of chemicals in file " & "because the maximum number of chemicals has been reached.")
			'Unload Me
			Exit Sub
		End If
		
		'DOES THE USER REALLY WANT TO IMPORT AT THIS TEMPERATURE AND PRESSURE?
		'---- I'VE DECIDED TO SKIP THIS STEP.  THE USER BEWARE.
		
		'DETERMINE WHICH COMPONENTS ARE IMPORTABLE.
		For i = 1 To link_ChemCount
			link_IsImportable(i) = True
			If (Not link_ChemPropAvailable(PROP_VAPORPRESSURE, i)) Then link_IsImportable(i) = False
			If (Not link_ChemPropAvailable(PROP_MOLECULARWEIGHT, i)) Then link_IsImportable(i) = False
			If (Not link_ChemPropAvailable(PROP_NORMALBOILINGPOINT, i)) Then link_IsImportable(i) = False
			If (Not link_ChemPropAvailable(PROP_LIQUIDDENSITY, i)) Then link_IsImportable(i) = False
			If (Not link_ChemPropAvailable(PROP_MOLARVOLUMEATNBP, i)) Then link_IsImportable(i) = False
			If (Not link_ChemPropAvailable(PROP_REFRACTIVEINDEX, i)) Then link_IsImportable(i) = False
			If (Not link_ChemPropAvailable(PROP_AQUEOUSSOLUBILITY, i)) Then link_IsImportable(i) = False
		Next i
		
		'IMPORT ALL IMPORTABLE COMPONENTS.
		Num_Imported = 0
		For i = 1 To link_ChemCount
			If (link_IsImportable(i)) Then
				Num_Imported = Num_Imported + 1
				Call SetComponentDefaults(ThisComp, -1)
				ThisComp.name = link_ChemName(i)
				ThisComp.Cas = CInt(Val(link_ChemCAS(i)))
				ThisComp.Vapor_Pressure = link_ChemProp(PROP_VAPORPRESSURE, i)
				ThisComp.MW = link_ChemProp(PROP_MOLECULARWEIGHT, i)
				ThisComp.BP = link_ChemProp(PROP_NORMALBOILINGPOINT, i)
				ThisComp.Liquid_Density = link_ChemProp(PROP_LIQUIDDENSITY, i) / 1000#
				ThisComp.MolarVolume = link_ChemProp(PROP_MOLARVOLUMEATNBP, i) * 1000#
				ThisComp.Refractive_Index = link_ChemProp(PROP_REFRACTIVEINDEX, i)
				ThisComp.Aqueous_Solubility = link_ChemProp(PROP_AQUEOUSSOLUBILITY, i)
				Number_Component = Number_Component + 1
				'UPGRADE_WARNING: Couldn't resolve default property of object Component(Number_Component). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Component(Number_Component) = ThisComp
				
				''Take care of miscellaneous screen B.S.
				'frmpfpsdm!cmdViewDimensionless.Enabled = True
				'frmpfpsdm!cmdEditComponent.Enabled = True
				'frmpfpsdm!cmdDeleteComponent.Enabled = True
				'frmpfpsdm!lstComponents.AddItem thiscomp.name
				'frmpfpsdm!cboSelectCompo.Enabled = True
				'frmpfpsdm!cboSelectCompo.AddItem thiscomp.name
				'If (Number_Component = Number_Compo_Max) Then
				'  frmpfpsdm!cmdAddComponent.Enabled = False
				'End If
				''Set index of the kinetic combo box to the new chemical
				'frmpfpsdm!cboSelectCompo.ListIndex = frmpfpsdm!cboSelectCompo.ListCount - 1
				''Update the corresponding kinetic data displayed
				'Call Update_Display_Kinetic
				'If (Number_Component > 0) Then
				'  frmpfpsdm!mnuRunItem(0).Enabled = True
				'  frmpfpsdm!mnuRunItem(1).Enabled = True
				'  frmpfpsdm!mnuRunItem(2).Enabled = True
				'  frmpfpsdm!mnuOptionsItem(0).Enabled = True
				'  frmpfpsdm!mnuOptionsItem(1).Enabled = True  'Variable Influent concentration
				'  frmpfpsdm!mnuOptionsItem(2).Enabled = True  'Variable Effluent concentration
				'End If
			End If
		Next i
		
		'DISPLAY WARNING/SUCCESS MESSAGE.
		vb3CrLf = Chr(13) & Chr(10)
		If (Num_Imported <> 0) Then
			msg = "Successfully imported " & Trim(Str(Num_Imported)) & " component"
			If (Num_Imported <> 1) Then msg = msg & "s"
			msg = msg & " from StEPP:" & vb3CrLf
			For i = 1 To link_ChemCount
				If (link_IsImportable(i)) Then
					msg = msg & "    " & Trim(link_ChemName(i)) & vb3CrLf
				End If
			Next i
			msg = msg & "The properties are for a "
			msg = msg & "pressure of " & Trim(Str(link_pressure)) & " Pa "
			msg = msg & "and a "
			msg = msg & "temperature of " & Trim(Str(link_temperature)) & " degrees Celcius." & vb3CrLf
			msg = msg & vb3CrLf
			msg = msg & "Don't forget to set the correct values of Freundlich K, "
			msg = msg & "Freundlich 1/n, and initial concentration for each "
			msg = msg & "of these components." & vb3CrLf
		Else
			msg = "Unable to import any components from StEPP." & vb3CrLf
		End If
		Num_Failed = link_ChemCount - Num_Imported
		If (Num_Failed <> 0) Then
			msg = msg & vb3CrLf
			msg = msg & "Failed to import the following component"
			If (Num_Failed <> 1) Then msg = msg & "s"
			msg = msg & ":" & vb3CrLf
			For i = 1 To link_ChemCount
				If (Not link_IsImportable(i)) Then
					msg = msg & "    " & Trim(link_ChemName(i)) & vb3CrLf
				End If
			Next i
			msg = msg & vb3CrLf
			msg = msg & "Important note: In order to successfully import a component "
			msg = msg & "from StEPP, the following properties must be available: "
			msg = msg & "vapor pressure, "
			msg = msg & "molecular weight, "
			msg = msg & "normal boiling point, "
			msg = msg & "liquid density, "
			msg = msg & "molar volume at the normal boiling point, "
			msg = msg & "refractive index, "
			msg = msg & "and aqueous solubility.  "
			msg = msg & "To force an import to occur, you may modify the user input "
			msg = msg & "value of the unavailable properties from within StEPP."
			msg = msg & vb3CrLf
		End If
		Call Show_Message(msg)
		Was_Aborted = False
		
exit_err_err_Do_ImportClipboard: 
		Exit Sub
err_nonfatal_err_Do_ImportClipboard: 
		Call Show_Error("An error occurred during the import process.")
		GoTo exit_err_err_Do_ImportClipboard
err_Do_ImportClipboard: 
		Call Show_Error("An error occurred during the import process.")
		Resume exit_err_err_Do_ImportClipboard
	End Sub
End Module