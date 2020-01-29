Attribute VB_Name = "StEPP_Clipboard"
Option Explicit


Sub Do_ImportClipboard( _
    Was_Aborted As Boolean, _
    Temp_Plant As TYPE_PlantDiagram)
Dim num_lines As Integer
Dim cliptext As String
Dim line_in As String
Dim r As Integer
Dim link_pressure As Double
Dim link_temperature As Double
Dim link_ChemCount As Integer
Const CHEMPROP_MIN = 0
Const CHEMPROP_MAX = 12
ReDim link_ChemProp(CHEMPROP_MIN To CHEMPROP_MAX, 1 To 1) As Double
ReDim link_ChemName(1 To 1) As String
ReDim link_ChemCAS(1 To 1) As String
ReDim link_ChemPropAvailable(CHEMPROP_MIN To CHEMPROP_MAX, 1 To 1) As Integer
ReDim link_AllPropsAvailable(1 To 1) As Integer
Dim i As Integer
Dim j As Integer
Const PROP_VAPORPRESSURE = 0
Const PROP_ACTIVITYCOEFFICIENT = 1
Const PROP_HENRYSCONSTANT = 2
Const PROP_MOLECULARWEIGHT = 3
Const PROP_NORMALBOILINGPOINT = 4
Const PROP_LIQUIDDENSITY = 5
Const PROP_MOLARVOLUMEATOPT = 6
Const PROP_MOLARVOLUMEATNBP = 7
Const PROP_REFRACTIVEINDEX = 8
Const PROP_AQUEOUSSOLUBILITY = 9
Const PROP_LOGKOW = 10
Const PROP_LIQUIDDIFFUSIVITY = 11
Const PROP_GASDIFFUSIVITY = 12
Dim Num_Imported As Integer
''''Dim ThisComp As ComponentPropertyType
Dim msg As String
Dim vb3CrLf As String
Dim Num_Failed As Integer

  Was_Aborted = True
  On Error GoTo err_Do_ImportClipboard
  cliptext = Clipboard.GetText()
  cliptext = Parser_RemoveCharacters(Chr$(10), cliptext)
  num_lines = Parser_GetNumArgs(Chr$(13), cliptext)
  r = 1
  Call Parser_GetArg(Chr$(13), cliptext, r, line_in)
  If (Trim$(UCase$(line_in)) <> Trim$(UCase$("1234567890:START_OF_STEPP_CLIPBOARD_EXPORT"))) Then
    GoTo err_nonfatal_err_Do_ImportClipboard
  End If
  r = r + 2
  Call Parser_GetArg(Chr$(13), cliptext, r, line_in)
  link_pressure = CDbl(Val(line_in))
  If (link_pressure <= 0#) Then GoTo err_nonfatal_err_Do_ImportClipboard
  r = r + 2
  Call Parser_GetArg(Chr$(13), cliptext, r, line_in)
  link_temperature = CDbl(Val(line_in))
  If (link_temperature <= 0#) Then GoTo err_nonfatal_err_Do_ImportClipboard
  r = r + 2
  Call Parser_GetArg(Chr$(13), cliptext, r, line_in)
  link_ChemCount = CInt(Val(line_in))
  If (link_ChemCount <= 0) Then GoTo err_nonfatal_err_Do_ImportClipboard
  ReDim link_ChemProp(CHEMPROP_MIN To CHEMPROP_MAX, 1 To link_ChemCount)
  ReDim link_ChemName(1 To link_ChemCount)
  ReDim link_ChemCAS(1 To link_ChemCount)
  ReDim link_ChemPropAvailable(CHEMPROP_MIN To CHEMPROP_MAX, 1 To link_ChemCount)
  ReDim link_AllPropsAvailable(1 To link_ChemCount)
  For i = 1 To link_ChemCount
    For j = CHEMPROP_MIN To CHEMPROP_MAX
      link_ChemPropAvailable(j, i) = True
    Next j
    r = r + 2
    Call Parser_GetArg(Chr$(13), cliptext, r, line_in)
    link_ChemName(i) = Trim$(UCase$(line_in))
    If (link_ChemName(i) = "") Then GoTo err_nonfatal_err_Do_ImportClipboard
    r = r + 2
    Call Parser_GetArg(Chr$(13), cliptext, r, line_in)
    link_ChemCAS(i) = Trim$(UCase$(line_in))
    'If (link_ChemCAS(i) = "") Then GoTo err_nonfatal_err_Do_ImportClipboard
    For j = CHEMPROP_MIN To CHEMPROP_MAX
      r = r + 2
      Call Parser_GetArg(Chr$(13), cliptext, r, line_in)
      line_in = Trim$(UCase$(line_in))
      If (Trim$(UCase$("UNAVAILABLE")) = line_in) Then
        link_ChemPropAvailable(j, i) = False
      Else
        link_ChemProp(j, i) = CDbl(Val(line_in))
      End If
    Next j
  Next i
  r = r + 1
  Call Parser_GetArg(Chr$(13), cliptext, r, line_in)
  If (Trim$(UCase$(line_in)) <> Trim$(UCase$("1234567890:END_OF_STEPP_CLIPBOARD_EXPORT"))) Then
    GoTo err_nonfatal_err_Do_ImportClipboard
  End If
  '
  ' ONLY ONE COMPONENT ALLOWED DURING THE IMPORT.
  '
  If (link_ChemCount <> 1) Then
    Call Show_Error("The clipboard contains " & _
        Trim$(Str$(link_ChemCount)) & _
        " chemicals.  Only one chemical may be " & _
        "imported into the FaVOr software.")
    Exit Sub
  End If
  '''''
  ''''' ARE THERE ENOUGH EMPTY COMPONENT SLOTS REMAINING?
  '''''
  ''''If (Number_Component + link_ChemCount > Number_Compo_Max) Then
  ''''  Call Show_Error("Unable to import all of chemicals in file " & _
  ''''      "because the maximum number of chemicals has been reached.")
  ''''  'Unload Me
  ''''  Exit Sub
  ''''End If
  '
  'DOES THE USER REALLY WANT TO IMPORT AT THIS TEMPERATURE AND PRESSURE?
      '---- I'VE DECIDED TO SKIP THIS STEP.  THE USER BEWARE.
  '
  '
  'DETERMINE WHICH COMPONENTS ARE IMPORTABLE.
  '
  For i = 1 To link_ChemCount
    link_AllPropsAvailable(i) = True
    If (Not link_ChemPropAvailable(PROP_HENRYSCONSTANT, i)) Then link_AllPropsAvailable(i) = False
    If (Not link_ChemPropAvailable(PROP_MOLECULARWEIGHT, i)) Then link_AllPropsAvailable(i) = False
    If (Not link_ChemPropAvailable(PROP_LOGKOW, i)) Then link_AllPropsAvailable(i) = False
    If (Not link_ChemPropAvailable(PROP_LIQUIDDIFFUSIVITY, i)) Then link_AllPropsAvailable(i) = False
    If (Not link_ChemPropAvailable(PROP_GASDIFFUSIVITY, i)) Then link_AllPropsAvailable(i) = False
  Next i
  '
  'IMPORT ALL IMPORTABLE COMPONENTS.
  '
  Num_Imported = 0
  For i = 1 To link_ChemCount
    Num_Imported = Num_Imported + 1
Dim ThisVal As Double
    '
    ' HANDLE THE NAME.
    Temp_Plant.ChemicalData.ContaminantName = link_ChemName(i)
    '
    ' HANDLE PROP_HENRYSCONSTANT.
    ThisVal = -1E+20
    If (link_ChemPropAvailable(PROP_HENRYSCONSTANT, i) = True) Then
      ThisVal = link_ChemProp(PROP_HENRYSCONSTANT, i)
    End If
    Temp_Plant.ChemicalData.DataSources(6).Val_StEPP = ThisVal
    Temp_Plant.ChemicalData.DataSources(6).SourceType = DATASOURCETYPE_STEPP
    '
    ' HANDLE PROP_MOLECULARWEIGHT.
    ThisVal = -1E+20
    If (link_ChemPropAvailable(PROP_MOLECULARWEIGHT, i) = True) Then
      ThisVal = link_ChemProp(PROP_MOLECULARWEIGHT, i)
    End If
    Temp_Plant.ChemicalData.DataSources(7).Val_StEPP = ThisVal
    Temp_Plant.ChemicalData.DataSources(7).SourceType = DATASOURCETYPE_STEPP
    '
    ' HANDLE PROP_LOGKOW.
    ThisVal = -1E+20
    If (link_ChemPropAvailable(PROP_LOGKOW, i) = True) Then
      ThisVal = link_ChemProp(PROP_LOGKOW, i)
    End If
    Temp_Plant.ChemicalData.DataSources(5).Val_StEPP = ThisVal
    Temp_Plant.ChemicalData.DataSources(5).SourceType = DATASOURCETYPE_STEPP
    '
    ' HANDLE PROP_LIQUIDDIFFUSIVITY.
    ThisVal = -1E+20
    If (link_ChemPropAvailable(PROP_LIQUIDDIFFUSIVITY, i) = True) Then
      '
      ' NOTE, THE NEXT LINE CONVERTS FROM m^2/s TO cm^2/s.
      ThisVal = link_ChemProp(PROP_LIQUIDDIFFUSIVITY, i) * 10000#
    End If
    Temp_Plant.ChemicalData.DataSources(8).Val_StEPP = ThisVal
    Temp_Plant.ChemicalData.DataSources(8).SourceType = DATASOURCETYPE_STEPP
    '
    ' HANDLE PROP_GASDIFFUSIVITY.
    ThisVal = -1E+20
    If (link_ChemPropAvailable(PROP_GASDIFFUSIVITY, i) = True) Then
      '
      ' NOTE, THE NEXT LINE CONVERTS FROM m^2/s TO cm^2/s.
      ThisVal = link_ChemProp(PROP_GASDIFFUSIVITY, i) * 10000#
    End If
    Temp_Plant.ChemicalData.DataSources(9).Val_StEPP = ThisVal
    Temp_Plant.ChemicalData.DataSources(9).SourceType = DATASOURCETYPE_STEPP
    
    ''''Call SetComponentDefaults(ThisComp, -1)
    ''''ThisComp.Name = link_ChemName(i)
    ''''ThisComp.Cas = CLng(Val(link_ChemCAS(i)))
    ''''ThisComp.Vapor_Pressure = link_ChemProp(PROP_VAPORPRESSURE, i)
    ''''ThisComp.MW = link_ChemProp(PROP_MOLECULARWEIGHT, i)
    ''''ThisComp.BP = link_ChemProp(PROP_NORMALBOILINGPOINT, i)
    ''''ThisComp.Liquid_Density = link_ChemProp(PROP_LIQUIDDENSITY, i) / 1000#
    ''''ThisComp.MolarVolume = link_ChemProp(PROP_MOLARVOLUMEATNBP, i) * 1000#
    ''''ThisComp.Refractive_Index = link_ChemProp(PROP_REFRACTIVEINDEX, i)
    ''''ThisComp.Aqueous_Solubility = link_ChemProp(PROP_AQUEOUSSOLUBILITY, i)
    ''''Number_Component = Number_Component + 1
    ''''Component(Number_Component) = ThisComp
  Next i

  'DISPLAY WARNING/SUCCESS MESSAGE.
  vb3CrLf = Chr$(13) & Chr$(10)
  If (Num_Imported <> 0) Then
    msg = "Successfully imported " & Trim$(Str$(Num_Imported)) & " component"
    If (Num_Imported <> 1) Then msg = msg & "s"
    msg = msg & " from StEPP:" & vb3CrLf
    For i = 1 To link_ChemCount
      msg = msg & "    " & Trim$(link_ChemName(i)) & vb3CrLf
    Next i
    msg = msg & vb3CrLf
    msg = msg & "The properties are for a "
    msg = msg & "pressure of " & Trim$(Str$(link_pressure)) & " Pa "
    msg = msg & "and a "
    msg = msg & "temperature of " & Trim$(Str$(link_temperature)) & " degrees Celcius.  "
    msg = msg & "If you change the temperature, you should re-import from StEPP.  "
    msg = msg & "Don't forget to set the correct values of "
    msg = msg & "the other parameters "
    msg = msg & "for this component." & vb3CrLf
    If (link_AllPropsAvailable(1) = False) Then
      msg = msg & vb3CrLf
      msg = msg & "The following parameters were unavailable in the StEPP " & _
          "database, and must be entered by you before hitting OK to " & _
          "exit this window." & vb3CrLf
      msg = msg & vb3CrLf
      If (Not link_ChemPropAvailable(PROP_HENRYSCONSTANT, 1)) Then
        msg = msg & "    Henry's Constant" & vb3CrLf
      End If
      If (Not link_ChemPropAvailable(PROP_MOLECULARWEIGHT, 1)) Then
        msg = msg & "    Molecular Weight" & vb3CrLf
      End If
      If (Not link_ChemPropAvailable(PROP_LOGKOW, 1)) Then
        msg = msg & "    Log Kow" & vb3CrLf
      End If
      If (Not link_ChemPropAvailable(PROP_LIQUIDDIFFUSIVITY, 1)) Then
        msg = msg & "    Water Diffusivity" & vb3CrLf
      End If
      If (Not link_ChemPropAvailable(PROP_GASDIFFUSIVITY, 1)) Then
        msg = msg & "    Air Diffusivity" & vb3CrLf
      End If
    End If
  End If
''''  Num_Failed = link_ChemCount - Num_Imported
''''  If (Num_Failed <> 0) Then
''''    msg = msg & vb3CrLf
''''    msg = msg & "Failed to import the following component"
''''    If (Num_Failed <> 1) Then msg = msg & "s"
''''    msg = msg & ":" & vb3CrLf
''''    For i = 1 To link_ChemCount
''''      If (Not link_AllPropsAvailable(i)) Then
''''        msg = msg & "    " & Trim$(link_ChemName(i)) & vb3CrLf
''''      End If
''''    Next i
''''    msg = msg & vb3CrLf
''''    msg = msg & "Important note: In order to successfully import a component "
''''    msg = msg & "from StEPP, the following properties must be available: "
''''    msg = msg & "vapor pressure, "
''''    msg = msg & "molecular weight, "
''''    msg = msg & "normal boiling point, "
''''    msg = msg & "liquid density, "
''''    msg = msg & "molar volume at the normal boiling point, "
''''    msg = msg & "refractive index, "
''''    msg = msg & "and aqueous solubility.  "
''''    msg = msg & "To force an import to occur, you may modify the user input "
''''    msg = msg & "value of the unavailable properties from within StEPP."
''''    msg = msg & vb3CrLf
''''  End If
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


Sub Do_ImportClipboard_Old_AdDesignS_Version(Was_Aborted As Boolean)
'Dim num_lines As Integer
'Dim cliptext As String
'Dim line_in As String
'Dim r As Integer
'Dim link_pressure As Double
'Dim link_temperature As Double
'Dim link_ChemCount As Integer
'Const CHEMPROP_MIN = 0
'Const CHEMPROP_MAX = 12
'ReDim link_ChemProp(CHEMPROP_MIN To CHEMPROP_MAX, 1 To 1) As Double
'ReDim link_ChemName(1 To 1) As String
'ReDim link_ChemCAS(1 To 1) As String
'ReDim link_ChemPropAvailable(CHEMPROP_MIN To CHEMPROP_MAX, 1 To 1) As Integer
'ReDim link_IsImportable(1 To 1) As Integer
'Dim i As Integer
'Dim j As Integer
'Const PROP_VAPORPRESSURE = 0
'Const PROP_ACTIVITYCOEFFICIENT = 1
'Const PROP_HENRYSCONSTANT = 2
'Const PROP_MOLECULARWEIGHT = 3
'Const PROP_NORMALBOILINGPOINT = 4
'Const PROP_LIQUIDDENSITY = 5
'Const PROP_MOLARVOLUMEATOPT = 6
'Const PROP_MOLARVOLUMEATNBP = 7
'Const PROP_REFRACTIVEINDEX = 8
'Const PROP_AQUEOUSSOLUBILITY = 9
'Const PROP_LOGKOW = 10
'Const PROP_LIQUIDDIFFUSIVITY = 11
'Const PROP_GASDIFFUSIVITY = 12
'Dim Num_Imported As Integer
'Dim ThisComp As ComponentPropertyType
'Dim msg As String
'Dim vb3CrLf As String
'Dim Num_Failed As Integer
'
'  Was_Aborted = True
'  On Error GoTo err_Do_ImportClipboard
'  cliptext = Clipboard.GetText()
'  cliptext = Parser_RemoveCharacters(Chr$(10), cliptext)
'  num_lines = Parser_GetNumArgs(Chr$(13), cliptext)
'  r = 1
'  Call Parser_GetArg(Chr$(13), cliptext, r, line_in)
'  If (Trim$(UCase$(line_in)) <> Trim$(UCase$("1234567890:START_OF_STEPP_CLIPBOARD_EXPORT"))) Then
'    GoTo err_nonfatal_err_Do_ImportClipboard
'  End If
'  r = r + 2
'  Call Parser_GetArg(Chr$(13), cliptext, r, line_in)
'  link_pressure = CDbl(Val(line_in))
'  If (link_pressure <= 0#) Then GoTo err_nonfatal_err_Do_ImportClipboard
'  r = r + 2
'  Call Parser_GetArg(Chr$(13), cliptext, r, line_in)
'  link_temperature = CDbl(Val(line_in))
'  If (link_temperature <= 0#) Then GoTo err_nonfatal_err_Do_ImportClipboard
'  r = r + 2
'  Call Parser_GetArg(Chr$(13), cliptext, r, line_in)
'  link_ChemCount = CInt(Val(line_in))
'  If (link_ChemCount <= 0) Then GoTo err_nonfatal_err_Do_ImportClipboard
'  ReDim link_ChemProp(CHEMPROP_MIN To CHEMPROP_MAX, 1 To link_ChemCount)
'  ReDim link_ChemName(1 To link_ChemCount)
'  ReDim link_ChemCAS(1 To link_ChemCount)
'  ReDim link_ChemPropAvailable(CHEMPROP_MIN To CHEMPROP_MAX, 1 To link_ChemCount)
'  ReDim link_IsImportable(1 To link_ChemCount)
'  For i = 1 To link_ChemCount
'    For j = CHEMPROP_MIN To CHEMPROP_MAX
'      link_ChemPropAvailable(j, i) = True
'    Next j
'    r = r + 2
'    Call Parser_GetArg(Chr$(13), cliptext, r, line_in)
'    link_ChemName(i) = Trim$(UCase$(line_in))
'    If (link_ChemName(i) = "") Then GoTo err_nonfatal_err_Do_ImportClipboard
'    r = r + 2
'    Call Parser_GetArg(Chr$(13), cliptext, r, line_in)
'    link_ChemCAS(i) = Trim$(UCase$(line_in))
'    'If (link_ChemCAS(i) = "") Then GoTo err_nonfatal_err_Do_ImportClipboard
'    For j = CHEMPROP_MIN To CHEMPROP_MAX
'      r = r + 2
'      Call Parser_GetArg(Chr$(13), cliptext, r, line_in)
'      line_in = Trim$(UCase$(line_in))
'      If (Trim$(UCase$("UNAVAILABLE")) = line_in) Then
'        link_ChemPropAvailable(j, i) = False
'      Else
'        link_ChemProp(j, i) = CDbl(Val(line_in))
'      End If
'    Next j
'  Next i
'  r = r + 1
'  Call Parser_GetArg(Chr$(13), cliptext, r, line_in)
'  If (Trim$(UCase$(line_in)) <> Trim$(UCase$("1234567890:END_OF_STEPP_CLIPBOARD_EXPORT"))) Then
'    GoTo err_nonfatal_err_Do_ImportClipboard
'  End If
'
'  'ARE THERE ENOUGH EMPTY COMPONENT SLOTS REMAINING?
'  If (Number_Component + link_ChemCount > Number_Compo_Max) Then
'    Call Show_Error("Unable to import all of chemicals in file " & _
'        "because the maximum number of chemicals has been reached.")
'    'Unload Me
'    Exit Sub
'  End If
'
'  'DOES THE USER REALLY WANT TO IMPORT AT THIS TEMPERATURE AND PRESSURE?
'      '---- I'VE DECIDED TO SKIP THIS STEP.  THE USER BEWARE.
'
'  'DETERMINE WHICH COMPONENTS ARE IMPORTABLE.
'  For i = 1 To link_ChemCount
'    link_IsImportable(i) = True
'    If (Not link_ChemPropAvailable(PROP_VAPORPRESSURE, i)) Then link_IsImportable(i) = False
'    If (Not link_ChemPropAvailable(PROP_MOLECULARWEIGHT, i)) Then link_IsImportable(i) = False
'    If (Not link_ChemPropAvailable(PROP_NORMALBOILINGPOINT, i)) Then link_IsImportable(i) = False
'    If (Not link_ChemPropAvailable(PROP_LIQUIDDENSITY, i)) Then link_IsImportable(i) = False
'    If (Not link_ChemPropAvailable(PROP_MOLARVOLUMEATNBP, i)) Then link_IsImportable(i) = False
'    If (Not link_ChemPropAvailable(PROP_REFRACTIVEINDEX, i)) Then link_IsImportable(i) = False
'    If (Not link_ChemPropAvailable(PROP_AQUEOUSSOLUBILITY, i)) Then link_IsImportable(i) = False
'  Next i
'
'  'IMPORT ALL IMPORTABLE COMPONENTS.
'  Num_Imported = 0
'  For i = 1 To link_ChemCount
'    If (link_IsImportable(i)) Then
'      Num_Imported = Num_Imported + 1
'      Call SetComponentDefaults(ThisComp, -1)
'      ThisComp.Name = link_ChemName(i)
'      ThisComp.Cas = CLng(Val(link_ChemCAS(i)))
'      ThisComp.Vapor_Pressure = link_ChemProp(PROP_VAPORPRESSURE, i)
'      ThisComp.MW = link_ChemProp(PROP_MOLECULARWEIGHT, i)
'      ThisComp.BP = link_ChemProp(PROP_NORMALBOILINGPOINT, i)
'      ThisComp.Liquid_Density = link_ChemProp(PROP_LIQUIDDENSITY, i) / 1000#
'      ThisComp.MolarVolume = link_ChemProp(PROP_MOLARVOLUMEATNBP, i) * 1000#
'      ThisComp.Refractive_Index = link_ChemProp(PROP_REFRACTIVEINDEX, i)
'      ThisComp.Aqueous_Solubility = link_ChemProp(PROP_AQUEOUSSOLUBILITY, i)
'      Number_Component = Number_Component + 1
'      Component(Number_Component) = ThisComp
'
'      ''Take care of miscellaneous screen B.S.
'      'frmpfpsdm!cmdViewDimensionless.Enabled = True
'      'frmpfpsdm!cmdEditComponent.Enabled = True
'      'frmpfpsdm!cmdDeleteComponent.Enabled = True
'      'frmpfpsdm!lstComponents.AddItem thiscomp.name
'      'frmpfpsdm!cboSelectCompo.Enabled = True
'      'frmpfpsdm!cboSelectCompo.AddItem thiscomp.name
'      'If (Number_Component = Number_Compo_Max) Then
'      '  frmpfpsdm!cmdAddComponent.Enabled = False
'      'End If
'      ''Set index of the kinetic combo box to the new chemical
'      'frmpfpsdm!cboSelectCompo.ListIndex = frmpfpsdm!cboSelectCompo.ListCount - 1
'      ''Update the corresponding kinetic data displayed
'      'Call Update_Display_Kinetic
'      'If (Number_Component > 0) Then
'      '  frmpfpsdm!mnuRunItem(0).Enabled = True
'      '  frmpfpsdm!mnuRunItem(1).Enabled = True
'      '  frmpfpsdm!mnuRunItem(2).Enabled = True
'      '  frmpfpsdm!mnuOptionsItem(0).Enabled = True
'      '  frmpfpsdm!mnuOptionsItem(1).Enabled = True  'Variable Influent concentration
'      '  frmpfpsdm!mnuOptionsItem(2).Enabled = True  'Variable Effluent concentration
'      'End If
'    End If
'  Next i
'
'  'DISPLAY WARNING/SUCCESS MESSAGE.
'  vb3CrLf = Chr$(13) & Chr$(10)
'  If (Num_Imported <> 0) Then
'    msg = "Successfully imported " & Trim$(Str$(Num_Imported)) & " component"
'    If (Num_Imported <> 1) Then msg = msg & "s"
'    msg = msg & " from StEPP:" & vb3CrLf
'    For i = 1 To link_ChemCount
'      If (link_IsImportable(i)) Then
'        msg = msg & "    " & Trim$(link_ChemName(i)) & vb3CrLf
'      End If
'    Next i
'    msg = msg & "The properties are for a "
'    msg = msg & "pressure of " & Trim$(Str$(link_pressure)) & " Pa "
'    msg = msg & "and a "
'    msg = msg & "temperature of " & Trim$(Str$(link_temperature)) & " degrees Celcius." & vb3CrLf
'    msg = msg & vb3CrLf
'    msg = msg & "Don't forget to set the correct values of Freundlich K, "
'    msg = msg & "Freundlich 1/n, and initial concentration for each "
'    msg = msg & "of these components." & vb3CrLf
'  Else
'    msg = "Unable to import any components from StEPP." & vb3CrLf
'  End If
'  Num_Failed = link_ChemCount - Num_Imported
'  If (Num_Failed <> 0) Then
'    msg = msg & vb3CrLf
'    msg = msg & "Failed to import the following component"
'    If (Num_Failed <> 1) Then msg = msg & "s"
'    msg = msg & ":" & vb3CrLf
'    For i = 1 To link_ChemCount
'      If (Not link_IsImportable(i)) Then
'        msg = msg & "    " & Trim$(link_ChemName(i)) & vb3CrLf
'      End If
'    Next i
'    msg = msg & vb3CrLf
'    msg = msg & "Important note: In order to successfully import a component "
'    msg = msg & "from StEPP, the following properties must be available: "
'    msg = msg & "vapor pressure, "
'    msg = msg & "molecular weight, "
'    msg = msg & "normal boiling point, "
'    msg = msg & "liquid density, "
'    msg = msg & "molar volume at the normal boiling point, "
'    msg = msg & "refractive index, "
'    msg = msg & "and aqueous solubility.  "
'    msg = msg & "To force an import to occur, you may modify the user input "
'    msg = msg & "value of the unavailable properties from within StEPP."
'    msg = msg & vb3CrLf
'  End If
'  Call Show_Message(msg)
'  Was_Aborted = False
'
'exit_err_err_Do_ImportClipboard:
'  Exit Sub
'err_nonfatal_err_Do_ImportClipboard:
'  Call Show_Error("An error occurred during the import process.")
'  GoTo exit_err_err_Do_ImportClipboard
'err_Do_ImportClipboard:
'  Call Show_Error("An error occurred during the import process.")
'  Resume exit_err_err_Do_ImportClipboard
End Sub



