Attribute VB_Name = "DisplayMod"
'DISPLAY.BAS

'This module contains the routines for DISPLAYING properties in Visual Basic
'on the appropriate forms and in the appropriate units.


'Format of Numbers in Visual Basic - If not specified for a property, that property will be formatted to three significant figures
Global Const WATER_DENSITY_FORMAT = "0.00"
Global Const MOLECULAR_WEIGHT_FORMAT = "0.00"
Global Const REFRACTIVE_INDEX_FORMAT = "0.0000"

Sub CheckActivityCoefficient(Index As Integer, ValueToDisplayIndex As Integer, DisplayedValueOnMainScreen As Integer)

    Select Case hie.ActivityCoefficient(Index).hierarchy
       Case ACTIVITY_COEFFICIENT_UNIFAC
          If PROPAVAILABLE(ACTIVITY_COEFFICIENT_UNIFAC) Then
             ValueToDisplayIndex = ACTIVITY_COEFFICIENT_UNIFAC
             DisplayedValueOnMainScreen = True
          End If
    End Select

End Sub

Sub CheckAirDensity(Index As Integer, ValueToDisplayIndex As Integer, DisplayedValueOnMainScreen As Integer)

    Select Case hie.AirDensity(Index).hierarchy
       Case AIR_DENSITY_CORRELATION
          If PROPAVAILABLE(AIR_DENSITY_CORRELATION) Then
             ValueToDisplayIndex = AIR_DENSITY_CORRELATION
             DisplayedValueOnMainScreen = True
          End If
       Case AIR_DENSITY_INPUT
          If PROPAVAILABLE(AIR_DENSITY_INPUT) Then
             ValueToDisplayIndex = AIR_DENSITY_INPUT
             DisplayedValueOnMainScreen = True
          End If
    End Select

End Sub

Sub CheckAirViscosity(Index As Integer, ValueToDisplayIndex As Integer, DisplayedValueOnMainScreen As Integer)

    Select Case hie.AirViscosity(Index).hierarchy
       Case AIR_VISCOSITY_CORRELATION
          If PROPAVAILABLE(AIR_VISCOSITY_CORRELATION) Then
             ValueToDisplayIndex = AIR_VISCOSITY_CORRELATION
             DisplayedValueOnMainScreen = True
          End If
       Case AIR_VISCOSITY_INPUT
          If PROPAVAILABLE(AIR_VISCOSITY_INPUT) Then
             ValueToDisplayIndex = AIR_VISCOSITY_INPUT
             DisplayedValueOnMainScreen = True
          End If
    End Select

End Sub

Sub CheckAqueousSolubility(Index As Integer, ValueToDisplayIndex As Integer, DisplayedValueOnMainScreen As Integer)

    Select Case hie.AqueousSolubility(Index).hierarchy
       Case AQUEOUS_SOLUBILITY_FIT
          If PROPAVAILABLE(AQUEOUS_SOLUBILITY_FIT) Then
             ValueToDisplayIndex = AQUEOUS_SOLUBILITY_FIT
             DisplayedValueOnMainScreen = True
          End If
       Case AQUEOUS_SOLUBILITY_OPT_UNIFAC
          If PROPAVAILABLE(AQUEOUS_SOLUBILITY_OPT_UNIFAC) Then
             ValueToDisplayIndex = AQUEOUS_SOLUBILITY_OPT_UNIFAC
             DisplayedValueOnMainScreen = True
          End If
       Case AQUEOUS_SOLUBILITY_DATABASE
          If PROPAVAILABLE(AQUEOUS_SOLUBILITY_DATABASE) Then
             ValueToDisplayIndex = AQUEOUS_SOLUBILITY_DATABASE
             DisplayedValueOnMainScreen = True
          End If
       Case AQUEOUS_SOLUBILITY_DBT_UNIFAC
          If PROPAVAILABLE(AQUEOUS_SOLUBILITY_DBT_UNIFAC) Then
             ValueToDisplayIndex = AQUEOUS_SOLUBILITY_DBT_UNIFAC
             DisplayedValueOnMainScreen = True
          End If
       Case AQUEOUS_SOLUBILITY_INPUT
          If PROPAVAILABLE(AQUEOUS_SOLUBILITY_INPUT) Then
             ValueToDisplayIndex = AQUEOUS_SOLUBILITY_INPUT
             DisplayedValueOnMainScreen = True
          End If

    End Select

End Sub

Sub CheckBoilingPoint(Index As Integer, ValueToDisplayIndex As Integer, DisplayedValueOnMainScreen As Integer)

    Select Case hie.BoilingPoint(Index).hierarchy
       Case BOILING_POINT_DATABASE
          If PROPAVAILABLE(BOILING_POINT_DATABASE) Then
             ValueToDisplayIndex = BOILING_POINT_DATABASE
             DisplayedValueOnMainScreen = True
          End If
       Case BOILING_POINT_INPUT
          If PROPAVAILABLE(BOILING_POINT_INPUT) Then
             ValueToDisplayIndex = BOILING_POINT_INPUT
             DisplayedValueOnMainScreen = True
          End If
    End Select

End Sub

Sub CheckGasDiffusivity(Index As Integer, ValueToDisplayIndex As Integer, DisplayedValueOnMainScreen As Integer)

    Select Case hie.GasDiffusivity(Index).hierarchy
       Case GAS_DIFFUSIVITY_WILKELEE
          If PROPAVAILABLE(GAS_DIFFUSIVITY_WILKELEE) Then
             ValueToDisplayIndex = GAS_DIFFUSIVITY_WILKELEE
             DisplayedValueOnMainScreen = True
          End If
       Case GAS_DIFFUSIVITY_INPUT
          If PROPAVAILABLE(GAS_DIFFUSIVITY_INPUT) Then
             ValueToDisplayIndex = GAS_DIFFUSIVITY_INPUT
             DisplayedValueOnMainScreen = True
          End If
    End Select

End Sub

Sub CheckHenrysConstant(Index As Integer, ValueToDisplayIndex As Integer, DisplayedValueOnMainScreen As Integer)

    Select Case hie.HenrysConstant(Index).hierarchy
       Case HENRYS_CONSTANT_REGRESS
          If PROPAVAILABLE(HENRYS_CONSTANT_REGRESS) Then
             ValueToDisplayIndex = HENRYS_CONSTANT_REGRESS
             DisplayedValueOnMainScreen = True
          End If
       Case HENRYS_CONSTANT_FIT
          If PROPAVAILABLE(HENRYS_CONSTANT_FIT) Then
             ValueToDisplayIndex = HENRYS_CONSTANT_FIT
             DisplayedValueOnMainScreen = True
          End If
       Case HENRYS_CONSTANT_OPT_UNIFAC
          If PROPAVAILABLE(HENRYS_CONSTANT_OPT_UNIFAC) Then
             ValueToDisplayIndex = HENRYS_CONSTANT_OPT_UNIFAC
             DisplayedValueOnMainScreen = True
          End If
       Case HENRYS_CONSTANT_DATABASE
          If PROPAVAILABLE(HENRYS_CONSTANT_DATABASE) Then
             ValueToDisplayIndex = HENRYS_CONSTANT_DATABASE
             DisplayedValueOnMainScreen = True
          End If
       Case HENRYS_CONSTANT_UNIFAC
          If PROPAVAILABLE(HENRYS_CONSTANT_UNIFAC) Then
             ValueToDisplayIndex = HENRYS_CONSTANT_UNIFAC
             DisplayedValueOnMainScreen = True
          End If
       Case HENRYS_CONSTANT_INPUT
          If PROPAVAILABLE(HENRYS_CONSTANT_INPUT) Then
             ValueToDisplayIndex = HENRYS_CONSTANT_INPUT
             DisplayedValueOnMainScreen = True
          End If

    End Select

End Sub

Sub CheckLiquidDensity(Index As Integer, ValueToDisplayIndex As Integer, DisplayedValueOnMainScreen As Integer)

    Select Case hie.LiquidDensity(Index).hierarchy
       Case LIQUID_DENSITY_DATABASE
          If PROPAVAILABLE(LIQUID_DENSITY_DATABASE) Then
             ValueToDisplayIndex = LIQUID_DENSITY_DATABASE
             DisplayedValueOnMainScreen = True
          End If
       Case LIQUID_DENSITY_UNIFAC
          If PROPAVAILABLE(LIQUID_DENSITY_UNIFAC) Then
             ValueToDisplayIndex = LIQUID_DENSITY_UNIFAC
             DisplayedValueOnMainScreen = True
          End If
       Case LIQUID_DENSITY_INPUT
          If PROPAVAILABLE(LIQUID_DENSITY_INPUT) Then
             ValueToDisplayIndex = LIQUID_DENSITY_INPUT
             DisplayedValueOnMainScreen = True
          End If
    End Select

End Sub

Sub CheckLiquidDiffusivity(HierarchyChoice As HierarchyType, ValueToDisplayIndex As Integer, DisplayedValueOnMainScreen As Integer)
    
    Select Case HierarchyChoice.hierarchy
       Case LIQUID_DIFFUSIVITY_HAYDUKLAUDIE
          If PROPAVAILABLE(LIQUID_DIFFUSIVITY_HAYDUKLAUDIE) Then
             ValueToDisplayIndex = LIQUID_DIFFUSIVITY_HAYDUKLAUDIE
             DisplayedValueOnMainScreen = True
          End If
       Case LIQUID_DIFFUSIVITY_WILKECHANG
          If PROPAVAILABLE(LIQUID_DIFFUSIVITY_WILKECHANG) Then
             ValueToDisplayIndex = LIQUID_DIFFUSIVITY_WILKECHANG
             DisplayedValueOnMainScreen = True
          End If
       Case LIQUID_DIFFUSIVITY_POLSON
          If PROPAVAILABLE(LIQUID_DIFFUSIVITY_POLSON) Then
             ValueToDisplayIndex = LIQUID_DIFFUSIVITY_POLSON
             DisplayedValueOnMainScreen = True
          End If
       Case LIQUID_DIFFUSIVITY_INPUT
          If PROPAVAILABLE(OCT_WATER_PART_COEFF_INPUT) Then
             ValueToDisplayIndex = OCT_WATER_PART_COEFF_INPUT
             DisplayedValueOnMainScreen = True
          End If

    End Select

End Sub

Sub CheckMolarVolumeNBP(Index As Integer, ValueToDisplayIndex As Integer, DisplayedValueOnMainScreen As Integer)

    Select Case hie.MolarVolumeBoilingPoint(Index).hierarchy
       Case MOLAR_VOLUME_NBP_UNIFAC
          If PROPAVAILABLE(MOLAR_VOLUME_NBP_UNIFAC) Then
             ValueToDisplayIndex = MOLAR_VOLUME_NBP_UNIFAC
             DisplayedValueOnMainScreen = True
          End If
       Case MOLAR_VOLUME_NBP_INPUT
          If PROPAVAILABLE(MOLAR_VOLUME_NBP_INPUT) Then
             ValueToDisplayIndex = MOLAR_VOLUME_NBP_INPUT
             DisplayedValueOnMainScreen = True
          End If
    End Select

End Sub

Sub CheckMolarVolumeOpT(Index As Integer, ValueToDisplayIndex As Integer, DisplayedValueOnMainScreen As Integer)

    Select Case hie.MolarVolumeOperatingT(Index).hierarchy
       Case MOLAR_VOLUME_OPT_DATABASE
          If PROPAVAILABLE(MOLAR_VOLUME_OPT_DATABASE) Then
             ValueToDisplayIndex = MOLAR_VOLUME_OPT_DATABASE
             DisplayedValueOnMainScreen = True
          End If
       Case MOLAR_VOLUME_OPT_UNIFAC
          If PROPAVAILABLE(MOLAR_VOLUME_OPT_UNIFAC) Then
             ValueToDisplayIndex = MOLAR_VOLUME_OPT_UNIFAC
             DisplayedValueOnMainScreen = True
          End If
       Case MOLAR_VOLUME_OPT_INPUT
          If PROPAVAILABLE(MOLAR_VOLUME_OPT_INPUT) Then
             ValueToDisplayIndex = MOLAR_VOLUME_OPT_INPUT
             DisplayedValueOnMainScreen = True
          End If
    End Select

End Sub

Sub CheckMolecularWeight(Index As Integer, ValueToDisplayIndex As Integer, DisplayedValueOnMainScreen As Integer)

    Select Case hie.MolecularWeight(Index).hierarchy
       Case MOLECULAR_WEIGHT_DATABASE
          If PROPAVAILABLE(MOLECULAR_WEIGHT_DATABASE) Then
             ValueToDisplayIndex = MOLECULAR_WEIGHT_DATABASE
             DisplayedValueOnMainScreen = True
          End If
       Case MOLECULAR_WEIGHT_UNIFAC
          If PROPAVAILABLE(MOLECULAR_WEIGHT_UNIFAC) Then
             ValueToDisplayIndex = MOLECULAR_WEIGHT_UNIFAC
             DisplayedValueOnMainScreen = True
          End If
       Case MOLECULAR_WEIGHT_INPUT
          If PROPAVAILABLE(MOLECULAR_WEIGHT_INPUT) Then
             ValueToDisplayIndex = MOLECULAR_WEIGHT_INPUT
             DisplayedValueOnMainScreen = True
          End If
    End Select

End Sub

Sub CheckOctWaterPartCoeff(Index As Integer, ValueToDisplayIndex As Integer, DisplayedValueOnMainScreen As Integer)

    Select Case hie.OctWaterPartCoeff(Index).hierarchy
       Case OCT_WATER_PART_COEFF_OPT_UNIFAC
          If PROPAVAILABLE(OCT_WATER_PART_COEFF_OPT_UNIFAC) Then
             ValueToDisplayIndex = OCT_WATER_PART_COEFF_OPT_UNIFAC
             DisplayedValueOnMainScreen = True
          End If
       Case OCT_WATER_PART_COEFF_DB
          If PROPAVAILABLE(OCT_WATER_PART_COEFF_DB) Then
             ValueToDisplayIndex = OCT_WATER_PART_COEFF_DB
             DisplayedValueOnMainScreen = True
          End If
       Case OCT_WATER_PART_COEFF_DBT_UNIFAC
          If PROPAVAILABLE(OCT_WATER_PART_COEFF_DBT_UNIFAC) Then
             ValueToDisplayIndex = OCT_WATER_PART_COEFF_DBT_UNIFAC
             DisplayedValueOnMainScreen = True
          End If
       Case OCT_WATER_PART_COEFF_INPUT
          If PROPAVAILABLE(OCT_WATER_PART_COEFF_INPUT) Then
             ValueToDisplayIndex = OCT_WATER_PART_COEFF_INPUT
             DisplayedValueOnMainScreen = True
          End If

    End Select

End Sub

Sub CheckRefractiveIndex(Index As Integer, ValueToDisplayIndex As Integer, DisplayedValueOnMainScreen As Integer)

    Select Case hie.RefractiveIndex(Index).hierarchy
       Case REFRACTIVE_INDEX_DATABASE
          If PROPAVAILABLE(REFRACTIVE_INDEX_DATABASE) Then
             ValueToDisplayIndex = REFRACTIVE_INDEX_DATABASE
             DisplayedValueOnMainScreen = True
          End If
       Case REFRACTIVE_INDEX_INPUT
          If PROPAVAILABLE(REFRACTIVE_INDEX_INPUT) Then
             ValueToDisplayIndex = REFRACTIVE_INDEX_INPUT
             DisplayedValueOnMainScreen = True
          End If
    End Select

End Sub

Sub CheckVaporPressure(Index As Integer, ValueToDisplayIndex As Integer, DisplayedValueOnMainScreen As Integer)

    Select Case hie.VaporPressure(Index).hierarchy
       Case VAPOR_PRESSURE_DATABASE
          If PROPAVAILABLE(VAPOR_PRESSURE_DATABASE) Then
             ValueToDisplayIndex = VAPOR_PRESSURE_DATABASE
             DisplayedValueOnMainScreen = True
          End If
       Case VAPOR_PRESSURE_INPUT
          If PROPAVAILABLE(VAPOR_PRESSURE_INPUT) Then
             ValueToDisplayIndex = VAPOR_PRESSURE_INPUT
             DisplayedValueOnMainScreen = True
          End If
    End Select

End Sub

Sub CheckWaterDensity(Index As Integer, ValueToDisplayIndex As Integer, DisplayedValueOnMainScreen As Integer)

    Select Case hie.WaterDensity(Index).hierarchy
       Case WATER_DENSITY_CORRELATION
          If PROPAVAILABLE(WATER_DENSITY_CORRELATION) Then
             ValueToDisplayIndex = WATER_DENSITY_CORRELATION
             DisplayedValueOnMainScreen = True
          End If
       Case WATER_DENSITY_INPUT
          If PROPAVAILABLE(WATER_DENSITY_INPUT) Then
             ValueToDisplayIndex = WATER_DENSITY_INPUT
             DisplayedValueOnMainScreen = True
          End If
    End Select

End Sub

Sub CheckWaterSurfaceTension(Index As Integer, ValueToDisplayIndex As Integer, DisplayedValueOnMainScreen As Integer)

    Select Case hie.WaterSurfaceTension(Index).hierarchy
       Case WATER_SURF_TENSION_CORRELATION
          If PROPAVAILABLE(WATER_SURF_TENSION_CORRELATION) Then
             ValueToDisplayIndex = WATER_SURF_TENSION_CORRELATION
             DisplayedValueOnMainScreen = True
          End If
       Case WATER_SURF_TENSION_INPUT
          If PROPAVAILABLE(WATER_SURF_TENSION_INPUT) Then
             ValueToDisplayIndex = WATER_SURF_TENSION_INPUT
             DisplayedValueOnMainScreen = True
          End If
    End Select

End Sub

Sub CheckWaterViscosity(Index As Integer, ValueToDisplayIndex As Integer, DisplayedValueOnMainScreen As Integer)

    Select Case hie.WaterViscosity(Index).hierarchy
       Case WATER_VISCOSITY_CORRELATION
          If PROPAVAILABLE(WATER_VISCOSITY_CORRELATION) Then
             ValueToDisplayIndex = WATER_VISCOSITY_CORRELATION
             DisplayedValueOnMainScreen = True
          End If
       Case WATER_VISCOSITY_INPUT
          If PROPAVAILABLE(WATER_VISCOSITY_INPUT) Then
             ValueToDisplayIndex = WATER_VISCOSITY_INPUT
             DisplayedValueOnMainScreen = True
          End If
    End Select

End Sub

Sub DisplayActivityCoefficient()
    Dim ValueToDisplayIndex As Integer
    Dim I As Integer
    Dim PropertySourceToHighlight As Integer
    Dim SIValue As Double
    Dim EnglishValue As Double
    Dim ValueToDisplay As Double

' ***** BEGIN Calculations to determine which value is displayed on
' ***** main screen according to hierarchy

    infinite_dilution_form!lblCurrentValues(0).Caption = ""
    infinite_dilution_form!lblCurrentValues(1).Caption = ""

    Call DisplayActivityCoefficientMainScreen(ValueToDisplayIndex)

' ***** END Calculations to determine which value is displayed on
' ***** main screen according to hierarchy


' ***** BEGIN Displaying Activity Coefficient Values in activity coefficient
' ***** form (Infinite_dilution_form)

'   *** Initialize all text and label boxes on Infinite_dilution_form
'   *** to gray and empty and disabled
    For I = 0 To 0
        infinite_dilution_form!Option1(I + 1).BackColor = &HC0C0C0
        infinite_dilution_form!Option1(I + 1).Enabled = False
        infinite_dilution_form!Option1(I + 1).Value = False
        infinite_dilution_form!lblSourceLabel(I).BackColor = &HC0C0C0
        infinite_dilution_form!lblActivityCoefficientValue(I).Caption = "Not Available"
'        Infinite_dilution_form!lblActivityCoefficientValue(I).Enabled = False
        infinite_dilution_form!lblActivityCoefficientValue(I).BackColor = &HC0C0C0
        infinite_dilution_form!lblACTemperature(I).Caption = ""
        infinite_dilution_form!lblACTemperature(I).Enabled = False
        infinite_dilution_form!lblACTemperature(I).BackColor = &HC0C0C0
    Next I

    If PROPAVAILABLE(ACTIVITY_COEFFICIENT_UNIFAC) Then

       SIValue = phprop.ActivityCoefficient.UNIFAC.Value
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call ACCONV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       infinite_dilution_form!lblActivityCoefficientValue(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       SIValue = phprop.ActivityCoefficient.UNIFAC.temperature
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call TEMPCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       infinite_dilution_form!lblACTemperature(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       '*** Set colors of available choices to white
       infinite_dilution_form!Option1(1).BackColor = &HFFFFFF
       infinite_dilution_form!Option1(1).Enabled = True
       infinite_dilution_form!lblSourceLabel(0).BackColor = &HFFFFFF
       infinite_dilution_form!lblActivityCoefficientValue(0).Enabled = True
       infinite_dilution_form!lblActivityCoefficientValue(0).BackColor = &HFFFFFF
       infinite_dilution_form!lblACTemperature(0).Enabled = True
       infinite_dilution_form!lblACTemperature(0).BackColor = &HFFFFFF
             
       If ValueToDisplayIndex = ACTIVITY_COEFFICIENT_UNIFAC Then
          infinite_dilution_form!Option1(1).Value = True
       Else
          infinite_dilution_form!Option1(1).Value = False
       End If
    End If

       For I = 0 To 0
           infinite_dilution_form!lblSourceLabel(I).ForeColor = &H80000008
       Next I

    '*** Highlight selected property source
    If ValueToDisplayIndex <> 0 Then
       infinite_dilution_form!lblSourceLabel(PropertySourceToHighlight).BackColor = &H800000
       infinite_dilution_form!lblSourceLabel(PropertySourceToHighlight).ForeColor = &H80000005
       hilight.ActivityCoefficient.PreviousIndex = PropertySourceToHighlight
    End If

' ***** END Displaying Activity Coefficient Values in activity coefficient
' ***** form (Infinite_dilution_form)


End Sub

Sub DisplayActivityCoefficientMainScreen(ValueToDisplayIndex As Integer)
    Dim DisplayedValueOnMainScreen As Integer
    Dim ValueToDisplay As Double
    Dim SourceOfValueToDisplay As Long
    Dim SIValue As Double
    Dim EnglishValue As Double

    If phprop.ActivityCoefficient.CurrentSelection.choice = 0 Then
       DisplayedValueOnMainScreen = False
       ValueToDisplayIndex = 0
    Else
       DisplayedValueOnMainScreen = True
       ValueToDisplayIndex = phprop.ActivityCoefficient.CurrentSelection.choice
    End If

    If Not DisplayedValueOnMainScreen Then
       Call CheckActivityCoefficient(1, ValueToDisplayIndex, DisplayedValueOnMainScreen)
    End If

    If Not DisplayedValueOnMainScreen Then
       contam_prop_form!lblContaminantProperties(1).Caption = "Not Available"
       HaveProperty(ACTIVITY_COEFFICIENT) = False
    Else
       Select Case ValueToDisplayIndex
          Case ACTIVITY_COEFFICIENT_UNIFAC

             SIValue = phprop.ActivityCoefficient.UNIFAC.Value
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call ACCONV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If

             SourceOfValueToDisplay = phprop.ActivityCoefficient.UNIFAC.source.short
             infinite_dilution_form!lblCurrentValues(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
             infinite_dilution_form!lblCurrentValues(1).Caption = infinite_dilution_form!lblSourceLabel(0).Caption
       End Select
       HaveProperty(ACTIVITY_COEFFICIENT) = True
       phprop.ActivityCoefficient.CurrentSelection.choice = ValueToDisplayIndex
       phprop.ActivityCoefficient.CurrentSelection.Value = SIValue
       phprop.ActivityCoefficient.CurrentSelection.source = SourceOfValueToDisplay

       contam_prop_form!lblContaminantProperties(1).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
    End If

End Sub

Sub DisplayAirDensity()
    Dim ValueToDisplayIndex As Integer
    Dim I As Integer
    Dim PropertySourceToHighlight As Integer
    Dim EnglishValue As Double, SIValue As Double
    Dim ValueToDisplay As Double

' ***** BEGIN Calculations to determine which value is displayed on
' ***** main screen according to hierarchy

    frmAirDensity!lblCurrentValues(0).Caption = ""
    frmAirDensity!lblCurrentValues(1).Caption = ""

    Call DisplayAirDensityMainScreen(ValueToDisplayIndex)

' ***** END Calculations to determine which value is displayed on
' ***** main screen according to hierarchy


' ***** BEGIN Displaying air density values in air density
' ***** form (frmAirDensity)

'   *** Initialize all text and label boxes on frmAirDensity to gray and empty
    For I = 0 To 0
        frmAirDensity!Option1(I + 1).BackColor = &HC0C0C0
        frmAirDensity!Option1(I + 1).Enabled = False
        frmAirDensity!Option1(I + 1).Value = False
        frmAirDensity!lblSource(I).BackColor = &HC0C0C0
        frmAirDensity!lblAirDensityValue(I).Caption = "Not Available"
'        frmAirDensity!lblAirDensityValue(I).Enabled = False
        frmAirDensity!lblAirDensityValue(I).BackColor = &HC0C0C0
        frmAirDensity!lblAirDensityTemperature(I).Caption = ""
        frmAirDensity!lblAirDensityTemperature(I).Enabled = False
        frmAirDensity!lblAirDensityTemperature(I).BackColor = &HC0C0C0
        frmAirDensity!lblAirDensityminimumT(I).Caption = ""
        frmAirDensity!lblAirDensityminimumT(I).Enabled = False
        frmAirDensity!lblAirDensityminimumT(I).BackColor = &HC0C0C0
        frmAirDensity!lblAirDensitymaximumT(I).Caption = ""
        frmAirDensity!lblAirDensitymaximumT(I).Enabled = False
        frmAirDensity!lblAirDensitymaximumT(I).BackColor = &HC0C0C0
    Next I

        frmAirDensity!Option1(2).BackColor = &HC0C0C0
        frmAirDensity!Option1(2).Enabled = False
        frmAirDensity!Option1(2).Value = False
        frmAirDensity!lblSource(1).BackColor = &HC0C0C0
        frmAirDensity!txtAirDensityValue(1).Text = ""
        frmAirDensity!txtAirDensityValue(1).Enabled = False
        frmAirDensity!txtAirDensityValue(1).BackColor = &HC0C0C0
        frmAirDensity!txtAirDensityTemperature(1).Text = ""
        frmAirDensity!txtAirDensityTemperature(1).Enabled = False
        frmAirDensity!txtAirDensityTemperature(1).BackColor = &HC0C0C0
        frmAirDensity!txtAirDensityminimumT(1).Text = ""
        frmAirDensity!txtAirDensityminimumT(1).Enabled = False
        frmAirDensity!txtAirDensityminimumT(1).BackColor = &HC0C0C0
        frmAirDensity!txtAirDensitymaximumT(1).Text = ""
        frmAirDensity!txtAirDensitymaximumT(1).Enabled = False
        frmAirDensity!txtAirDensitymaximumT(1).BackColor = &HC0C0C0

    If PROPAVAILABLE(AIR_DENSITY_CORRELATION) Then

       SIValue = phprop.AirDensity.correlation.Value
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call ADENSCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       frmAirDensity!lblAirDensityValue(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       SIValue = phprop.AirDensity.correlation.temperature
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call TEMPCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       frmAirDensity!lblAirDensityTemperature(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       frmAirDensity!lblAirDensityminimumT(0).Caption = "N/A"
       frmAirDensity!lblAirDensitymaximumT(0).Caption = "N/A"
       '*** Set colors of available choices to white
       frmAirDensity!Option1(1).BackColor = &HFFFFFF
       frmAirDensity!Option1(1).Enabled = True
       frmAirDensity!lblSource(0).BackColor = &HFFFFFF
       frmAirDensity!lblAirDensityValue(0).Enabled = True
       frmAirDensity!lblAirDensityValue(0).BackColor = &HFFFFFF
       frmAirDensity!lblAirDensityTemperature(0).Enabled = True
       frmAirDensity!lblAirDensityTemperature(0).BackColor = &HFFFFFF
       frmAirDensity!lblAirDensityminimumT(0).Enabled = True
       frmAirDensity!lblAirDensityminimumT(0).BackColor = &HFFFFFF
       frmAirDensity!lblAirDensitymaximumT(0).Enabled = True
       frmAirDensity!lblAirDensitymaximumT(0).BackColor = &HFFFFFF
       
       If ValueToDisplayIndex = AIR_DENSITY_CORRELATION Then
          frmAirDensity!Option1(1).Value = True
          PropertySourceToHighlight = 0
       Else
          frmAirDensity!Option1(1).Value = False
       End If
    End If

'  *** User input always possible so set backcolor to white
    frmAirDensity!Option1(2).BackColor = &HFFFFFF
    frmAirDensity!Option1(2).Enabled = True
    frmAirDensity!lblSource(1).BackColor = &HFFFFFF
    frmAirDensity!txtAirDensityValue(1).Enabled = True
    frmAirDensity!txtAirDensityValue(1).BackColor = &HFFFFFF
    frmAirDensity!txtAirDensityTemperature(1).Enabled = True
    frmAirDensity!txtAirDensityTemperature(1).BackColor = &HFFFFFF
    frmAirDensity!txtAirDensityminimumT(1).Enabled = True
    frmAirDensity!txtAirDensityminimumT(1).BackColor = &HFFFFFF
    frmAirDensity!txtAirDensitymaximumT(1).Enabled = True
    frmAirDensity!txtAirDensitymaximumT(1).BackColor = &HFFFFFF

    If PROPAVAILABLE(AIR_DENSITY_INPUT) Then

       SIValue = phprop.AirDensity.input.Value
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call ADENSCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       frmAirDensity!txtAirDensityValue(1).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       If HaveTemp(phprop.AirDensity.input.temperature) Then
          SIValue = phprop.AirDensity.input.temperature
          If CurrentUnits = SIUnits Then
             ValueToDisplay = SIValue
          ElseIf CurrentUnits = EnglishUnits Then
             Call TEMPCNV(EnglishValue, SIValue)
             ValueToDisplay = EnglishValue
          End If
          frmAirDensity!txtAirDensityTemperature(1).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
       Else
          frmAirDensity!txtAirDensityTemperature(1).Text = ""
       End If

       frmAirDensity!txtAirDensityminimumT(1).Text = ""
       frmAirDensity!txtAirDensitymaximumT(1).Text = ""

       If ValueToDisplayIndex = AIR_DENSITY_INPUT Then
          frmAirDensity!Option1(2).Value = True
          PropertySourceToHighlight = 1
       Else
          frmAirDensity!Option1(2).Value = False
       End If

    End If

       For I = 0 To 1
           frmAirDensity!lblSource(I).ForeColor = &H80000008
       Next I

    '*** Highlight selected property source
    If ValueToDisplayIndex <> 0 Then
       frmAirDensity!lblSource(PropertySourceToHighlight).BackColor = &H800000
       frmAirDensity!lblSource(PropertySourceToHighlight).ForeColor = &H80000005
       hilight.AirDensity.PreviousIndex = PropertySourceToHighlight
    End If

' ***** END Displaying air density values in air density
' ***** form (frmAirDensity)


End Sub

Sub DisplayAirDensityMainScreen(ValueToDisplayIndex As Integer)
    Dim DisplayedValueOnMainScreen As Integer
    Dim ValueToDisplay As Double
    Dim SourceOfValueToDisplay As Long
    Dim SIValue As Double, EnglishValue As Double

    If phprop.AirDensity.CurrentSelection.choice = 0 Then
       DisplayedValueOnMainScreen = False
       ValueToDisplayIndex = 0
    Else
       DisplayedValueOnMainScreen = True
       ValueToDisplayIndex = phprop.AirDensity.CurrentSelection.choice
    End If
    
    If Not DisplayedValueOnMainScreen Then
       Call CheckAirDensity(1, ValueToDisplayIndex, DisplayedValueOnMainScreen)
    End If

    If Not DisplayedValueOnMainScreen Then
       Call CheckAirDensity(2, ValueToDisplayIndex, DisplayedValueOnMainScreen)
    End If

    If Not DisplayedValueOnMainScreen Then
       contam_prop_form!lblAirWaterProperties(3).Caption = "Not Available"
       HaveProperty(AIR_DENSITY) = False
    Else
       Select Case ValueToDisplayIndex

          Case AIR_DENSITY_CORRELATION

             SIValue = phprop.AirDensity.correlation.Value
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call ADENSCNV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If
            
             SourceOfValueToDisplay = phprop.AirDensity.correlation.source.short
             frmAirDensity!lblCurrentValues(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
             frmAirDensity!lblCurrentValues(1).Caption = frmAirDensity!lblSource(0).Caption

          Case AIR_DENSITY_INPUT

             SIValue = phprop.AirDensity.input.Value
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call ADENSCNV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If
            
             SourceOfValueToDisplay = phprop.AirDensity.input.source.short
             frmAirDensity!lblCurrentValues(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
             frmAirDensity!lblCurrentValues(1).Caption = frmAirDensity!lblSource(1).Caption
       End Select

       HaveProperty(AIR_DENSITY) = True
       phprop.AirDensity.CurrentSelection.choice = ValueToDisplayIndex
       phprop.AirDensity.CurrentSelection.Value = SIValue
       phprop.AirDensity.CurrentSelection.source = SourceOfValueToDisplay

       contam_prop_form!lblAirWaterProperties(3).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
    End If

End Sub

Sub DisplayAirViscosity()
    Dim ValueToDisplayIndex As Integer
    Dim I As Integer
    Dim PropertySourceToHighlight As Integer
    Dim SIValue As Double, EnglishValue As Double
    Dim ValueToDisplay As Double


' ***** BEGIN Calculations to determine which value is displayed on
' ***** main screen according to hierarchy

    frmAirViscosity!lblCurrentValues(0).Caption = ""
    frmAirViscosity!lblCurrentValues(1).Caption = ""

    Call DisplayAirViscosityMainScreen(ValueToDisplayIndex)

' ***** END Calculations to determine which value is displayed on
' ***** main screen according to hierarchy


' ***** BEGIN Displaying air viscosity values in air viscosity
' ***** form (frmAirViscosity)

'   *** Initialize all text and label boxes on frmAirViscosity to gray and empty
    For I = 0 To 0
        frmAirViscosity!Option1(I + 1).BackColor = &HC0C0C0
        frmAirViscosity!Option1(I + 1).Enabled = False
        frmAirViscosity!Option1(I + 1).Value = False
        frmAirViscosity!lblSource(I).BackColor = &HC0C0C0
        frmAirViscosity!lblAirViscosityValue(I).Caption = "Not Available"
'        frmAirViscosity!lblAirViscosityValue(I).Enabled = False
        frmAirViscosity!lblAirViscosityValue(I).BackColor = &HC0C0C0
        frmAirViscosity!lblAirViscosityTemperature(I).Caption = ""
        frmAirViscosity!lblAirViscosityTemperature(I).Enabled = False
        frmAirViscosity!lblAirViscosityTemperature(I).BackColor = &HC0C0C0
        frmAirViscosity!lblAirViscosityminimumT(I).Caption = ""
        frmAirViscosity!lblAirViscosityminimumT(I).Enabled = False
        frmAirViscosity!lblAirViscosityminimumT(I).BackColor = &HC0C0C0
        frmAirViscosity!lblAirViscositymaximumT(I).Caption = ""
        frmAirViscosity!lblAirViscositymaximumT(I).Enabled = False
        frmAirViscosity!lblAirViscositymaximumT(I).BackColor = &HC0C0C0
    Next I

        frmAirViscosity!Option1(2).BackColor = &HC0C0C0
        frmAirViscosity!Option1(2).Enabled = False
        frmAirViscosity!Option1(2).Value = False
        frmAirViscosity!lblSource(1).BackColor = &HC0C0C0
        frmAirViscosity!txtAirViscosityValue(1).Text = ""
        frmAirViscosity!txtAirViscosityValue(1).Enabled = False
        frmAirViscosity!txtAirViscosityValue(1).BackColor = &HC0C0C0
        frmAirViscosity!txtAirViscosityTemperature(1).Text = ""
        frmAirViscosity!txtAirViscosityTemperature(1).Enabled = False
        frmAirViscosity!txtAirViscosityTemperature(1).BackColor = &HC0C0C0
        frmAirViscosity!txtAirViscosityminimumT(1).Text = ""
        frmAirViscosity!txtAirViscosityminimumT(1).Enabled = False
        frmAirViscosity!txtAirViscosityminimumT(1).BackColor = &HC0C0C0
        frmAirViscosity!txtAirViscositymaximumT(1).Text = ""
        frmAirViscosity!txtAirViscositymaximumT(1).Enabled = False
        frmAirViscosity!txtAirViscositymaximumT(1).BackColor = &HC0C0C0

    If PROPAVAILABLE(AIR_VISCOSITY_CORRELATION) Then

       SIValue = phprop.AirViscosity.correlation.Value
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call AVISCCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       frmAirViscosity!lblAirViscosityValue(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       SIValue = phprop.AirViscosity.correlation.temperature
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call TEMPCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       frmAirViscosity!lblAirViscosityTemperature(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       frmAirViscosity!lblAirViscosityminimumT(0).Caption = "N/A"
       frmAirViscosity!lblAirViscositymaximumT(0).Caption = "N/A"
       '*** Set colors of available choices to white
       frmAirViscosity!Option1(1).BackColor = &HFFFFFF
       frmAirViscosity!Option1(1).Enabled = True
       frmAirViscosity!lblSource(0).BackColor = &HFFFFFF
       frmAirViscosity!lblAirViscosityValue(0).Enabled = True
       frmAirViscosity!lblAirViscosityValue(0).BackColor = &HFFFFFF
       frmAirViscosity!lblAirViscosityTemperature(0).Enabled = True
       frmAirViscosity!lblAirViscosityTemperature(0).BackColor = &HFFFFFF
       frmAirViscosity!lblAirViscosityminimumT(0).Enabled = True
       frmAirViscosity!lblAirViscosityminimumT(0).BackColor = &HFFFFFF
       frmAirViscosity!lblAirViscositymaximumT(0).Enabled = True
       frmAirViscosity!lblAirViscositymaximumT(0).BackColor = &HFFFFFF
       
       If ValueToDisplayIndex = AIR_VISCOSITY_CORRELATION Then
          frmAirViscosity!Option1(1).Value = True
          PropertySourceToHighlight = 0
       Else
          frmAirViscosity!Option1(1).Value = False
       End If
    End If

'  *** User input always possible so set backcolor to white
    frmAirViscosity!Option1(2).BackColor = &HFFFFFF
    frmAirViscosity!Option1(2).Enabled = True
    frmAirViscosity!lblSource(1).BackColor = &HFFFFFF
    frmAirViscosity!txtAirViscosityValue(1).Enabled = True
    frmAirViscosity!txtAirViscosityValue(1).BackColor = &HFFFFFF
    frmAirViscosity!txtAirViscosityTemperature(1).Enabled = True
    frmAirViscosity!txtAirViscosityTemperature(1).BackColor = &HFFFFFF
    frmAirViscosity!txtAirViscosityminimumT(1).Enabled = True
    frmAirViscosity!txtAirViscosityminimumT(1).BackColor = &HFFFFFF
    frmAirViscosity!txtAirViscositymaximumT(1).Enabled = True
    frmAirViscosity!txtAirViscositymaximumT(1).BackColor = &HFFFFFF

    If PROPAVAILABLE(AIR_VISCOSITY_INPUT) Then

       SIValue = phprop.AirViscosity.input.Value
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call AVISCCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       frmAirViscosity!txtAirViscosityValue(1).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       If HaveTemp(phprop.AirViscosity.input.temperature) Then
          SIValue = phprop.AirViscosity.input.temperature
          If CurrentUnits = SIUnits Then
             ValueToDisplay = SIValue
          ElseIf CurrentUnits = EnglishUnits Then
             Call TEMPCNV(EnglishValue, SIValue)
             ValueToDisplay = EnglishValue
          End If
          frmAirViscosity!txtAirViscosityTemperature(1).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
       Else
          frmAirViscosity!txtAirViscosityTemperature(1).Text = ""
       End If

       frmAirViscosity!txtAirViscosityminimumT(1).Text = ""
       frmAirViscosity!txtAirViscositymaximumT(1).Text = ""

       If ValueToDisplayIndex = AIR_VISCOSITY_INPUT Then
          frmAirViscosity!Option1(2).Value = True
          PropertySourceToHighlight = 1
       Else
          frmAirViscosity!Option1(2).Value = False
       End If

    End If

       For I = 0 To 1
           frmAirViscosity!lblSource(I).ForeColor = &H80000008
       Next I

    '*** Highlight selected property source
    If ValueToDisplayIndex <> 0 Then
       frmAirViscosity!lblSource(PropertySourceToHighlight).BackColor = &H800000
       frmAirViscosity!lblSource(PropertySourceToHighlight).ForeColor = &H80000005
       hilight.AirViscosity.PreviousIndex = PropertySourceToHighlight
    End If


End Sub

Sub DisplayAirViscosityMainScreen(ValueToDisplayIndex As Integer)
    Dim DisplayedValueOnMainScreen As Integer
    Dim ValueToDisplay As Double
    Dim SourceOfValueToDisplay As Long
    Dim EnglishValue As Double, SIValue As Double

    If phprop.AirViscosity.CurrentSelection.choice = 0 Then
       DisplayedValueOnMainScreen = False
       ValueToDisplayIndex = 0
    Else
       DisplayedValueOnMainScreen = True
       ValueToDisplayIndex = phprop.AirViscosity.CurrentSelection.choice
    End If
    
    If Not DisplayedValueOnMainScreen Then
       Call CheckAirViscosity(1, ValueToDisplayIndex, DisplayedValueOnMainScreen)
    End If

    If Not DisplayedValueOnMainScreen Then
       Call CheckAirViscosity(2, ValueToDisplayIndex, DisplayedValueOnMainScreen)
    End If

    If Not DisplayedValueOnMainScreen Then
       contam_prop_form!lblAirWaterProperties(4).Caption = "Not Available"
       HaveProperty(AIR_VISCOSITY) = False
    Else
       Select Case ValueToDisplayIndex

          Case AIR_VISCOSITY_CORRELATION

             SIValue = phprop.AirViscosity.correlation.Value
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call AVISCCNV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If
             
             SourceOfValueToDisplay = phprop.AirViscosity.correlation.source.short
             frmAirViscosity!lblCurrentValues(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
             frmAirViscosity!lblCurrentValues(1).Caption = frmAirViscosity!lblSource(0).Caption

          Case AIR_VISCOSITY_INPUT

             SIValue = phprop.AirViscosity.input.Value
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call AVISCCNV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If
             
             SourceOfValueToDisplay = phprop.AirViscosity.input.source.short
             frmAirViscosity!lblCurrentValues(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
             frmAirViscosity!lblCurrentValues(1).Caption = frmAirViscosity!lblSource(1).Caption
       End Select

       HaveProperty(AIR_VISCOSITY) = True
       phprop.AirViscosity.CurrentSelection.choice = ValueToDisplayIndex
       phprop.AirViscosity.CurrentSelection.Value = SIValue
       phprop.AirViscosity.CurrentSelection.source = SourceOfValueToDisplay

       contam_prop_form!lblAirWaterProperties(4).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
    End If

End Sub

Sub DisplayAllProperties()
    Dim MolecularWeight As Double

      If HaveProperty(MOLECULAR_WEIGHT) Then
         MolecularWeight = phprop.MolecularWeight.CurrentSelection.Value
      Else
         MolecularWeight = -1#
      End If

    Call DisplayVaporPressure
    Call DisplayActivityCoefficient
    Call DisplayHenrysConstant
    Call DisplayMolecularWeight
    Call DisplayBoilingPoint
    Call DisplayLiquidDensity
    Call DisplayMolarVolumeOpT
    Call DisplayMolarVolumeNBP
    Call DisplayRefractiveIndex
    Call DisplayAqueousSolubility
    Call DisplayOctWaterPartCoeff
    Call DisplayLiquidDiffusivity(MolecularWeight)
    Call DisplayGasDiffusivity
    Call DisplayWaterDensity
    Call DisplayWaterViscosity
    Call DisplayWaterSurfaceTension
    Call DisplayAirDensity
    Call DisplayAirViscosity

End Sub

Sub DisplayAqueousSolubility()
    Dim ValueToDisplayIndex As Integer
    Dim PropertySourceToHighlight As Integer
    Dim I As Integer
    Dim SIValue As Double, EnglishValue As Double
    Dim ValueToDisplay As Double


' ***** BEGIN Calculations to determine which value is displayed on
' ***** main screen according to hierarchy

    aqsol_form!lblCurrentValues(0).Caption = ""
    aqsol_form!lblCurrentValues(1).Caption = ""

    Call DisplayAqueousSolubilityMainScreen(ValueToDisplayIndex)

' ***** END Calculations to determine which value is displayed on
' ***** main screen according to hierarchy


' ***** BEGIN Displaying Aqueous Solubility values in aqueous solubility
' ***** form (aqsol_form)

'   *** Initialize all text and label boxes on aqsol_form to gray and empty
    For I = 0 To 3
        aqsol_form!Option1(I + 1).BackColor = &HC0C0C0
        aqsol_form!Option1(I + 1).Enabled = False
        aqsol_form!Option1(I + 1).Value = False
        aqsol_form!lblSource(I).BackColor = &HC0C0C0
        aqsol_form!lblAqueousSolubilityValue(I).Caption = "Not Available"
'        aqsol_form!lblAqueousSolubilityValue(I).Enabled = False
        aqsol_form!lblAqueousSolubilityValue(I).BackColor = &HC0C0C0
        aqsol_form!lblAqSolTemperature(I).Caption = ""
        aqsol_form!lblAqSolTemperature(I).Enabled = False
        aqsol_form!lblAqSolTemperature(I).BackColor = &HC0C0C0
    Next I

        aqsol_form!Option1(5).BackColor = &HC0C0C0
        aqsol_form!Option1(5).Enabled = False
        aqsol_form!Option1(5).Value = False
        aqsol_form!lblSource(4).BackColor = &HC0C0C0
        aqsol_form!txtAqueousSolubilityValue(4).Text = ""
        aqsol_form!txtAqueousSolubilityValue(4).Enabled = False
        aqsol_form!txtAqueousSolubilityValue(4).BackColor = &HC0C0C0
        aqsol_form!txtAqSolTemperature(4).Text = ""
        aqsol_form!txtAqSolTemperature(4).Enabled = False
        aqsol_form!txtAqSolTemperature(4).BackColor = &HC0C0C0

    If PROPAVAILABLE(AQUEOUS_SOLUBILITY_FIT) Then

       SIValue = phprop.AqueousSolubility.fit.UNIFAC.Value
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call AQSCONV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       aqsol_form!lblAqueousSolubilityValue(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       SIValue = phprop.AqueousSolubility.fit.UNIFAC.temperature
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call TEMPCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       aqsol_form!lblAqSolTemperature(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       '*** Set colors of available choices to white
       aqsol_form!Option1(1).BackColor = &HFFFFFF
       aqsol_form!Option1(1).Enabled = True
       aqsol_form!lblSource(0).BackColor = &HFFFFFF
       aqsol_form!lblAqueousSolubilityValue(0).Enabled = True
       aqsol_form!lblAqueousSolubilityValue(0).BackColor = &HFFFFFF
       aqsol_form!lblAqSolTemperature(0).Enabled = True
       aqsol_form!lblAqSolTemperature(0).BackColor = &HFFFFFF
             
       If ValueToDisplayIndex = AQUEOUS_SOLUBILITY_FIT Then
          aqsol_form!Option1(1).Value = True
          PropertySourceToHighlight = 0
       Else
          aqsol_form!Option1(1).Value = False
       End If
    End If

    If PROPAVAILABLE(AQUEOUS_SOLUBILITY_OPT_UNIFAC) Then

       SIValue = phprop.AqueousSolubility.operatingT.UNIFAC.Value
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call AQSCONV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       aqsol_form!lblAqueousSolubilityValue(1).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       SIValue = phprop.AqueousSolubility.operatingT.UNIFAC.temperature
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call TEMPCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       aqsol_form!lblAqSolTemperature(1).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       '*** Set colors of available choices to white
       aqsol_form!Option1(2).BackColor = &HFFFFFF
       aqsol_form!Option1(2).Enabled = True
       aqsol_form!lblSource(1).BackColor = &HFFFFFF
       aqsol_form!lblAqueousSolubilityValue(1).Enabled = True
       aqsol_form!lblAqueousSolubilityValue(1).BackColor = &HFFFFFF
       aqsol_form!lblAqSolTemperature(1).Enabled = True
       aqsol_form!lblAqSolTemperature(1).BackColor = &HFFFFFF
             
       If ValueToDisplayIndex = AQUEOUS_SOLUBILITY_OPT_UNIFAC Then
          aqsol_form!Option1(2).Value = True
          PropertySourceToHighlight = 1
       Else
          aqsol_form!Option1(2).Value = False
       End If
    End If

    If PROPAVAILABLE(AQUEOUS_SOLUBILITY_DATABASE) Then

       SIValue = phprop.AqueousSolubility.database.Value
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call AQSCONV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       aqsol_form!lblAqueousSolubilityValue(2).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       SIValue = phprop.AqueousSolubility.database.temperature
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call TEMPCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       aqsol_form!lblAqSolTemperature(2).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       '*** Set colors of available choices to white
       aqsol_form!Option1(3).BackColor = &HFFFFFF
       aqsol_form!Option1(3).Enabled = True
       aqsol_form!lblSource(2).BackColor = &HFFFFFF
       aqsol_form!lblAqueousSolubilityValue(2).Enabled = True
       aqsol_form!lblAqueousSolubilityValue(2).BackColor = &HFFFFFF
       aqsol_form!lblAqSolTemperature(2).Enabled = True
       aqsol_form!lblAqSolTemperature(2).BackColor = &HFFFFFF
             
       If ValueToDisplayIndex = AQUEOUS_SOLUBILITY_DATABASE Then
          aqsol_form!Option1(3).Value = True
          PropertySourceToHighlight = 2
       Else
          aqsol_form!Option1(3).Value = False
       End If
    End If

    If PROPAVAILABLE(AQUEOUS_SOLUBILITY_DBT_UNIFAC) Then

       SIValue = phprop.AqueousSolubility.UNIFAC.Value
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call AQSCONV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       aqsol_form!lblAqueousSolubilityValue(3).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       SIValue = phprop.AqueousSolubility.UNIFAC.temperature
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call TEMPCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       aqsol_form!lblAqSolTemperature(3).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       '*** Set colors of available choices to white
       aqsol_form!Option1(4).BackColor = &HFFFFFF
       aqsol_form!Option1(4).Enabled = True
       aqsol_form!lblSource(3).BackColor = &HFFFFFF
       aqsol_form!lblAqueousSolubilityValue(3).Enabled = True
       aqsol_form!lblAqueousSolubilityValue(3).BackColor = &HFFFFFF
       aqsol_form!lblAqSolTemperature(3).Enabled = True
       aqsol_form!lblAqSolTemperature(3).BackColor = &HFFFFFF
             
       If ValueToDisplayIndex = AQUEOUS_SOLUBILITY_DBT_UNIFAC Then
          aqsol_form!Option1(4).Value = True
          PropertySourceToHighlight = 3
       Else
          aqsol_form!Option1(4).Value = False
       End If
    End If

'  *** User input always possible so set backcolor to white
    aqsol_form!Option1(5).BackColor = &HFFFFFF
    aqsol_form!Option1(5).Enabled = True
    aqsol_form!lblSource(4).BackColor = &HFFFFFF
    aqsol_form!txtAqueousSolubilityValue(4).Enabled = True
    aqsol_form!txtAqueousSolubilityValue(4).BackColor = &HFFFFFF
    aqsol_form!txtAqSolTemperature(4).Enabled = True
    aqsol_form!txtAqSolTemperature(4).BackColor = &HFFFFFF

    If PROPAVAILABLE(AQUEOUS_SOLUBILITY_INPUT) Then

       SIValue = phprop.AqueousSolubility.input.Value
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call AQSCONV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       aqsol_form!txtAqueousSolubilityValue(4).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       If HaveTemp(phprop.AqueousSolubility.input.temperature) Then
          SIValue = phprop.AqueousSolubility.input.temperature
          If CurrentUnits = SIUnits Then
             ValueToDisplay = SIValue
          ElseIf CurrentUnits = EnglishUnits Then
             Call TEMPCNV(EnglishValue, SIValue)
             ValueToDisplay = EnglishValue
          End If
          aqsol_form!txtAqSolTemperature(4).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
       Else
          aqsol_form!txtAqSolTemperature(4).Text = ""
       End If

       If ValueToDisplayIndex = AQUEOUS_SOLUBILITY_INPUT Then
          aqsol_form!Option1(5).Value = True
          PropertySourceToHighlight = 4
       Else
          aqsol_form!Option1(5).Value = False
       End If

    End If

       For I = 0 To 4
           aqsol_form!lblSource(I).ForeColor = &H80000008
       Next I

    '*** Highlight selected property source
    If ValueToDisplayIndex <> 0 Then
       aqsol_form!lblSource(PropertySourceToHighlight).BackColor = &H800000
       aqsol_form!lblSource(PropertySourceToHighlight).ForeColor = &H80000005
       hilight.AqueousSolubility.PreviousIndex = PropertySourceToHighlight
    End If


' ***** END Displaying Aqueous Solubility values in aqueous solubility
' ***** form (aqsol_form)

End Sub

Sub DisplayAqueousSolubilityMainScreen(ValueToDisplayIndex As Integer)
    Dim DisplayedValueOnMainScreen As Integer
    Dim ValueToDisplay As Double
    Dim SourceOfValueToDisplay As Long
    Dim EnglishValue As Double, SIValue As Double

    If phprop.AqueousSolubility.CurrentSelection.choice = 0 Then
       DisplayedValueOnMainScreen = False
       ValueToDisplayIndex = 0
    Else
       DisplayedValueOnMainScreen = True
       ValueToDisplayIndex = phprop.AqueousSolubility.CurrentSelection.choice
    End If
    
    If Not DisplayedValueOnMainScreen Then
       Call CheckAqueousSolubility(1, ValueToDisplayIndex, DisplayedValueOnMainScreen)
    End If

    If Not DisplayedValueOnMainScreen Then
       Call CheckAqueousSolubility(2, ValueToDisplayIndex, DisplayedValueOnMainScreen)
    End If

    If Not DisplayedValueOnMainScreen Then
       Call CheckAqueousSolubility(3, ValueToDisplayIndex, DisplayedValueOnMainScreen)
    End If

    If Not DisplayedValueOnMainScreen Then
       Call CheckAqueousSolubility(4, ValueToDisplayIndex, DisplayedValueOnMainScreen)
    End If

    If Not DisplayedValueOnMainScreen Then
       Call CheckAqueousSolubility(5, ValueToDisplayIndex, DisplayedValueOnMainScreen)
    End If

    If Not DisplayedValueOnMainScreen Then
       contam_prop_form!lblContaminantProperties(9).Caption = "Not Available"
       HaveProperty(AQUEOUS_SOLUBILITY) = False
    Else
       Select Case ValueToDisplayIndex

          Case AQUEOUS_SOLUBILITY_FIT

             SIValue = phprop.AqueousSolubility.fit.UNIFAC.Value
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call AQSCONV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If
             
             SourceOfValueToDisplay = phprop.AqueousSolubility.fit.UNIFAC.source.short
             aqsol_form!lblCurrentValues(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
             aqsol_form!lblCurrentValues(1).Caption = aqsol_form!lblSource(0).Caption

          Case AQUEOUS_SOLUBILITY_OPT_UNIFAC

             SIValue = phprop.AqueousSolubility.operatingT.UNIFAC.Value
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call AQSCONV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If
           
             SourceOfValueToDisplay = phprop.AqueousSolubility.operatingT.UNIFAC.source.short
             aqsol_form!lblCurrentValues(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
             aqsol_form!lblCurrentValues(1).Caption = aqsol_form!lblSource(1).Caption

          Case AQUEOUS_SOLUBILITY_DATABASE

             SIValue = phprop.AqueousSolubility.database.Value
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call AQSCONV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If
           
             SourceOfValueToDisplay = phprop.AqueousSolubility.database.source.short
             aqsol_form!lblCurrentValues(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
             aqsol_form!lblCurrentValues(1).Caption = aqsol_form!lblSource(2).Caption

          Case AQUEOUS_SOLUBILITY_DBT_UNIFAC

             SIValue = phprop.AqueousSolubility.UNIFAC.Value
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call AQSCONV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If
          
             SourceOfValueToDisplay = phprop.AqueousSolubility.UNIFAC.source.short
             aqsol_form!lblCurrentValues(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
             aqsol_form!lblCurrentValues(1).Caption = aqsol_form!lblSource(3).Caption

          Case AQUEOUS_SOLUBILITY_INPUT

             SIValue = phprop.AqueousSolubility.input.Value
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call AQSCONV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If
            
             SourceOfValueToDisplay = phprop.AqueousSolubility.input.source.short
             aqsol_form!lblCurrentValues(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
             aqsol_form!lblCurrentValues(1).Caption = aqsol_form!lblSource(4).Caption
       End Select

       HaveProperty(AQUEOUS_SOLUBILITY) = True
       phprop.AqueousSolubility.CurrentSelection.choice = ValueToDisplayIndex
       phprop.AqueousSolubility.CurrentSelection.Value = SIValue
       phprop.AqueousSolubility.CurrentSelection.source = SourceOfValueToDisplay

       contam_prop_form!lblContaminantProperties(9).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
    End If
    
End Sub

Sub DisplayBoilingPoint()
    Dim ValueToDisplayIndex As Integer
    Dim PropertySourceToHighlight As Integer
    Dim I As Integer
    Dim SIValue As Double, EnglishValue As Double
    Dim ValueToDisplay As Double


' ***** BEGIN Calculations to determine which value is displayed on
' ***** main screen according to hierarchy

    nbp_form!lblCurrentValues(0).Caption = ""
    nbp_form!lblCurrentValues(1).Caption = ""

    Call DisplayBoilingPointMainScreen(ValueToDisplayIndex)

' ***** END Calculations to determine which value is displayed on
' ***** main screen according to hierarchy


' ***** BEGIN Displaying Normal Boiling Point Values in
' ***** normal boiling point form (nbp_form)

'   *** Initialize all text and label boxes on nbp_form
'   *** to gray and empty and disabled
    For I = 0 To 0
        nbp_form!Option1(I + 1).BackColor = &HC0C0C0
        nbp_form!Option1(I + 1).Enabled = False
        nbp_form!Option1(I + 1).Value = False
        nbp_form!lblSource(I).BackColor = &HC0C0C0
        nbp_form!lblNormalBPValue(I).Caption = "Not Available"
'        nbp_form!lblNormalBPValue(I).Enabled = False
        nbp_form!lblNormalBPValue(I).BackColor = &HC0C0C0
    Next I

        nbp_form!Option1(2).BackColor = &HC0C0C0
        nbp_form!Option1(2).Enabled = False
        nbp_form!Option1(2).Value = False
        nbp_form!lblSource(1).BackColor = &HC0C0C0
        nbp_form!txtNormalBPValue(1).Text = ""
        nbp_form!txtNormalBPValue(1).Enabled = False
        nbp_form!txtNormalBPValue(1).BackColor = &HC0C0C0

    If PROPAVAILABLE(BOILING_POINT_DATABASE) Then

       SIValue = phprop.BoilingPoint.database.Value
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call NBPCONV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       nbp_form!lblNormalBPValue(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       '*** Set colors of available choices to white
       nbp_form!Option1(1).BackColor = &HFFFFFF
       nbp_form!Option1(1).Enabled = True
       nbp_form!lblSource(0).BackColor = &HFFFFFF
       nbp_form!lblNormalBPValue(0).Enabled = True
       nbp_form!lblNormalBPValue(0).BackColor = &HFFFFFF
             
       If ValueToDisplayIndex = BOILING_POINT_DATABASE Then
          nbp_form!Option1(1).Value = True
          PropertySourceToHighlight = 0
       Else
          nbp_form!Option1(1).Value = False
       End If
    End If

'  *** User input always possible so set backcolor to white
    nbp_form!Option1(2).BackColor = &HFFFFFF
    nbp_form!Option1(2).Enabled = True
    nbp_form!lblSource(1).BackColor = &HFFFFFF
    nbp_form!txtNormalBPValue(1).Enabled = True
    nbp_form!txtNormalBPValue(1).BackColor = &HFFFFFF

    If PROPAVAILABLE(BOILING_POINT_INPUT) Then

       SIValue = phprop.BoilingPoint.input.Value
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call NBPCONV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       nbp_form!txtNormalBPValue(1).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       If ValueToDisplayIndex = BOILING_POINT_INPUT Then
          nbp_form!Option1(2).Value = True
          PropertySourceToHighlight = 1
       Else
          nbp_form!Option1(2).Value = False
       End If

    End If

       For I = 0 To 1
           nbp_form!lblSource(I).ForeColor = &H80000008
       Next I

    '*** Highlight selected property source
    If ValueToDisplayIndex <> 0 Then
       nbp_form!lblSource(PropertySourceToHighlight).BackColor = &H800000
       nbp_form!lblSource(PropertySourceToHighlight).ForeColor = &H80000005
       hilight.BoilingPoint.PreviousIndex = PropertySourceToHighlight
    End If


' ***** END Displaying Normal Boiling Point Values in
' ***** normal boiling point form (nbp_form)

End Sub

Sub DisplayBoilingPointMainScreen(ValueToDisplayIndex As Integer)
    Dim DisplayedValueOnMainScreen As Integer
    Dim ValueToDisplay As Double
    Dim SourceOfValueToDisplay As Long
    Dim SIValue As Double, EnglishValue As Double
    
    If phprop.BoilingPoint.CurrentSelection.choice = 0 Then
       DisplayedValueOnMainScreen = False
       ValueToDisplayIndex = 0
    Else
       DisplayedValueOnMainScreen = True
       ValueToDisplayIndex = phprop.BoilingPoint.CurrentSelection.choice
    End If
    
    If Not DisplayedValueOnMainScreen Then
       Call CheckBoilingPoint(1, ValueToDisplayIndex, DisplayedValueOnMainScreen)
    End If

    If Not DisplayedValueOnMainScreen Then
       Call CheckBoilingPoint(2, ValueToDisplayIndex, DisplayedValueOnMainScreen)
    End If

    If Not DisplayedValueOnMainScreen Then
       contam_prop_form!lblContaminantProperties(4).Caption = "Not Available"
       HaveProperty(BOILING_POINT) = False
    Else
       Select Case ValueToDisplayIndex
          Case BOILING_POINT_DATABASE

             SIValue = phprop.BoilingPoint.database.Value
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call NBPCONV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If
           
             SourceOfValueToDisplay = phprop.BoilingPoint.database.source.short
             nbp_form!lblCurrentValues(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
             nbp_form!lblCurrentValues(1).Caption = nbp_form!lblSource(0).Caption
          Case BOILING_POINT_INPUT

             SIValue = phprop.BoilingPoint.input.Value
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call NBPCONV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If
            
             SourceOfValueToDisplay = phprop.BoilingPoint.input.source.short
             nbp_form!lblCurrentValues(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
             nbp_form!lblCurrentValues(1).Caption = nbp_form!lblSource(1).Caption
       End Select
       HaveProperty(BOILING_POINT) = True
       phprop.BoilingPoint.CurrentSelection.choice = ValueToDisplayIndex
       phprop.BoilingPoint.CurrentSelection.Value = SIValue
       phprop.BoilingPoint.CurrentSelection.source = SourceOfValueToDisplay

       contam_prop_form!lblContaminantProperties(4).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
    End If


End Sub

Sub DisplayGasDiffusivity()
    Dim ValueToDisplayIndex As Integer
    Dim I As Integer
    Dim PropertySourceToHighlight As Integer
    Dim EnglishValue As Double, SIValue As Double
    Dim ValueToDisplay As Double

' ***** BEGIN Calculations to determine which value is displayed on
' ***** main screen according to hierarchy

    gas_diff_form!lblCurrentValues(0).Caption = ""
    gas_diff_form!lblCurrentValues(1).Caption = ""

    Call DisplayGasDiffusivityMainScreen(ValueToDisplayIndex)

' ***** END Calculations to determine which value is displayed on
' ***** main screen according to hierarchy


' ***** BEGIN Displaying Gas Diffusivity values
' ***** in gas diffusivity form (gas_diff_form)

'   *** Initialize all text and label boxes on gas_diff_form to gray and empty
    For I = 0 To 0
        gas_diff_form!Option1(I + 1).BackColor = &HC0C0C0
        gas_diff_form!Option1(I + 1).Enabled = False
        gas_diff_form!Option1(I + 1).Value = False
        gas_diff_form!lblSource(I).BackColor = &HC0C0C0
        gas_diff_form!lblGasDiffusivityValue(I).Caption = "Not Available"
'        gas_diff_form!lblGasDiffusivityValue(I).Enabled = False
        gas_diff_form!lblGasDiffusivityValue(I).BackColor = &HC0C0C0
        gas_diff_form!lblGasDiffTemperature(I).Caption = ""
        gas_diff_form!lblGasDiffTemperature(I).Enabled = False
        gas_diff_form!lblGasDiffTemperature(I).BackColor = &HC0C0C0
    Next I

        gas_diff_form!Option1(2).BackColor = &HC0C0C0
        gas_diff_form!Option1(2).Enabled = False
        gas_diff_form!Option1(2).Value = False
        gas_diff_form!lblSource(1).BackColor = &HC0C0C0
        gas_diff_form!txtGasDiffusivityValue(1).Text = ""
        gas_diff_form!txtGasDiffusivityValue(1).Enabled = False
        gas_diff_form!txtGasDiffusivityValue(1).BackColor = &HC0C0C0
        gas_diff_form!txtGasDiffTemperature(1).Text = ""
        gas_diff_form!txtGasDiffTemperature(1).Enabled = False
        gas_diff_form!txtGasDiffTemperature(1).BackColor = &HC0C0C0

    If PROPAVAILABLE(GAS_DIFFUSIVITY_WILKELEE) Then

       SIValue = phprop.GasDiffusivity.wilkeLee.Value
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call GDIFFCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       gas_diff_form!lblGasDiffusivityValue(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       SIValue = phprop.GasDiffusivity.wilkeLee.temperature
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call TEMPCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       gas_diff_form!lblGasDiffTemperature(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       '*** Set colors of available choices to white
       gas_diff_form!Option1(1).BackColor = &HFFFFFF
       gas_diff_form!Option1(1).Enabled = True
       gas_diff_form!lblSource(0).BackColor = &HFFFFFF
       gas_diff_form!lblGasDiffusivityValue(0).Enabled = True
       gas_diff_form!lblGasDiffusivityValue(0).BackColor = &HFFFFFF
       gas_diff_form!lblGasDiffTemperature(0).Enabled = True
       gas_diff_form!lblGasDiffTemperature(0).BackColor = &HFFFFFF
             
       If ValueToDisplayIndex = GAS_DIFFUSIVITY_WILKELEE Then
          gas_diff_form!Option1(1).Value = True
          PropertySourceToHighlight = 0
       Else
          gas_diff_form!Option1(1).Value = False
       End If
    End If

'  *** User input always possible so set backcolor to white
    gas_diff_form!Option1(2).BackColor = &HFFFFFF
    gas_diff_form!Option1(2).Enabled = True
    gas_diff_form!lblSource(1).BackColor = &HFFFFFF
    gas_diff_form!txtGasDiffusivityValue(1).Enabled = True
    gas_diff_form!txtGasDiffusivityValue(1).BackColor = &HFFFFFF
    gas_diff_form!txtGasDiffTemperature(1).Enabled = True
    gas_diff_form!txtGasDiffTemperature(1).BackColor = &HFFFFFF

    If PROPAVAILABLE(GAS_DIFFUSIVITY_INPUT) Then

       SIValue = phprop.GasDiffusivity.input.Value
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call GDIFFCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       gas_diff_form!txtGasDiffusivityValue(1).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       If HaveTemp(phprop.GasDiffusivity.input.temperature) Then
          SIValue = phprop.GasDiffusivity.input.temperature
          If CurrentUnits = SIUnits Then
             ValueToDisplay = SIValue
          ElseIf CurrentUnits = EnglishUnits Then
             Call TEMPCNV(EnglishValue, SIValue)
             ValueToDisplay = EnglishValue
          End If
          gas_diff_form!txtGasDiffTemperature(1).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
       Else
          gas_diff_form!txtGasDiffTemperature(1).Text = ""
       End If

       If ValueToDisplayIndex = GAS_DIFFUSIVITY_INPUT Then
          gas_diff_form!Option1(2).Value = True
          PropertySourceToHighlight = 1
       Else
          gas_diff_form!Option1(2).Value = False
       End If

    End If

       For I = 0 To 1
           gas_diff_form!lblSource(I).ForeColor = &H80000008
       Next I

    '*** Highlight selected property source
    If ValueToDisplayIndex <> 0 Then
       gas_diff_form!lblSource(PropertySourceToHighlight).BackColor = &H800000
       gas_diff_form!lblSource(PropertySourceToHighlight).ForeColor = &H80000005
       hilight.GasDiffusivity.PreviousIndex = PropertySourceToHighlight
    End If


' ***** END Displaying Gas Diffusivity values
' ***** in gas diffusivity form (gas_diff_form)


End Sub

Sub DisplayGasDiffusivityMainScreen(ValueToDisplayIndex As Integer)
    Dim DisplayedValueOnMainScreen As Integer
    Dim ValueToDisplay As Double
    Dim SourceOfValueToDisplay As Long
    Dim EnglishValue As Double, SIValue As Double

    If phprop.GasDiffusivity.CurrentSelection.choice = 0 Then
       DisplayedValueOnMainScreen = False
       ValueToDisplayIndex = 0
    Else
       DisplayedValueOnMainScreen = True
       ValueToDisplayIndex = phprop.GasDiffusivity.CurrentSelection.choice
    End If
    
    If Not DisplayedValueOnMainScreen Then
       Call CheckGasDiffusivity(1, ValueToDisplayIndex, DisplayedValueOnMainScreen)
    End If

    If Not DisplayedValueOnMainScreen Then
       Call CheckGasDiffusivity(2, ValueToDisplayIndex, DisplayedValueOnMainScreen)
    End If

    If Not DisplayedValueOnMainScreen Then
       contam_prop_form!lblContaminantProperties(12).Caption = "Not Available"
       HaveProperty(GAS_DIFFUSIVITY) = False
    Else
       Select Case ValueToDisplayIndex

          Case GAS_DIFFUSIVITY_WILKELEE

             SIValue = phprop.GasDiffusivity.wilkeLee.Value
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call GDIFFCNV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If
           
             SourceOfValueToDisplay = phprop.GasDiffusivity.wilkeLee.source.short
             gas_diff_form!lblCurrentValues(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
             gas_diff_form!lblCurrentValues(1).Caption = gas_diff_form!lblSource(0).Caption

          Case GAS_DIFFUSIVITY_INPUT

             SIValue = phprop.GasDiffusivity.input.Value
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call GDIFFCNV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If
           
             SourceOfValueToDisplay = phprop.GasDiffusivity.input.source.short
             gas_diff_form!lblCurrentValues(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
             gas_diff_form!lblCurrentValues(1).Caption = gas_diff_form!lblSource(1).Caption
       End Select

       HaveProperty(GAS_DIFFUSIVITY) = True
       phprop.GasDiffusivity.CurrentSelection.choice = ValueToDisplayIndex
       phprop.GasDiffusivity.CurrentSelection.Value = SIValue
       phprop.GasDiffusivity.CurrentSelection.source = SourceOfValueToDisplay

       contam_prop_form!lblContaminantProperties(12).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
    End If

End Sub

Sub DisplayHenrysConstant()
    
    Dim ValueToDisplayIndex As Integer
    Dim PropertySourceToHighlight As Integer
    Dim I As Integer
    Dim SIValue As Double, EnglishValue As Double
    Dim ValueToDisplay As Double
    Dim hc_database_value As String * 40
    Dim hc_database_temp As String
    Dim hc_string As String
    Dim hc_unifac_value As String * 40
    Dim hc_unifac_temp As String


' ***** BEGIN Calculations to determine which value is displayed on
' ***** main screen according to hierarchy

    hc_form!lblCurrentValues(0).Caption = ""
    hc_form!lblCurrentValues(1).Caption = ""

    Call DisplayHenrysConstantMainScreen(ValueToDisplayIndex)

' ***** END Calculations to determine which value is displayed on
' ***** main screen according to hierarchy


' ***** BEGIN Displaying Henry's constant values in Henry's constant
' ***** form (hc_form)

'   *** Initialize all text and label boxes on hc_form to gray and empty
    For I = 0 To 4
        hc_form!Option1(I + 1).BackColor = &HC0C0C0
        hc_form!Option1(I + 1).Enabled = False
        hc_form!Option1(I + 1).Value = False
        hc_form!lblSource(I).BackColor = &HC0C0C0
        hc_form!lblHenrysConstantValue(I).Caption = "Not Available"
'        hc_form!lblHenrysConstantValue(I).Enabled = False
        hc_form!lblHenrysConstantValue(I).BackColor = &HC0C0C0
        hc_form!lblHCTemperature(I).Caption = ""
        hc_form!lblHCTemperature(I).Enabled = False
        hc_form!lblHCTemperature(I).BackColor = &HC0C0C0
        hc_form!lblHCminimumT(I).Caption = ""
        hc_form!lblHCminimumT(I).Enabled = False
        hc_form!lblHCminimumT(I).BackColor = &HC0C0C0
        hc_form!lblHCmaximumT(I).Caption = ""
        hc_form!lblHCmaximumT(I).Enabled = False
        hc_form!lblHCmaximumT(I).BackColor = &HC0C0C0
    Next I

        hc_form!Option1(6).BackColor = &HC0C0C0
        hc_form!Option1(6).Enabled = False
        hc_form!Option1(6).Value = False
        hc_form!lblSource(5).BackColor = &HC0C0C0
        hc_form!txtHenrysConstantValue(5).Text = ""
        hc_form!txtHenrysConstantValue(5).Enabled = False
        hc_form!txtHenrysConstantValue(5).BackColor = &HC0C0C0
        hc_form!txtHCTemperature(5).Text = ""
        hc_form!txtHCTemperature(5).Enabled = False
        hc_form!txtHCTemperature(5).BackColor = &HC0C0C0
        hc_form!txtHCminimumT(5).Text = ""
        hc_form!txtHCminimumT(5).Enabled = False
        hc_form!txtHCminimumT(5).BackColor = &HC0C0C0
        hc_form!txtHCmaximumT(5).Text = ""
        hc_form!txtHCmaximumT(5).Enabled = False
        hc_form!txtHCmaximumT(5).BackColor = &HC0C0C0
        hc_form!lblUNIFAC.BackColor = &HC0C0C0
        hc_form!cboUNIFAC.BackColor = &HC0C0C0
        hc_form!cboUNIFAC.Enabled = False
        hc_form!lblDatabase.BackColor = &HC0C0C0
        hc_form!hc_list.BackColor = &HC0C0C0
        hc_form!hc_list.Enabled = False


    If PROPAVAILABLE(HENRYS_CONSTANT_REGRESS) Then

       SIValue = phprop.HenrysConstant.regress.Value
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call HCCONV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       hc_form!lblHenrysConstantValue(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       SIValue = phprop.HenrysConstant.regress.temperature
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call TEMPCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       hc_form!lblHCTemperature(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       hc_form!lblHCminimumT(0).Caption = "N/A"
       hc_form!lblHCmaximumT(0).Caption = "N/A"
       '*** Set colors of available choices to white
       hc_form!Option1(1).BackColor = &HFFFFFF
       hc_form!Option1(1).Enabled = True
       hc_form!lblSource(0).BackColor = &HFFFFFF
       hc_form!lblHenrysConstantValue(0).Enabled = True
       hc_form!lblHenrysConstantValue(0).BackColor = &HFFFFFF
       hc_form!lblHCTemperature(0).Enabled = True
       hc_form!lblHCTemperature(0).BackColor = &HFFFFFF
       hc_form!lblHCminimumT(0).Enabled = True
       hc_form!lblHCminimumT(0).BackColor = &HFFFFFF
       hc_form!lblHCmaximumT(0).Enabled = True
       hc_form!lblHCmaximumT(0).BackColor = &HFFFFFF
       
       If ValueToDisplayIndex = HENRYS_CONSTANT_REGRESS Then
          hc_form!Option1(1).Value = True
          PropertySourceToHighlight = 0
       Else
          hc_form!Option1(1).Value = False
       End If
    End If


    If PROPAVAILABLE(HENRYS_CONSTANT_FIT) Then

       SIValue = phprop.HenrysConstant.fit.UNIFAC.Value
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call HCCONV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       hc_form!lblHenrysConstantValue(1).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       SIValue = phprop.HenrysConstant.fit.UNIFAC.temperature
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call TEMPCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       hc_form!lblHCTemperature(1).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       hc_form!lblHCminimumT(1).Caption = "N/A"
       hc_form!lblHCmaximumT(1).Caption = "N/A"
       '*** Set colors of available choices to white
       hc_form!Option1(2).BackColor = &HFFFFFF
       hc_form!Option1(2).Enabled = True
       hc_form!lblSource(1).BackColor = &HFFFFFF
       hc_form!lblHenrysConstantValue(1).Enabled = True
       hc_form!lblHenrysConstantValue(1).BackColor = &HFFFFFF
       hc_form!lblHCTemperature(1).Enabled = True
       hc_form!lblHCTemperature(1).BackColor = &HFFFFFF
       hc_form!lblHCminimumT(1).Enabled = True
       hc_form!lblHCminimumT(1).BackColor = &HFFFFFF
       hc_form!lblHCmaximumT(1).Enabled = True
       hc_form!lblHCmaximumT(1).BackColor = &HFFFFFF
       
       If ValueToDisplayIndex = HENRYS_CONSTANT_FIT Then
          hc_form!Option1(2).Value = True
          PropertySourceToHighlight = 1
       Else
          hc_form!Option1(2).Value = False
       End If
    End If

    If PROPAVAILABLE(HENRYS_CONSTANT_OPT_UNIFAC) Then

       SIValue = phprop.HenrysConstant.operatingT.UNIFAC.Value
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call HCCONV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       hc_form!lblHenrysConstantValue(2).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       SIValue = phprop.HenrysConstant.operatingT.UNIFAC.temperature
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call TEMPCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       hc_form!lblHCTemperature(2).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       hc_form!lblHCminimumT(2).Caption = "N/A"
       hc_form!lblHCmaximumT(2).Caption = "N/A"
       '*** Set colors of available choices to white
       hc_form!Option1(3).BackColor = &HFFFFFF
       hc_form!Option1(3).Enabled = True
       hc_form!lblSource(2).BackColor = &HFFFFFF
       hc_form!lblHenrysConstantValue(2).Enabled = True
       hc_form!lblHenrysConstantValue(2).BackColor = &HFFFFFF
       hc_form!lblHCTemperature(2).Enabled = True
       hc_form!lblHCTemperature(2).BackColor = &HFFFFFF
       hc_form!lblHCminimumT(2).Enabled = True
       hc_form!lblHCminimumT(2).BackColor = &HFFFFFF
       hc_form!lblHCmaximumT(2).Enabled = True
       hc_form!lblHCmaximumT(2).BackColor = &HFFFFFF
       
       If ValueToDisplayIndex = HENRYS_CONSTANT_OPT_UNIFAC Then
          hc_form!Option1(3).Value = True
          PropertySourceToHighlight = 2
       Else
          hc_form!Option1(3).Value = False
       End If
    End If

    If PROPAVAILABLE(HENRYS_CONSTANT_DATABASE) Then

       SIValue = phprop.HenrysConstant.database(phprop.HenrysConstant.chosenDatabaseIndex).Value
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call HCCONV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       hc_form!lblHenrysConstantValue(3).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       SIValue = phprop.HenrysConstant.database(phprop.HenrysConstant.chosenDatabaseIndex).temperature
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call TEMPCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       hc_form!lblHCTemperature(3).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       hc_form!lblHCminimumT(3).Caption = "N/A"
       hc_form!lblHCmaximumT(3).Caption = "N/A"
       '*** Set colors of available choices to white
       hc_form!Option1(4).BackColor = &HFFFFFF
       hc_form!Option1(4).Enabled = True
       hc_form!lblSource(3).BackColor = &HFFFFFF
       hc_form!lblHenrysConstantValue(3).Enabled = True
       hc_form!lblHenrysConstantValue(3).BackColor = &HFFFFFF
       hc_form!lblHCTemperature(3).Enabled = True
       hc_form!lblHCTemperature(3).BackColor = &HFFFFFF
       hc_form!lblHCminimumT(3).Enabled = True
       hc_form!lblHCminimumT(3).BackColor = &HFFFFFF
       hc_form!lblHCmaximumT(3).Enabled = True
       hc_form!lblHCmaximumT(3).BackColor = &HFFFFFF
       hc_form!lblDatabase.BackColor = &HFFFFFF
       hc_form!hc_list.BackColor = &HFFFFFF
       hc_form!hc_list.Enabled = True

       '***  Build combo box on HC_FORM for database values

       hc_form!hc_list.Clear

       If phprop.HenrysConstant.database(1).source.short = 3 Then   'RTI
          hc_form!lblDatabase.Caption = "RTI"
       ElseIf phprop.HenrysConstant.database(1).source.short = 1 Then   'Yaws
          hc_form!lblDatabase.Caption = "Yaws"
       ElseIf phprop.HenrysConstant.database(1).source.short = 2 Then   'Superfund
          hc_form!lblDatabase.Caption = "Superfund"
       End If

       For I = 1 To phprop.HenrysConstant.NumberOfDatabaseHenrysConstants
             
           SIValue = phprop.HenrysConstant.database(I).Value
           If CurrentUnits = SIUnits Then
              ValueToDisplay = SIValue
           ElseIf CurrentUnits = EnglishUnits Then
              Call HCCONV(EnglishValue, SIValue)
              ValueToDisplay = EnglishValue
           End If
           LSet hc_database_value = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

           SIValue = phprop.HenrysConstant.database(I).temperature
           If CurrentUnits = SIUnits Then
              ValueToDisplay = SIValue
           ElseIf CurrentUnits = EnglishUnits Then
              Call TEMPCNV(EnglishValue, SIValue)
              ValueToDisplay = EnglishValue
           End If
           hc_database_temp = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

           hc_string = hc_database_value + hc_database_temp
           hc_form!hc_list.AddItem hc_string
           If phprop.HenrysConstant.chosenDatabaseIndex = I Then
              hc_form!hc_list.ListIndex = I - 1
           End If
             
       Next I

       
       If ValueToDisplayIndex = HENRYS_CONSTANT_DATABASE Then
          hc_form!Option1(4).Value = True
          PropertySourceToHighlight = 3
       Else
          hc_form!Option1(4).Value = False
       End If
    End If

    If PROPAVAILABLE(HENRYS_CONSTANT_UNIFAC) Then

       SIValue = phprop.HenrysConstant.UNIFAC(phprop.HenrysConstant.chosenUNIFACIndex).Value
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call HCCONV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       hc_form!lblHenrysConstantValue(4).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       SIValue = phprop.HenrysConstant.UNIFAC(phprop.HenrysConstant.chosenUNIFACIndex).temperature
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call TEMPCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       hc_form!lblHCTemperature(4).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       hc_form!lblHCminimumT(4).Caption = "N/A"
       hc_form!lblHCmaximumT(4).Caption = "N/A"
       '*** Set colors of available choices to white
       hc_form!Option1(5).BackColor = &HFFFFFF
       hc_form!Option1(5).Enabled = True
       hc_form!lblSource(4).BackColor = &HFFFFFF
       hc_form!lblHenrysConstantValue(4).Enabled = True
       hc_form!lblHenrysConstantValue(4).BackColor = &HFFFFFF
       hc_form!lblHCTemperature(4).Enabled = True
       hc_form!lblHCTemperature(4).BackColor = &HFFFFFF
       hc_form!lblHCminimumT(4).Enabled = True
       hc_form!lblHCminimumT(4).BackColor = &HFFFFFF
       hc_form!lblHCmaximumT(4).Enabled = True
       hc_form!lblHCmaximumT(4).BackColor = &HFFFFFF
       hc_form!lblUNIFAC.BackColor = &HFFFFFF
       hc_form!cboUNIFAC.BackColor = &HFFFFFF
       hc_form!cboUNIFAC.Enabled = True

       hc_form!cboUNIFAC.Clear
       '*** Build combo box of UNIFAC Henry's constants on hc_form
       
       For I = 1 To phprop.HenrysConstant.NumberOfDatabaseHenrysConstants
           If phprop.HenrysConstant.UNIFAC(I).error >= 0 Then
              If hc_form!lblUNIFAC.Caption = "" Then hc_form!lblUNIFAC.Caption = "UNIFAC"

              SIValue = phprop.HenrysConstant.UNIFAC(I).Value
              If CurrentUnits = SIUnits Then
                 ValueToDisplay = SIValue
              ElseIf CurrentUnits = EnglishUnits Then
                 Call HCCONV(EnglishValue, SIValue)
                 ValueToDisplay = EnglishValue
              End If
              LSet hc_unifac_value = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

              SIValue = phprop.HenrysConstant.UNIFAC(I).temperature
              If CurrentUnits = SIUnits Then
                 ValueToDisplay = SIValue
              ElseIf CurrentUnits = EnglishUnits Then
                 Call TEMPCNV(EnglishValue, SIValue)
                 ValueToDisplay = EnglishValue
              End If
              hc_unifac_temp = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

              hc_string = hc_unifac_value + hc_unifac_temp
              hc_form!cboUNIFAC.AddItem hc_string
              If phprop.HenrysConstant.chosenUNIFACIndex = I Then
                 hc_form!cboUNIFAC.ListIndex = I - 1
              End If
           Else
              LSet hc_unifac_value = "N/A    "

              SIValue = phprop.HenrysConstant.database(I).temperature
              If CurrentUnits = SIUnits Then
                 ValueToDisplay = SIValue
              ElseIf CurrentUnits = EnglishUnits Then
                 Call TEMPCNV(EnglishValue, SIValue)
                 ValueToDisplay = EnglishValue
              End If
              hc_unifac_temp = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

              hc_string = hc_unifac_value + hc_unifac_temp
              hc_form!cboUNIFAC.AddItem hc_string
           End If
       Next I


       If ValueToDisplayIndex = HENRYS_CONSTANT_UNIFAC Then
          hc_form!Option1(5).Value = True
          PropertySourceToHighlight = 4
       Else
          hc_form!Option1(5).Value = False
       End If
    End If

'  *** User input always possible so set backcolor to white
    hc_form!Option1(6).BackColor = &HFFFFFF
    hc_form!Option1(6).Enabled = True
    hc_form!lblSource(5).BackColor = &HFFFFFF
    hc_form!txtHenrysConstantValue(5).Enabled = True
    hc_form!txtHenrysConstantValue(5).BackColor = &HFFFFFF
    hc_form!txtHCTemperature(5).Enabled = True
    hc_form!txtHCTemperature(5).BackColor = &HFFFFFF
    hc_form!txtHCminimumT(5).Enabled = True
    hc_form!txtHCminimumT(5).BackColor = &HFFFFFF
    hc_form!txtHCmaximumT(5).Enabled = True
    hc_form!txtHCmaximumT(5).BackColor = &HFFFFFF

    If PROPAVAILABLE(HENRYS_CONSTANT_INPUT) Then

       SIValue = phprop.HenrysConstant.input.Value
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call HCCONV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       hc_form!txtHenrysConstantValue(5).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       If HaveTemp(phprop.HenrysConstant.input.temperature) Then

          SIValue = phprop.HenrysConstant.input.temperature
          If CurrentUnits = SIUnits Then
             ValueToDisplay = SIValue
          ElseIf CurrentUnits = EnglishUnits Then
             Call TEMPCNV(EnglishValue, SIValue)
             ValueToDisplay = EnglishValue
          End If
          hc_form!txtHCTemperature(5).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
       Else
          hc_form!txtHCTemperature(5).Text = ""
       End If
       hc_form!txtHCminimumT(5).Text = "N/A"
       hc_form!txtHCmaximumT(5).Text = "N/A"

       If ValueToDisplayIndex = HENRYS_CONSTANT_INPUT Then
          hc_form!Option1(6).Value = True
          PropertySourceToHighlight = 5
       Else
          hc_form!Option1(5).Value = False
       End If

    End If

       For I = 0 To 5
           hc_form!lblSource(I).ForeColor = &H80000008
       Next I

    '*** Highlight selected property source
    If ValueToDisplayIndex <> 0 Then
       hc_form!lblSource(PropertySourceToHighlight).BackColor = &H800000
       hc_form!lblSource(PropertySourceToHighlight).ForeColor = &H80000005
       hilight.HenrysConstant.PreviousIndex = PropertySourceToHighlight
    End If


' ***** END Displaying Henry's constant values in Henry's constant
' ***** form (hc_form)

End Sub

Sub DisplayHenrysConstantMainScreen(ValueToDisplayIndex As Integer)
    Dim DisplayedValueOnMainScreen As Integer
    Dim ValueToDisplay As Double
    Dim HenrysConstantCloseIndex As Integer
    Dim HenrysConstantUNIFACCloseIndex As Integer
    Dim SourceOfValueToDisplay As Long
    Dim SIValue As Double, EnglishValue As Double

    If phprop.HenrysConstant.CurrentSelection.choice = 0 Then
       DisplayedValueOnMainScreen = False
       ValueToDisplayIndex = 0
    Else
       DisplayedValueOnMainScreen = True
       ValueToDisplayIndex = phprop.HenrysConstant.CurrentSelection.choice
    End If
    
    If Not DisplayedValueOnMainScreen Then
       Call CheckHenrysConstant(1, ValueToDisplayIndex, DisplayedValueOnMainScreen)
    End If

    If Not DisplayedValueOnMainScreen Then
       Call CheckHenrysConstant(2, ValueToDisplayIndex, DisplayedValueOnMainScreen)
    End If

    If Not DisplayedValueOnMainScreen Then
       Call CheckHenrysConstant(3, ValueToDisplayIndex, DisplayedValueOnMainScreen)
    End If

    If Not DisplayedValueOnMainScreen Then
       Call CheckHenrysConstant(4, ValueToDisplayIndex, DisplayedValueOnMainScreen)
    End If

    If Not DisplayedValueOnMainScreen Then
       Call CheckHenrysConstant(5, ValueToDisplayIndex, DisplayedValueOnMainScreen)
    End If

    If Not DisplayedValueOnMainScreen Then
       Call CheckHenrysConstant(6, ValueToDisplayIndex, DisplayedValueOnMainScreen)
    End If


    If Not DisplayedValueOnMainScreen Then
       contam_prop_form!lblContaminantProperties(2).Caption = "Not Available"
       HaveProperty(HENRYS_CONSTANT) = False
    Else
       Select Case ValueToDisplayIndex
          Case HENRYS_CONSTANT_REGRESS

             SIValue = phprop.HenrysConstant.regress.Value
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call HCCONV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If

             SourceOfValueToDisplay = phprop.HenrysConstant.regress.source.short
             hc_form!lblCurrentValues(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
             hc_form!lblCurrentValues(1).Caption = hc_form!lblSource(0).Caption
          Case HENRYS_CONSTANT_FIT

             SIValue = phprop.HenrysConstant.fit.UNIFAC.Value
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call HCCONV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If

             SourceOfValueToDisplay = phprop.HenrysConstant.fit.UNIFAC.source.short
             hc_form!lblCurrentValues(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
             hc_form!lblCurrentValues(1).Caption = hc_form!lblSource(1).Caption
          Case HENRYS_CONSTANT_OPT_UNIFAC

             SIValue = phprop.HenrysConstant.operatingT.UNIFAC.Value
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call HCCONV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If

             SourceOfValueToDisplay = phprop.HenrysConstant.operatingT.UNIFAC.source.short
             hc_form!lblCurrentValues(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
             hc_form!lblCurrentValues(1).Caption = hc_form!lblSource(2).Caption
          Case HENRYS_CONSTANT_DATABASE

             HenrysConstantCloseIndex = phprop.HenrysConstant.chosenDatabaseIndex
             SIValue = phprop.HenrysConstant.database(HenrysConstantCloseIndex).Value
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call HCCONV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If

             SourceOfValueToDisplay = phprop.HenrysConstant.database(HenrysConstantCloseIndex).source.short
             hc_form!lblCurrentValues(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
             hc_form!lblCurrentValues(1).Caption = hc_form!lblSource(3).Caption
          Case HENRYS_CONSTANT_UNIFAC

             HenrysConstantUNIFACCloseIndex = phprop.HenrysConstant.chosenUNIFACIndex
             SIValue = phprop.HenrysConstant.UNIFAC(HenrysConstantUNIFACCloseIndex).Value
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call HCCONV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If

             SourceOfValueToDisplay = phprop.HenrysConstant.UNIFAC(HenrysConstantUNIFACCloseIndex).source.short
             hc_form!lblCurrentValues(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
             hc_form!lblCurrentValues(1).Caption = hc_form!lblSource(4).Caption
          Case HENRYS_CONSTANT_INPUT

             SIValue = phprop.HenrysConstant.input.Value
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call HCCONV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If

             SourceOfValueToDisplay = phprop.HenrysConstant.input.source.short
             hc_form!lblCurrentValues(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
             hc_form!lblCurrentValues(1).Caption = hc_form!lblSource(5).Caption
       End Select
       HaveProperty(HENRYS_CONSTANT) = True
       phprop.HenrysConstant.CurrentSelection.choice = ValueToDisplayIndex
       phprop.HenrysConstant.CurrentSelection.Value = SIValue
       phprop.HenrysConstant.CurrentSelection.source = SourceOfValueToDisplay

       contam_prop_form!lblContaminantProperties(2).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
    End If

End Sub

Sub DisplayLiquidDensity()
    Dim ValueToDisplayIndex As Integer
    Dim PropertySourceToHighlight As Integer
    Dim I As Integer
    Dim SIValue As Double, EnglishValue As Double
    Dim ValueToDisplay As Double


' ***** BEGIN Calculations to determine which value is displayed on
' ***** main screen according to hierarchy

    ldens_form!lblCurrentValues(0).Caption = ""
    ldens_form!lblCurrentValues(1).Caption = ""

    Call DisplayLiquidDensityMainScreen(ValueToDisplayIndex)

' ***** END Calculations to determine which value is displayed on
' ***** main screen according to hierarchy


' ***** BEGIN Displaying liquid density values in liquid density
' ***** form (ldens_form)

'   *** Initialize all text and label boxes on ldens_form to gray and empty
    For I = 0 To 1
        ldens_form!Option1(I + 1).BackColor = &HC0C0C0
        ldens_form!Option1(I + 1).Enabled = False
        ldens_form!Option1(I + 1).Value = False
        ldens_form!lblSource(I).BackColor = &HC0C0C0
        ldens_form!lblLiquidDensityValue(I).Caption = "Not Available"
'        ldens_form!lblLiquidDensityValue(I).Enabled = False
        ldens_form!lblLiquidDensityValue(I).BackColor = &HC0C0C0
        ldens_form!lblLDTemperature(I).Caption = ""
        ldens_form!lblLDTemperature(I).Enabled = False
        ldens_form!lblLDTemperature(I).BackColor = &HC0C0C0
        ldens_form!lblLDminimumT(I).Caption = ""
        ldens_form!lblLDminimumT(I).Enabled = False
        ldens_form!lblLDminimumT(I).BackColor = &HC0C0C0
        ldens_form!lblLDmaximumT(I).Caption = ""
        ldens_form!lblLDmaximumT(I).Enabled = False
        ldens_form!lblLDmaximumT(I).BackColor = &HC0C0C0
    Next I

        ldens_form!Option1(3).BackColor = &HC0C0C0
        ldens_form!Option1(3).Enabled = False
        ldens_form!Option1(3).Value = False
        ldens_form!lblSource(2).BackColor = &HC0C0C0
        ldens_form!txtLiquidDensityValue(2).Text = ""
        ldens_form!txtLiquidDensityValue(2).Enabled = False
        ldens_form!txtLiquidDensityValue(2).BackColor = &HC0C0C0
        ldens_form!txtLDTemperature(2).Text = ""
        ldens_form!txtLDTemperature(2).Enabled = False
        ldens_form!txtLDTemperature(2).BackColor = &HC0C0C0
        ldens_form!txtLDminimumT(2).Text = ""
        ldens_form!txtLDminimumT(2).Enabled = False
        ldens_form!txtLDminimumT(2).BackColor = &HC0C0C0
        ldens_form!txtLDmaximumT(2).Text = ""
        ldens_form!txtLDmaximumT(2).Enabled = False
        ldens_form!txtLDmaximumT(2).BackColor = &HC0C0C0

    If PROPAVAILABLE(LIQUID_DENSITY_DATABASE) Then

       SIValue = phprop.LiquidDensity.database.Value
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call LDENSCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       ldens_form!lblLiquidDensityValue(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       SIValue = phprop.LiquidDensity.database.temperature
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call TEMPCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       ldens_form!lblLDTemperature(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       SIValue = phprop.LiquidDensity.dbase_minT
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call TEMPCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       ldens_form!lblLDminimumT(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       SIValue = phprop.LiquidDensity.dbase_maxT
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call TEMPCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       ldens_form!lblLDmaximumT(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       '*** Set colors of available choices to white
       ldens_form!Option1(1).BackColor = &HFFFFFF
       ldens_form!Option1(1).Enabled = True
       ldens_form!lblSource(0).BackColor = &HFFFFFF
       ldens_form!lblLiquidDensityValue(0).Enabled = True
       ldens_form!lblLiquidDensityValue(0).BackColor = &HFFFFFF
       ldens_form!lblLDTemperature(0).Enabled = True
       ldens_form!lblLDTemperature(0).BackColor = &HFFFFFF
       ldens_form!lblLDminimumT(0).Enabled = True
       ldens_form!lblLDminimumT(0).BackColor = &HFFFFFF
       ldens_form!lblLDmaximumT(0).Enabled = True
       ldens_form!lblLDmaximumT(0).BackColor = &HFFFFFF
       
       If ValueToDisplayIndex = LIQUID_DENSITY_DATABASE Then
          ldens_form!Option1(1).Value = True
          PropertySourceToHighlight = 0
       Else
          ldens_form!Option1(1).Value = False
       End If
    End If

    If PROPAVAILABLE(LIQUID_DENSITY_UNIFAC) Then

       SIValue = phprop.LiquidDensity.UNIFAC.Value
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call LDENSCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       ldens_form!lblLiquidDensityValue(1).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       SIValue = phprop.LiquidDensity.UNIFAC.temperature
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call TEMPCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       ldens_form!lblLDTemperature(1).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       ldens_form!lblLDminimumT(1).Caption = "N/A"
       ldens_form!lblLDmaximumT(1).Caption = "N/A"
       '*** Set colors of available choices to white
       ldens_form!Option1(2).BackColor = &HFFFFFF
       ldens_form!Option1(2).Enabled = True
       ldens_form!lblSource(1).BackColor = &HFFFFFF
       ldens_form!lblLiquidDensityValue(1).Enabled = True
       ldens_form!lblLiquidDensityValue(1).BackColor = &HFFFFFF
       ldens_form!lblLDTemperature(1).Enabled = True
       ldens_form!lblLDTemperature(1).BackColor = &HFFFFFF
       ldens_form!lblLDminimumT(1).Enabled = True
       ldens_form!lblLDminimumT(1).BackColor = &HFFFFFF
       ldens_form!lblLDmaximumT(1).Enabled = True
       ldens_form!lblLDmaximumT(1).BackColor = &HFFFFFF
       
       If ValueToDisplayIndex = LIQUID_DENSITY_UNIFAC Then
          ldens_form!Option1(2).Value = True
          PropertySourceToHighlight = 1
       Else
          ldens_form!Option1(2).Value = False
       End If
    End If

'  *** User input always possible so set backcolor to white
    ldens_form!Option1(3).BackColor = &HFFFFFF
    ldens_form!Option1(3).Enabled = True
    ldens_form!lblSource(2).BackColor = &HFFFFFF
    ldens_form!txtLiquidDensityValue(2).Enabled = True
    ldens_form!txtLiquidDensityValue(2).BackColor = &HFFFFFF
    ldens_form!txtLDTemperature(2).Enabled = True
    ldens_form!txtLDTemperature(2).BackColor = &HFFFFFF
    ldens_form!txtLDminimumT(2).Enabled = True
    ldens_form!txtLDminimumT(2).BackColor = &HFFFFFF
    ldens_form!txtLDmaximumT(2).Enabled = True
    ldens_form!txtLDmaximumT(2).BackColor = &HFFFFFF

    If PROPAVAILABLE(LIQUID_DENSITY_INPUT) Then

       SIValue = phprop.LiquidDensity.input.Value
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call LDENSCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       ldens_form!txtLiquidDensityValue(2).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       If HaveTemp(phprop.LiquidDensity.input.temperature) Then
          SIValue = phprop.LiquidDensity.input.temperature
          If CurrentUnits = SIUnits Then
             ValueToDisplay = SIValue
          ElseIf CurrentUnits = EnglishUnits Then
             Call TEMPCNV(EnglishValue, SIValue)
             ValueToDisplay = EnglishValue
          End If
          ldens_form!txtLDTemperature(2).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
       Else
          ldens_form!txtLDTemperature(2).Text = ""
       End If

       ldens_form!txtLDminimumT(2).Text = "N/A"
       ldens_form!txtLDmaximumT(2).Text = "N/A"

       If ValueToDisplayIndex = LIQUID_DENSITY_INPUT Then
          ldens_form!Option1(3).Value = True
          PropertySourceToHighlight = 2
       Else
          ldens_form!Option1(3).Value = False
       End If

    End If

       For I = 0 To 2
           ldens_form!lblSource(I).ForeColor = &H80000008
       Next I

    '*** Highlight selected property source
    If ValueToDisplayIndex <> 0 Then
       ldens_form!lblSource(PropertySourceToHighlight).BackColor = &H800000
       ldens_form!lblSource(PropertySourceToHighlight).ForeColor = &H80000005
       hilight.LiquidDensity.PreviousIndex = PropertySourceToHighlight
    End If

' ***** END Displaying liquid density values in liquid density
' ***** form (ldens_form)


End Sub

Sub DisplayLiquidDensityMainScreen(ValueToDisplayIndex As Integer)
    Dim DisplayedValueOnMainScreen As Integer
    Dim ValueToDisplay As Double
    Dim SourceOfValueToDisplay As Long
    Dim SIValue As Double, EnglishValue As Double
    
    If phprop.LiquidDensity.CurrentSelection.choice = 0 Then
       DisplayedValueOnMainScreen = False
       ValueToDisplayIndex = 0
    Else
       DisplayedValueOnMainScreen = True
       ValueToDisplayIndex = phprop.LiquidDensity.CurrentSelection.choice
    End If
    
    If Not DisplayedValueOnMainScreen Then
       Call CheckLiquidDensity(1, ValueToDisplayIndex, DisplayedValueOnMainScreen)
    End If

    If Not DisplayedValueOnMainScreen Then
       Call CheckLiquidDensity(2, ValueToDisplayIndex, DisplayedValueOnMainScreen)
    End If

    If Not DisplayedValueOnMainScreen Then
       Call CheckLiquidDensity(3, ValueToDisplayIndex, DisplayedValueOnMainScreen)
    End If

    If Not DisplayedValueOnMainScreen Then
       contam_prop_form!lblContaminantProperties(5).Caption = "Not Available"
       HaveProperty(LIQUID_DENSITY) = False
    Else
       Select Case ValueToDisplayIndex

          Case LIQUID_DENSITY_DATABASE

             SIValue = phprop.LiquidDensity.database.Value
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call LDENSCNV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If
             
             SourceOfValueToDisplay = phprop.LiquidDensity.database.source.short
             ldens_form!lblCurrentValues(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
             ldens_form!lblCurrentValues(1).Caption = ldens_form!lblSource(0).Caption

          Case LIQUID_DENSITY_UNIFAC

             SIValue = phprop.LiquidDensity.UNIFAC.Value
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call LDENSCNV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If
            
             SourceOfValueToDisplay = phprop.LiquidDensity.UNIFAC.source.short
             ldens_form!lblCurrentValues(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
             ldens_form!lblCurrentValues(1).Caption = ldens_form!lblSource(1).Caption

          Case LIQUID_DENSITY_INPUT

             SIValue = phprop.LiquidDensity.input.Value
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call LDENSCNV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If
            
             SourceOfValueToDisplay = phprop.LiquidDensity.input.source.short
             ldens_form!lblCurrentValues(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
             ldens_form!lblCurrentValues(1).Caption = ldens_form!lblSource(2).Caption
      End Select
      HaveProperty(LIQUID_DENSITY) = True
      phprop.LiquidDensity.CurrentSelection.choice = ValueToDisplayIndex
      phprop.LiquidDensity.CurrentSelection.Value = SIValue
      phprop.LiquidDensity.CurrentSelection.source = SourceOfValueToDisplay

      contam_prop_form!lblContaminantProperties(5).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
    End If

End Sub

Sub DisplayLiquidDiffusivity(MolecularWeight As Double)
    Dim ValueToDisplayIndex As Integer
    Dim I As Integer
    Dim PropertySourceToHighlight As Integer
    Dim EnglishValue As Double, SIValue As Double
    Dim ValueToDisplay As Double

' ***** BEGIN Calculations to determine which value is displayed on
' ***** main screen according to hierarchy

    liquid_diff_form!lblCurrentValues(0).Caption = ""
    liquid_diff_form!lblCurrentValues(1).Caption = ""

    Call DisplayLiquidDiffusivityMainScreen(ValueToDisplayIndex)

' ***** END Calculations to determine which value is displayed on
' ***** main screen according to hierarchy


' ***** BEGIN Displaying Liquid Diffusivity values
' ***** in liquid diffusivity form (liquid_diff_form)

'   *** Initialize all text and label boxes on liquid_diff_form to gray and empty
    For I = 0 To 2
        liquid_diff_form!Option1(I + 1).BackColor = &HC0C0C0
        liquid_diff_form!Option1(I + 1).Enabled = False
        liquid_diff_form!Option1(I + 1).Value = False
        liquid_diff_form!lblSource(I).BackColor = &HC0C0C0
        liquid_diff_form!lblLiquidDiffusivityValue(I).Caption = "Not Available"
'        liquid_diff_form!lblLiquidDiffusivityValue(I).Enabled = False
        liquid_diff_form!lblLiquidDiffusivityValue(I).BackColor = &HC0C0C0
        liquid_diff_form!lblLiqDiffTemperature(I).Caption = ""
        liquid_diff_form!lblLiqDiffTemperature(I).Enabled = False
        liquid_diff_form!lblLiqDiffTemperature(I).BackColor = &HC0C0C0
    Next I

        liquid_diff_form!Option1(4).BackColor = &HC0C0C0
        liquid_diff_form!Option1(4).Enabled = False
        liquid_diff_form!Option1(4).Value = False
        liquid_diff_form!lblSource(3).BackColor = &HC0C0C0
        liquid_diff_form!txtLiquidDiffusivityValue(3).Text = ""
        liquid_diff_form!txtLiquidDiffusivityValue(3).Enabled = False
        liquid_diff_form!txtLiquidDiffusivityValue(3).BackColor = &HC0C0C0
        liquid_diff_form!txtLiqDiffTemperature(3).Text = ""
        liquid_diff_form!txtLiqDiffTemperature(3).Enabled = False
        liquid_diff_form!txtLiqDiffTemperature(3).BackColor = &HC0C0C0

    If PROPAVAILABLE(LIQUID_DIFFUSIVITY_HAYDUKLAUDIE) Then

       SIValue = phprop.LiquidDiffusivity.haydukLaudie.Value
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call LDIFFCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       liquid_diff_form!lblLiquidDiffusivityValue(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       SIValue = phprop.LiquidDiffusivity.haydukLaudie.temperature
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call TEMPCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       liquid_diff_form!lblLiqDiffTemperature(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       '*** Set colors of available choices to white
       liquid_diff_form!Option1(1).BackColor = &HFFFFFF
       liquid_diff_form!Option1(1).Enabled = True
       liquid_diff_form!lblSource(0).BackColor = &HFFFFFF
       liquid_diff_form!lblLiquidDiffusivityValue(0).Enabled = True
       liquid_diff_form!lblLiquidDiffusivityValue(0).BackColor = &HFFFFFF
       liquid_diff_form!lblLiqDiffTemperature(0).Enabled = True
       liquid_diff_form!lblLiqDiffTemperature(0).BackColor = &HFFFFFF
             
       If ValueToDisplayIndex = LIQUID_DIFFUSIVITY_HAYDUKLAUDIE Then
          liquid_diff_form!Option1(1).Value = True
          PropertySourceToHighlight = 0
       Else
          liquid_diff_form!Option1(1).Value = False
       End If
    End If

    If PROPAVAILABLE(LIQUID_DIFFUSIVITY_POLSON) Then

       SIValue = phprop.LiquidDiffusivity.polson.Value
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call LDIFFCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       liquid_diff_form!lblLiquidDiffusivityValue(1).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       SIValue = phprop.LiquidDiffusivity.polson.temperature
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call TEMPCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       liquid_diff_form!lblLiqDiffTemperature(1).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       '*** Set colors of available choices to white
       liquid_diff_form!Option1(2).BackColor = &HFFFFFF
       liquid_diff_form!Option1(2).Enabled = True
       liquid_diff_form!lblSource(1).BackColor = &HFFFFFF
       liquid_diff_form!lblLiquidDiffusivityValue(1).Enabled = True
       liquid_diff_form!lblLiquidDiffusivityValue(1).BackColor = &HFFFFFF
       liquid_diff_form!lblLiqDiffTemperature(1).Enabled = True
       liquid_diff_form!lblLiqDiffTemperature(1).BackColor = &HFFFFFF
             
       If ValueToDisplayIndex = LIQUID_DIFFUSIVITY_POLSON Then
          liquid_diff_form!Option1(2).Value = True
          PropertySourceToHighlight = 1
       Else
          liquid_diff_form!Option1(2).Value = False
       End If
    End If

    If PROPAVAILABLE(LIQUID_DIFFUSIVITY_WILKECHANG) Then

       SIValue = phprop.LiquidDiffusivity.wilkeChang.Value
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call LDIFFCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       liquid_diff_form!lblLiquidDiffusivityValue(2).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       SIValue = phprop.LiquidDiffusivity.wilkeChang.temperature
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call TEMPCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       liquid_diff_form!lblLiqDiffTemperature(2).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       '*** Set colors of available choices to white
       liquid_diff_form!Option1(3).BackColor = &HFFFFFF
       liquid_diff_form!Option1(3).Enabled = True
       liquid_diff_form!lblSource(2).BackColor = &HFFFFFF
       liquid_diff_form!lblLiquidDiffusivityValue(2).Enabled = True
       liquid_diff_form!lblLiquidDiffusivityValue(2).BackColor = &HFFFFFF
       liquid_diff_form!lblLiqDiffTemperature(2).Enabled = True
       liquid_diff_form!lblLiqDiffTemperature(2).BackColor = &HFFFFFF
             
       If ValueToDisplayIndex = LIQUID_DIFFUSIVITY_WILKECHANG Then
          liquid_diff_form!Option1(3).Value = True
          PropertySourceToHighlight = 2
       Else
          liquid_diff_form!Option1(3).Value = False
       End If
    End If

'  *** User input always possible so set backcolor to white
    liquid_diff_form!Option1(4).BackColor = &HFFFFFF
    liquid_diff_form!Option1(4).Enabled = True
    liquid_diff_form!lblSource(3).BackColor = &HFFFFFF
    liquid_diff_form!txtLiquidDiffusivityValue(3).Enabled = True
    liquid_diff_form!txtLiquidDiffusivityValue(3).BackColor = &HFFFFFF
    liquid_diff_form!txtLiqDiffTemperature(3).Enabled = True
    liquid_diff_form!txtLiqDiffTemperature(3).BackColor = &HFFFFFF

    If PROPAVAILABLE(LIQUID_DIFFUSIVITY_INPUT) Then

       SIValue = phprop.LiquidDiffusivity.input.Value
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call LDIFFCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       liquid_diff_form!txtLiquidDiffusivityValue(3).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       If HaveTemp(phprop.LiquidDiffusivity.input.temperature) Then
          SIValue = phprop.LiquidDiffusivity.input.temperature
          If CurrentUnits = SIUnits Then
             ValueToDisplay = SIValue
          ElseIf CurrentUnits = EnglishUnits Then
             Call TEMPCNV(EnglishValue, SIValue)
             ValueToDisplay = EnglishValue
          End If
          liquid_diff_form!txtLiqDiffTemperature(3).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
       Else
          liquid_diff_form!txtLiqDiffTemperature(3).Text = ""
       End If

       If ValueToDisplayIndex = LIQUID_DIFFUSIVITY_INPUT Then
          liquid_diff_form!Option1(4).Value = True
          PropertySourceToHighlight = 3
       Else
          liquid_diff_form!Option1(4).Value = False
       End If

    End If

       For I = 0 To 3
           liquid_diff_form!lblSource(I).ForeColor = &H80000008
       Next I

    '*** Highlight selected property source
    If ValueToDisplayIndex <> 0 Then
       liquid_diff_form!lblSource(PropertySourceToHighlight).BackColor = &H800000
       liquid_diff_form!lblSource(PropertySourceToHighlight).ForeColor = &H80000005
       hilight.LiquidDiffusivity.PreviousIndex = PropertySourceToHighlight
    End If


' ***** END Displaying Liquid Diffusivity values
' ***** in liquid diffusivity form (liquid_diff_form)

End Sub

Sub DisplayLiquidDiffusivityMainScreen(ValueToDisplayIndex As Integer)
    Dim DisplayedValueOnMainScreen As Integer
    Dim ValueToDisplay As Double
    Static HierarchyArray(1 To 4) As HierarchyType
    Dim SourceOfValueToDisplay As Long
    Dim EnglishValue As Double, SIValue As Double

    If phprop.LiquidDiffusivity.CurrentSelection.choice = 0 Then
       DisplayedValueOnMainScreen = False
       ValueToDisplayIndex = 0
    Else
       DisplayedValueOnMainScreen = True
       ValueToDisplayIndex = phprop.LiquidDiffusivity.CurrentSelection.choice
    End If

    If MolecularWeight < 1000# Then
       For I = 1 To 4
           HierarchyArray(I) = hie.LiquidDiffusivityMWTlt1000(I)
       Next I
    Else
       For I = 1 To 4
           HierarchyArray(I) = hie.LiquidDiffusivityMWTgt1000(I)
       Next I
    End If

    If Not DisplayedValueOnMainScreen Then
       Call CheckLiquidDiffusivity(HierarchyArray(1), ValueToDisplayIndex, DisplayedValueOnMainScreen)
    End If

    If Not DisplayedValueOnMainScreen Then
       Call CheckLiquidDiffusivity(HierarchyArray(2), ValueToDisplayIndex, DisplayedValueOnMainScreen)
    End If

    If Not DisplayedValueOnMainScreen Then
       Call CheckLiquidDiffusivity(HierarchyArray(3), ValueToDisplayIndex, DisplayedValueOnMainScreen)
    End If

    If Not DisplayedValueOnMainScreen Then
       Call CheckLiquidDiffusivity(HierarchyArray(4), ValueToDisplayIndex, DisplayedValueOnMainScreen)
    End If

    If Not DisplayedValueOnMainScreen Then
       contam_prop_form!lblContaminantProperties(11).Caption = "Not Available"
       HaveProperty(LIQUID_DIFFUSIVITY) = False
    Else
       Select Case ValueToDisplayIndex

          Case LIQUID_DIFFUSIVITY_HAYDUKLAUDIE

             SIValue = phprop.LiquidDiffusivity.haydukLaudie.Value
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call LDIFFCNV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If

             SourceOfValueToDisplay = phprop.LiquidDiffusivity.haydukLaudie.source.short
             liquid_diff_form!lblCurrentValues(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
             liquid_diff_form!lblCurrentValues(1).Caption = liquid_diff_form!lblSource(0).Caption

          Case LIQUID_DIFFUSIVITY_WILKECHANG

             SIValue = phprop.LiquidDiffusivity.wilkeChang.Value
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call LDIFFCNV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If
             
             SourceOfValueToDisplay = phprop.LiquidDiffusivity.wilkeChang.source.short
             liquid_diff_form!lblCurrentValues(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
             liquid_diff_form!lblCurrentValues(1).Caption = liquid_diff_form!lblSource(2).Caption

          Case LIQUID_DIFFUSIVITY_POLSON

             SIValue = phprop.LiquidDiffusivity.polson.Value
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call LDIFFCNV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If
             
             SourceOfValueToDisplay = phprop.LiquidDiffusivity.polson.source.short
             liquid_diff_form!lblCurrentValues(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
             liquid_diff_form!lblCurrentValues(1).Caption = liquid_diff_form!lblSource(1).Caption

          Case LIQUID_DIFFUSIVITY_INPUT

             SIValue = phprop.LiquidDiffusivity.input.Value
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call LDIFFCNV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If
           
             SourceOfValueToDisplay = phprop.LiquidDiffusivity.input.source.short
             liquid_diff_form!lblCurrentValues(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
             liquid_diff_form!lblCurrentValues(1).Caption = liquid_diff_form!lblSource(3).Caption
       End Select

       HaveProperty(LIQUID_DIFFUSIVITY) = True
       phprop.LiquidDiffusivity.CurrentSelection.choice = ValueToDisplayIndex
       phprop.LiquidDiffusivity.CurrentSelection.Value = SIValue
       phprop.LiquidDiffusivity.CurrentSelection.source = SourceOfValueToDisplay

       contam_prop_form!lblContaminantProperties(11).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
    End If

End Sub

Sub DisplayMolarVolumeNBP()
    Dim ValueToDisplayIndex As Integer
    Dim PropertySourceToHighlight As Integer
    Dim I As Integer
    Dim EnglishValue As Double, SIValue As Double
    Dim ValueToDisplay As Double


' ***** BEGIN Calculations to determine which value is displayed on
' ***** main screen according to hierarchy

    mv_nbp_form!lblCurrentValues(0).Caption = ""
    mv_nbp_form!lblCurrentValues(1).Caption = ""

    Call DisplayMolarVolumeNBPMainScreen(ValueToDisplayIndex)

' ***** END Calculations to determine which value is displayed on
' ***** main screen according to hierarchy


' ***** BEGIN Displaying molar volume at normal boiling point Values
' ***** in molar volume at normal boiling point form (mv_nbp_form)

'   *** Initialize all text and label boxes on mv_nbp_form
'   *** to gray and empty and disabled
    For I = 0 To 0
        mv_nbp_form!Option1(I + 1).BackColor = &HC0C0C0
        mv_nbp_form!Option1(I + 1).Enabled = False
        mv_nbp_form!Option1(I + 1).Value = False
        mv_nbp_form!lblSource(I).BackColor = &HC0C0C0
        mv_nbp_form!lblMolarVolumeNBPValue(I).Caption = "Not Available"
'        mv_nbp_form!lblMolarVolumeNBPValue(I).Enabled = False
        mv_nbp_form!lblMolarVolumeNBPValue(I).BackColor = &HC0C0C0
        mv_nbp_form!lblMVNBPTemperature(I).Caption = ""
        mv_nbp_form!lblMVNBPTemperature(I).Enabled = False
        mv_nbp_form!lblMVNBPTemperature(I).BackColor = &HC0C0C0
    Next I

        mv_nbp_form!Option1(2).BackColor = &HC0C0C0
        mv_nbp_form!Option1(2).Enabled = False
        mv_nbp_form!Option1(2).Value = False
        mv_nbp_form!lblSource(1).BackColor = &HC0C0C0
        mv_nbp_form!txtMolarVolumeNBPValue(1).Text = ""
        mv_nbp_form!txtMolarVolumeNBPValue(1).Enabled = False
        mv_nbp_form!txtMolarVolumeNBPValue(1).BackColor = &HC0C0C0
        mv_nbp_form!txtMVNBPTemperature(1).Text = ""
        mv_nbp_form!txtMVNBPTemperature(1).Enabled = False
        mv_nbp_form!txtMVNBPTemperature(1).BackColor = &HC0C0C0

    If PROPAVAILABLE(MOLAR_VOLUME_NBP_UNIFAC) Then

       SIValue = phprop.MolarVolume.BoilingPoint.UNIFAC.Value
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call MVNBPCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       mv_nbp_form!lblMolarVolumeNBPValue(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       If HaveProperty(BOILING_POINT) Then
          SIValue = phprop.MolarVolume.BoilingPoint.UNIFAC.temperature
          If CurrentUnits = SIUnits Then
             ValueToDisplay = SIValue
          ElseIf CurrentUnits = EnglishUnits Then
             Call TEMPCNV(EnglishValue, SIValue)
             ValueToDisplay = EnglishValue
          End If
          mv_nbp_form!lblMVNBPTemperature(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
       Else
          mv_nbp_form!lblMVNBPTemperature(0).Caption = "N/A"
       End If

       '*** Set colors of available choices to white
       mv_nbp_form!Option1(1).BackColor = &HFFFFFF
       mv_nbp_form!Option1(1).Enabled = True
       mv_nbp_form!lblSource(0).BackColor = &HFFFFFF
       mv_nbp_form!lblMolarVolumeNBPValue(0).Enabled = True
       mv_nbp_form!lblMolarVolumeNBPValue(0).BackColor = &HFFFFFF
       mv_nbp_form!lblMVNBPTemperature(0).Enabled = True
       mv_nbp_form!lblMVNBPTemperature(0).BackColor = &HFFFFFF
             
       If ValueToDisplayIndex = MOLAR_VOLUME_NBP_UNIFAC Then
          mv_nbp_form!Option1(1).Value = True
          PropertySourceToHighlight = 0
       Else
          mv_nbp_form!Option1(1).Value = False
       End If
    End If

'  *** User input always possible so set backcolor to white
    mv_nbp_form!Option1(2).BackColor = &HFFFFFF
    mv_nbp_form!Option1(2).Enabled = True
    mv_nbp_form!lblSource(1).BackColor = &HFFFFFF
    mv_nbp_form!txtMolarVolumeNBPValue(1).Enabled = True
    mv_nbp_form!txtMolarVolumeNBPValue(1).BackColor = &HFFFFFF
    mv_nbp_form!txtMVNBPTemperature(1).Enabled = True
    mv_nbp_form!txtMVNBPTemperature(1).BackColor = &HFFFFFF

    If PROPAVAILABLE(MOLAR_VOLUME_NBP_INPUT) Then

       SIValue = phprop.MolarVolume.BoilingPoint.input.Value
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call MVNBPCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       mv_nbp_form!txtMolarVolumeNBPValue(1).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       If HaveTemp(phprop.MolarVolume.BoilingPoint.input.temperature) Then
          SIValue = phprop.MolarVolume.BoilingPoint.input.temperature
          If CurrentUnits = SIUnits Then
             ValueToDisplay = SIValue
          ElseIf CurrentUnits = EnglishUnits Then
             Call TEMPCNV(EnglishValue, SIValue)
             ValueToDisplay = EnglishValue
          End If
          mv_nbp_form!txtMVNBPTemperature(1).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
       Else
          mv_nbp_form!txtMVNBPTemperature(1).Text = ""
       End If

       If ValueToDisplayIndex = MOLAR_VOLUME_NBP_INPUT Then
          mv_nbp_form!Option1(2).Value = True
          PropertySourceToHighlight = 1
       Else
          mv_nbp_form!Option1(2).Value = False
       End If

    End If

       For I = 0 To 1
           mv_nbp_form!lblSource(I).ForeColor = &H80000008
       Next I

    '*** Highlight selected property source
    If ValueToDisplayIndex <> 0 Then
       mv_nbp_form!lblSource(PropertySourceToHighlight).BackColor = &H800000
       mv_nbp_form!lblSource(PropertySourceToHighlight).ForeColor = &H80000005
       hilight.MolarVolumeBoilingPoint.PreviousIndex = PropertySourceToHighlight
    End If


' ***** END Displaying molar volume at normal boiling point Values
' ***** in molar volume at normal boiling point form (mv_nbp_form)


End Sub

Sub DisplayMolarVolumeNBPMainScreen(ValueToDisplayIndex As Integer)
    Dim DisplayedValueOnMainScreen As Integer
    Dim ValueToDisplay As Double
    Dim SourceOfValueToDisplay As Long
    Dim EnglishValue As Double, SIValue As Double
    
    If phprop.MolarVolume.BoilingPoint.CurrentSelection.choice = 0 Then
       DisplayedValueOnMainScreen = False
       ValueToDisplayIndex = 0
    Else
       DisplayedValueOnMainScreen = True
       ValueToDisplayIndex = phprop.MolarVolume.BoilingPoint.CurrentSelection.choice
    End If
    
    If Not DisplayedValueOnMainScreen Then
       Call CheckMolarVolumeNBP(1, ValueToDisplayIndex, DisplayedValueOnMainScreen)
    End If

    If Not DisplayedValueOnMainScreen Then
       Call CheckMolarVolumeNBP(2, ValueToDisplayIndex, DisplayedValueOnMainScreen)
    End If

    If Not DisplayedValueOnMainScreen Then
       contam_prop_form!lblContaminantProperties(7).Caption = "Not Available"
       HaveProperty(MOLAR_VOLUME_BOILING_POINT) = False
    Else
       Select Case ValueToDisplayIndex

          Case MOLAR_VOLUME_NBP_UNIFAC

             SIValue = phprop.MolarVolume.BoilingPoint.UNIFAC.Value
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call MVNBPCNV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If
             
             SourceOfValueToDisplay = phprop.MolarVolume.BoilingPoint.UNIFAC.source.short
             mv_nbp_form!lblCurrentValues(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
             mv_nbp_form!lblCurrentValues(1).Caption = mv_nbp_form!lblSource(0).Caption

          Case MOLAR_VOLUME_NBP_INPUT

             SIValue = phprop.MolarVolume.BoilingPoint.input.Value
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call MVNBPCNV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If
             
             SourceOfValueToDisplay = phprop.MolarVolume.BoilingPoint.input.source.short
             mv_nbp_form!lblCurrentValues(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
             mv_nbp_form!lblCurrentValues(1).Caption = mv_nbp_form!lblSource(1).Caption
       End Select
       HaveProperty(MOLAR_VOLUME_BOILING_POINT) = True
       phprop.MolarVolume.BoilingPoint.CurrentSelection.choice = ValueToDisplayIndex
       phprop.MolarVolume.BoilingPoint.CurrentSelection.Value = SIValue
       phprop.MolarVolume.BoilingPoint.CurrentSelection.source = SourceOfValueToDisplay

       contam_prop_form!lblContaminantProperties(7).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
    End If

End Sub

Sub DisplayMolarVolumeOpT()
    Dim ValueToDisplayIndex As Integer
    Dim PropertySourceToHighlight As Integer
    Dim I As Integer
    Dim SIValue As Double, EnglishValue As Double
    Dim ValueToDisplay As Double


' ***** BEGIN Calculations to determine which value is displayed on
' ***** main screen according to hierarchy

    molar_vol_form!lblCurrentValues(0).Caption = ""
    molar_vol_form!lblCurrentValues(1).Caption = ""

    Call DisplayMolarVolumeOpTMainScreen(ValueToDisplayIndex)

' ***** END Calculations to determine which value is displayed on
' ***** main screen according to hierarchy


' ***** BEGIN Displaying molar volume at operating temperature values
' ***** in molar volume at operating temperature form (molar_vol_form)
                                         
'   *** Initialize all text and label boxes on molar_vol_form to gray and empty
    For I = 0 To 1
        molar_vol_form!Option1(I + 1).BackColor = &HC0C0C0
        molar_vol_form!Option1(I + 1).Enabled = False
        molar_vol_form!Option1(I + 1).Value = False
        molar_vol_form!lblSource(I).BackColor = &HC0C0C0
        molar_vol_form!lblMolarVolumeOpTValue(I).Caption = "Not Available"
'        molar_vol_form!lblMolarVolumeOpTValue(I).Enabled = False
        molar_vol_form!lblMolarVolumeOpTValue(I).BackColor = &HC0C0C0
        molar_vol_form!lblMVOpTTemperature(I).Caption = ""
        molar_vol_form!lblMVOpTTemperature(I).Enabled = False
        molar_vol_form!lblMVOpTTemperature(I).BackColor = &HC0C0C0
        molar_vol_form!lblMVOpTminimumT(I).Caption = ""
        molar_vol_form!lblMVOpTminimumT(I).Enabled = False
        molar_vol_form!lblMVOpTminimumT(I).BackColor = &HC0C0C0
        molar_vol_form!lblMVOpTmaximumT(I).Caption = ""
        molar_vol_form!lblMVOpTmaximumT(I).Enabled = False
        molar_vol_form!lblMVOpTmaximumT(I).BackColor = &HC0C0C0
    Next I

        molar_vol_form!Option1(3).BackColor = &HC0C0C0
        molar_vol_form!Option1(3).Enabled = False
        molar_vol_form!Option1(3).Value = False
        molar_vol_form!lblSource(2).BackColor = &HC0C0C0
        molar_vol_form!txtMolarVolumeOpTValue(2).Text = ""
        molar_vol_form!txtMolarVolumeOpTValue(2).Enabled = False
        molar_vol_form!txtMolarVolumeOpTValue(2).BackColor = &HC0C0C0
        molar_vol_form!txtMVOpTTemperature(2).Text = ""
        molar_vol_form!txtMVOpTTemperature(2).Enabled = False
        molar_vol_form!txtMVOpTTemperature(2).BackColor = &HC0C0C0
        molar_vol_form!txtMVOpTminimumT(2).Text = ""
        molar_vol_form!txtMVOpTminimumT(2).Enabled = False
        molar_vol_form!txtMVOpTminimumT(2).BackColor = &HC0C0C0
        molar_vol_form!txtMVOpTmaximumT(2).Text = ""
        molar_vol_form!txtMVOpTmaximumT(2).Enabled = False
        molar_vol_form!txtMVOpTmaximumT(2).BackColor = &HC0C0C0

    If PROPAVAILABLE(MOLAR_VOLUME_OPT_DATABASE) Then

       SIValue = phprop.MolarVolume.operatingT.database.Value
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call MVOPTCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       molar_vol_form!lblMolarVolumeOpTValue(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       SIValue = phprop.MolarVolume.operatingT.database.temperature
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call TEMPCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       molar_vol_form!lblMVOpTTemperature(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       SIValue = phprop.LiquidDensity.dbase_minT
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call TEMPCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       molar_vol_form!lblMVOpTminimumT(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       SIValue = phprop.LiquidDensity.dbase_maxT
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call TEMPCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       molar_vol_form!lblMVOpTmaximumT(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       '*** Set colors of available choices to white
       molar_vol_form!Option1(1).BackColor = &HFFFFFF
       molar_vol_form!Option1(1).Enabled = True
       molar_vol_form!lblSource(0).BackColor = &HFFFFFF
       molar_vol_form!lblMolarVolumeOpTValue(0).Enabled = True
       molar_vol_form!lblMolarVolumeOpTValue(0).BackColor = &HFFFFFF
       molar_vol_form!lblMVOpTTemperature(0).Enabled = True
       molar_vol_form!lblMVOpTTemperature(0).BackColor = &HFFFFFF
       molar_vol_form!lblMVOpTminimumT(0).Enabled = True
       molar_vol_form!lblMVOpTminimumT(0).BackColor = &HFFFFFF
       molar_vol_form!lblMVOpTmaximumT(0).Enabled = True
       molar_vol_form!lblMVOpTmaximumT(0).BackColor = &HFFFFFF
       
       If ValueToDisplayIndex = MOLAR_VOLUME_OPT_DATABASE Then
          molar_vol_form!Option1(1).Value = True
          PropertySourceToHighlight = 0
       Else
          molar_vol_form!Option1(1).Value = False
       End If
    End If

    If PROPAVAILABLE(MOLAR_VOLUME_OPT_UNIFAC) Then

       SIValue = phprop.MolarVolume.operatingT.UNIFAC.Value
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call MVOPTCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       molar_vol_form!lblMolarVolumeOpTValue(1).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       SIValue = phprop.MolarVolume.operatingT.UNIFAC.temperature
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call TEMPCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       molar_vol_form!lblMVOpTTemperature(1).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       molar_vol_form!lblMVOpTminimumT(1).Caption = "N/A"
       molar_vol_form!lblMVOpTmaximumT(1).Caption = "N/A"

       '*** Set colors of available choices to white
       molar_vol_form!Option1(2).BackColor = &HFFFFFF
       molar_vol_form!Option1(2).Enabled = True
       molar_vol_form!lblSource(1).BackColor = &HFFFFFF
       molar_vol_form!lblMolarVolumeOpTValue(1).Enabled = True
       molar_vol_form!lblMolarVolumeOpTValue(1).BackColor = &HFFFFFF
       molar_vol_form!lblMVOpTTemperature(1).Enabled = True
       molar_vol_form!lblMVOpTTemperature(1).BackColor = &HFFFFFF
       molar_vol_form!lblMVOpTminimumT(1).Enabled = True
       molar_vol_form!lblMVOpTminimumT(1).BackColor = &HFFFFFF
       molar_vol_form!lblMVOpTmaximumT(1).Enabled = True
       molar_vol_form!lblMVOpTmaximumT(1).BackColor = &HFFFFFF
       
       If ValueToDisplayIndex = MOLAR_VOLUME_OPT_UNIFAC Then
          molar_vol_form!Option1(2).Value = True
          PropertySourceToHighlight = 1
       Else
          molar_vol_form!Option1(2).Value = False
       End If
    End If

'  *** User input always possible so set backcolor to white
    molar_vol_form!Option1(3).BackColor = &HFFFFFF
    molar_vol_form!Option1(3).Enabled = True
    molar_vol_form!lblSource(2).BackColor = &HFFFFFF
    molar_vol_form!txtMolarVolumeOpTValue(2).Enabled = True
    molar_vol_form!txtMolarVolumeOpTValue(2).BackColor = &HFFFFFF
    molar_vol_form!txtMVOpTTemperature(2).Enabled = True
    molar_vol_form!txtMVOpTTemperature(2).BackColor = &HFFFFFF
    molar_vol_form!txtMVOpTminimumT(2).Enabled = True
    molar_vol_form!txtMVOpTminimumT(2).BackColor = &HFFFFFF
    molar_vol_form!txtMVOpTmaximumT(2).Enabled = True
    molar_vol_form!txtMVOpTmaximumT(2).BackColor = &HFFFFFF

    If PROPAVAILABLE(MOLAR_VOLUME_OPT_INPUT) Then

       SIValue = phprop.MolarVolume.operatingT.input.Value
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call MVOPTCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       molar_vol_form!txtMolarVolumeOpTValue(2).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       If HaveTemp(phprop.MolarVolume.operatingT.input.temperature) Then
          SIValue = phprop.MolarVolume.operatingT.input.temperature
          If CurrentUnits = SIUnits Then
             ValueToDisplay = SIValue
          ElseIf CurrentUnits = EnglishUnits Then
             Call TEMPCNV(EnglishValue, SIValue)
             ValueToDisplay = EnglishValue
          End If
          molar_vol_form!txtMVOpTTemperature(2).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
       Else
          molar_vol_form!txtMVOpTTemperature(2).Text = ""
       End If

       molar_vol_form!txtMVOpTminimumT(2).Text = "N/A"
       molar_vol_form!txtMVOpTmaximumT(2).Text = "N/A"

       If ValueToDisplayIndex = MOLAR_VOLUME_OPT_INPUT Then
          molar_vol_form!Option1(3).Value = True
          PropertySourceToHighlight = 2
       Else
          molar_vol_form!Option1(3).Value = False
       End If

    End If

       For I = 0 To 2
           molar_vol_form!lblSource(I).ForeColor = &H80000008
       Next I

    '*** Highlight selected property source
    If ValueToDisplayIndex <> 0 Then
       molar_vol_form!lblSource(PropertySourceToHighlight).BackColor = &H800000
       molar_vol_form!lblSource(PropertySourceToHighlight).ForeColor = &H80000005
       hilight.MolarVolumeOperatingT.PreviousIndex = PropertySourceToHighlight
    End If


' ***** END Displaying molar volume at operating temperature values
' ***** in molar volume at operating temperature form (molar_vol_form)

End Sub

Sub DisplayMolarVolumeOpTMainScreen(ValueToDisplayIndex As Integer)
    Dim ValueToDisplay As Double
    Dim DisplayedValueOnMainScreen As Integer
    Dim SourceOfValueToDisplay As Long
    Dim SIValue As Double, EnglishValue As Double

    If phprop.MolarVolume.operatingT.CurrentSelection.choice = 0 Then
       DisplayedValueOnMainScreen = False
       ValueToDisplayIndex = 0
    Else
       DisplayedValueOnMainScreen = True
       ValueToDisplayIndex = phprop.MolarVolume.operatingT.CurrentSelection.choice
    End If
    
    If Not DisplayedValueOnMainScreen Then
       Call CheckMolarVolumeOpT(1, ValueToDisplayIndex, DisplayedValueOnMainScreen)
    End If

    If Not DisplayedValueOnMainScreen Then
       Call CheckMolarVolumeOpT(2, ValueToDisplayIndex, DisplayedValueOnMainScreen)
    End If

    If Not DisplayedValueOnMainScreen Then
       Call CheckMolarVolumeOpT(3, ValueToDisplayIndex, DisplayedValueOnMainScreen)
    End If

    If Not DisplayedValueOnMainScreen Then
       contam_prop_form!lblContaminantProperties(6).Caption = "Not Available"
       HaveProperty(MOLAR_VOLUME_OPT) = False
    Else
       Select Case ValueToDisplayIndex

          Case MOLAR_VOLUME_OPT_DATABASE

             SIValue = phprop.MolarVolume.operatingT.database.Value
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call MVOPTCNV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If
             
             SourceOfValueToDisplay = phprop.MolarVolume.operatingT.database.source.short
             molar_vol_form!lblCurrentValues(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
             molar_vol_form!lblCurrentValues(1).Caption = molar_vol_form!lblSource(0).Caption

          Case MOLAR_VOLUME_OPT_UNIFAC

             SIValue = phprop.MolarVolume.operatingT.UNIFAC.Value
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call MVOPTCNV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If
            
             SourceOfValueToDisplay = phprop.MolarVolume.operatingT.UNIFAC.source.short
             molar_vol_form!lblCurrentValues(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
             molar_vol_form!lblCurrentValues(1).Caption = molar_vol_form!lblSource(1).Caption

          Case MOLAR_VOLUME_OPT_INPUT

             SIValue = phprop.MolarVolume.operatingT.input.Value
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call MVOPTCNV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If
            
             SourceOfValueToDisplay = phprop.MolarVolume.operatingT.input.source.short
             molar_vol_form!lblCurrentValues(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
             molar_vol_form!lblCurrentValues(1).Caption = molar_vol_form!lblSource(2).Caption
       End Select
       HaveProperty(MOLAR_VOLUME_OPT) = True
       phprop.MolarVolume.operatingT.CurrentSelection.choice = ValueToDisplayIndex
       phprop.MolarVolume.operatingT.CurrentSelection.Value = SIValue
       phprop.MolarVolume.operatingT.CurrentSelection.source = SourceOfValueToDisplay

       contam_prop_form!lblContaminantProperties(6).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
    End If

End Sub

Sub DisplayMolecularWeight()
    Dim ValueToDisplayIndex As Integer
    Dim PropertySourceToHighlight As Integer
    Dim I As Integer
    Dim SIValue As Double, EnglishValue As Double
    Dim ValueToDisplay As Double


' ***** BEGIN Calculations to determine which value is displayed on
' ***** main screen according to hierarchy

    mwt_form!lblCurrentValues(0).Caption = ""
    mwt_form!lblCurrentValues(1).Caption = ""

    Call DisplayMolecularWeightMainScreen(ValueToDisplayIndex)

' ***** END Calculations to determine which value is displayed on
' ***** main screen according to hierarchy


' ***** BEGIN Displaying Molecular Weight Values in molecular weight
' ***** form (mwt_form)

'   *** Initialize all text and label boxes on mwt_form
'   *** to gray and empty and disabled
    For I = 0 To 1
        mwt_form!Option1(I + 1).BackColor = &HC0C0C0
        mwt_form!Option1(I + 1).Enabled = False
        mwt_form!Option1(I + 1).Value = False
        mwt_form!lblSourceLabel(I).BackColor = &HC0C0C0
        mwt_form!lblMolecularWeightValue(I).Caption = "Not Available"
'        mwt_form!lblMolecularWeightValue(I).Enabled = False
        mwt_form!lblMolecularWeightValue(I).BackColor = &HC0C0C0
    Next I

        mwt_form!Option1(3).BackColor = &HC0C0C0
        mwt_form!Option1(3).Enabled = False
        mwt_form!Option1(3).Value = False
        mwt_form!lblSourceLabel(2).BackColor = &HC0C0C0
        mwt_form!txtMolecularWeightValue(2).Text = ""
        mwt_form!txtMolecularWeightValue(2).Enabled = False
        mwt_form!txtMolecularWeightValue(2).BackColor = &HC0C0C0


    If PROPAVAILABLE(MOLECULAR_WEIGHT_DATABASE) Then

       SIValue = phprop.MolecularWeight.database.Value
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call MWCONV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       mwt_form!lblMolecularWeightValue(0).Caption = Format$(ValueToDisplay, MOLECULAR_WEIGHT_FORMAT)

       '*** Set colors of available choices to white
       mwt_form!Option1(1).BackColor = &HFFFFFF
       mwt_form!Option1(1).Enabled = True
       mwt_form!lblSourceLabel(0).BackColor = &HFFFFFF
       mwt_form!lblMolecularWeightValue(0).Enabled = True
       mwt_form!lblMolecularWeightValue(0).BackColor = &HFFFFFF
             
       If ValueToDisplayIndex = MOLECULAR_WEIGHT_DATABASE Then
          mwt_form!Option1(1).Value = True
          PropertySourceToHighlight = 0
       Else
          mwt_form!Option1(1).Value = False
       End If
    End If

    If PROPAVAILABLE(MOLECULAR_WEIGHT_UNIFAC) Then

       SIValue = phprop.MolecularWeight.UNIFAC.Value
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call MWCONV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       mwt_form!lblMolecularWeightValue(1).Caption = Format$(ValueToDisplay, MOLECULAR_WEIGHT_FORMAT)

       '*** Set colors of available choices to white
       mwt_form!Option1(2).BackColor = &HFFFFFF
       mwt_form!Option1(2).Enabled = True
       mwt_form!lblSourceLabel(1).BackColor = &HFFFFFF
       mwt_form!lblMolecularWeightValue(1).Enabled = True
       mwt_form!lblMolecularWeightValue(1).BackColor = &HFFFFFF
             
       If ValueToDisplayIndex = MOLECULAR_WEIGHT_UNIFAC Then
          mwt_form!Option1(2).Value = True
          PropertySourceToHighlight = 1
       Else
          mwt_form!Option1(2).Value = False
       End If
    End If

'  *** User input always possible so set backcolor to white
    mwt_form!Option1(3).BackColor = &HFFFFFF
    mwt_form!Option1(3).Enabled = True
    mwt_form!lblSourceLabel(2).BackColor = &HFFFFFF
    mwt_form!txtMolecularWeightValue(2).Enabled = True
    mwt_form!txtMolecularWeightValue(2).BackColor = &HFFFFFF

    If PROPAVAILABLE(MOLECULAR_WEIGHT_INPUT) Then

       SIValue = phprop.MolecularWeight.input.Value
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call MWCONV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       mwt_form!txtMolecularWeightValue(2).Text = Format$(ValueToDisplay, MOLECULAR_WEIGHT_FORMAT)

       If ValueToDisplayIndex = MOLECULAR_WEIGHT_INPUT Then
          mwt_form!Option1(3).Value = True
          PropertySourceToHighlight = 2
       Else
          mwt_form!Option1(3).Value = False
       End If

    End If

       For I = 0 To 2
           mwt_form!lblSourceLabel(I).ForeColor = &H80000008
       Next I

    '*** Highlight selected property source
    If ValueToDisplayIndex <> 0 Then
       mwt_form!lblSourceLabel(PropertySourceToHighlight).BackColor = &H800000
       mwt_form!lblSourceLabel(PropertySourceToHighlight).ForeColor = &H80000005
       hilight.MolecularWeight.PreviousIndex = PropertySourceToHighlight
    End If

End Sub

Sub DisplayMolecularWeightMainScreen(ValueToDisplayIndex As Integer)
    Dim DisplayedValueOnMainScreen As Integer
    Dim ValueToDisplay As Double
    Dim SourceOfValueToDisplay As Long
    Dim SIValue As Double, EnglishValue As Double

    If phprop.MolecularWeight.CurrentSelection.choice = 0 Then
       DisplayedValueOnMainScreen = False
       ValueToDisplayIndex = 0
    Else
       DisplayedValueOnMainScreen = True
       ValueToDisplayIndex = phprop.MolecularWeight.CurrentSelection.choice
    End If
    
    If Not DisplayedValueOnMainScreen Then
       Call CheckMolecularWeight(1, ValueToDisplayIndex, DisplayedValueOnMainScreen)
    End If

    If Not DisplayedValueOnMainScreen Then
       Call CheckMolecularWeight(2, ValueToDisplayIndex, DisplayedValueOnMainScreen)
    End If

    If Not DisplayedValueOnMainScreen Then
       Call CheckMolecularWeight(3, ValueToDisplayIndex, DisplayedValueOnMainScreen)
    End If

    If Not DisplayedValueOnMainScreen Then
       contam_prop_form!lblContaminantProperties(3).Caption = "Not Available"
       HaveProperty(MOLECULAR_WEIGHT) = False
    Else
       Select Case ValueToDisplayIndex

          Case MOLECULAR_WEIGHT_DATABASE

             SIValue = phprop.MolecularWeight.database.Value
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call MWCONV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If
            
             SourceOfValueToDisplay = phprop.MolecularWeight.database.source.short
             mwt_form!lblCurrentValues(0).Caption = Format$(ValueToDisplay, MOLECULAR_WEIGHT_FORMAT)
             mwt_form!lblCurrentValues(1).Caption = mwt_form!lblSourceLabel(0).Caption

          Case MOLECULAR_WEIGHT_UNIFAC

             SIValue = phprop.MolecularWeight.UNIFAC.Value
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call MWCONV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If
             
             SourceOfValueToDisplay = phprop.MolecularWeight.UNIFAC.source.short
             mwt_form!lblCurrentValues(0).Caption = Format$(ValueToDisplay, MOLECULAR_WEIGHT_FORMAT)
             mwt_form!lblCurrentValues(1).Caption = mwt_form!lblSourceLabel(1).Caption

          Case MOLECULAR_WEIGHT_INPUT

             SIValue = phprop.MolecularWeight.input.Value
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call MWCONV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If

             SourceOfValueToDisplay = phprop.MolecularWeight.input.source.short
             mwt_form!lblCurrentValues(0).Caption = Format$(ValueToDisplay, MOLECULAR_WEIGHT_FORMAT)
             mwt_form!lblCurrentValues(1).Caption = mwt_form!lblSourceLabel(2).Caption
       End Select
       HaveProperty(MOLECULAR_WEIGHT) = True
       phprop.MolecularWeight.CurrentSelection.choice = ValueToDisplayIndex
       phprop.MolecularWeight.CurrentSelection.Value = SIValue
       phprop.MolecularWeight.CurrentSelection.source = SourceOfValueToDisplay

       contam_prop_form!lblContaminantProperties(3).Caption = Format$(ValueToDisplay, MOLECULAR_WEIGHT_FORMAT)
    End If

End Sub

Sub DisplayOctWaterPartCoeff()
    Dim ValueToDisplayIndex As Integer
    Dim PropertySourceToHighlight As Integer
    Dim I As Integer
    Dim EnglishValue As Double, SIValue As Double
    Dim ValueToDisplay As Double

' ***** BEGIN Calculations to determine which value is displayed on
' ***** main screen according to hierarchy

    octanol_form!lblCurrentValues(0).Caption = ""
    octanol_form!lblCurrentValues(1).Caption = ""

    Call DisplayOctWaterPartCoeffMainScreen(ValueToDisplayIndex)

' ***** END Calculations to determine which value is displayed on
' ***** main screen according to hierarchy


' ***** BEGIN Displaying Octanol Water Partition Coefficient values
' ***** in octanol water partition coefficient form (octanol_form)

'   *** Initialize all text and label boxes on octanol_form to gray and empty
    For I = 0 To 2
        octanol_form!Option1(I + 1).BackColor = &HC0C0C0
        octanol_form!Option1(I + 1).Enabled = False
        octanol_form!Option1(I + 1).Value = False
        octanol_form!lblSource(I).BackColor = &HC0C0C0
        octanol_form!lblOctWatPartCoeffValue(I).Caption = "Not Available"
'        octanol_form!lblOctWatPartCoeffValue(I).Enabled = False
        octanol_form!lblOctWatPartCoeffValue(I).BackColor = &HC0C0C0
        octanol_form!lblOWPCTemperature(I).Caption = ""
        octanol_form!lblOWPCTemperature(I).Enabled = False
        octanol_form!lblOWPCTemperature(I).BackColor = &HC0C0C0
    Next I

        octanol_form!Option1(4).BackColor = &HC0C0C0
        octanol_form!Option1(4).Enabled = False
        octanol_form!Option1(4).Value = False
        octanol_form!lblSource(3).BackColor = &HC0C0C0
        octanol_form!txtOctWatPartCoeffValue(3).Text = ""
        octanol_form!txtOctWatPartCoeffValue(3).Enabled = False
        octanol_form!txtOctWatPartCoeffValue(3).BackColor = &HC0C0C0
        octanol_form!txtOWPCTemperature(3).Text = ""
        octanol_form!txtOWPCTemperature(3).Enabled = False
        octanol_form!txtOWPCTemperature(3).BackColor = &HC0C0C0

    If PROPAVAILABLE(OCT_WATER_PART_COEFF_OPT_UNIFAC) Then

       SIValue = phprop.OctWaterPartCoeff.operatingT.UNIFAC.Value
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call KOWCONV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       octanol_form!lblOctWatPartCoeffValue(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       SIValue = phprop.OctWaterPartCoeff.operatingT.UNIFAC.temperature
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call TEMPCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       octanol_form!lblOWPCTemperature(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       '*** Set colors of available choices to white
       octanol_form!Option1(1).BackColor = &HFFFFFF
       octanol_form!Option1(1).Enabled = True
       octanol_form!lblSource(0).BackColor = &HFFFFFF
       octanol_form!lblOctWatPartCoeffValue(0).Enabled = True
       octanol_form!lblOctWatPartCoeffValue(0).BackColor = &HFFFFFF
       octanol_form!lblOWPCTemperature(0).Enabled = True
       octanol_form!lblOWPCTemperature(0).BackColor = &HFFFFFF
             
       If ValueToDisplayIndex = OCT_WATER_PART_COEFF_OPT_UNIFAC Then
          octanol_form!Option1(1).Value = True
          PropertySourceToHighlight = 0
       Else
          octanol_form!Option1(1).Value = False
       End If
    End If

    If PROPAVAILABLE(OCT_WATER_PART_COEFF_DB) Then

       SIValue = phprop.OctWaterPartCoeff.database.Value
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call KOWCONV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       octanol_form!lblOctWatPartCoeffValue(1).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       SIValue = phprop.OctWaterPartCoeff.database.temperature
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call TEMPCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       octanol_form!lblOWPCTemperature(1).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       '*** Set colors of available choices to white
       octanol_form!Option1(2).BackColor = &HFFFFFF
       octanol_form!Option1(2).Enabled = True
       octanol_form!lblSource(1).BackColor = &HFFFFFF
       octanol_form!lblOctWatPartCoeffValue(1).Enabled = True
       octanol_form!lblOctWatPartCoeffValue(1).BackColor = &HFFFFFF
       octanol_form!lblOWPCTemperature(1).Enabled = True
       octanol_form!lblOWPCTemperature(1).BackColor = &HFFFFFF
             
       If ValueToDisplayIndex = OCT_WATER_PART_COEFF_DB Then
          octanol_form!Option1(2).Value = True
          PropertySourceToHighlight = 1
       Else
          octanol_form!Option1(2).Value = False
       End If
    End If

    If PROPAVAILABLE(OCT_WATER_PART_COEFF_DBT_UNIFAC) Then

       SIValue = phprop.OctWaterPartCoeff.databaseT.UNIFAC.Value
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call KOWCONV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       octanol_form!lblOctWatPartCoeffValue(2).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       SIValue = phprop.OctWaterPartCoeff.databaseT.UNIFAC.temperature
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call TEMPCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       octanol_form!lblOWPCTemperature(2).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       '*** Set colors of available choices to white
       octanol_form!Option1(3).BackColor = &HFFFFFF
       octanol_form!Option1(3).Enabled = True
       octanol_form!lblSource(2).BackColor = &HFFFFFF
       octanol_form!lblOctWatPartCoeffValue(2).Enabled = True
       octanol_form!lblOctWatPartCoeffValue(2).BackColor = &HFFFFFF
       octanol_form!lblOWPCTemperature(2).Enabled = True
       octanol_form!lblOWPCTemperature(2).BackColor = &HFFFFFF
             
       If ValueToDisplayIndex = OCT_WATER_PART_COEFF_DBT_UNIFAC Then
          octanol_form!Option1(3).Value = True
          PropertySourceToHighlight = 2
       Else
          octanol_form!Option1(3).Value = False
       End If
    End If

'  *** User input always possible so set backcolor to white
    octanol_form!Option1(4).BackColor = &HFFFFFF
    octanol_form!Option1(4).Enabled = True
    octanol_form!lblSource(3).BackColor = &HFFFFFF
    octanol_form!txtOctWatPartCoeffValue(3).Enabled = True
    octanol_form!txtOctWatPartCoeffValue(3).BackColor = &HFFFFFF
    octanol_form!txtOWPCTemperature(3).Enabled = True
    octanol_form!txtOWPCTemperature(3).BackColor = &HFFFFFF

    If PROPAVAILABLE(OCT_WATER_PART_COEFF_INPUT) Then

       SIValue = phprop.OctWaterPartCoeff.input.Value
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call KOWCONV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       octanol_form!txtOctWatPartCoeffValue(3).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       If HaveTemp(phprop.OctWaterPartCoeff.input.temperature) Then
          SIValue = phprop.OctWaterPartCoeff.input.temperature
          If CurrentUnits = SIUnits Then
             ValueToDisplay = SIValue
          ElseIf CurrentUnits = EnglishUnits Then
             Call TEMPCNV(EnglishValue, SIValue)
             ValueToDisplay = EnglishValue
          End If
          octanol_form!txtOWPCTemperature(3).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
       Else
          octanol_form!txtOWPCTemperature(3).Text = ""
       End If

       If ValueToDisplayIndex = OCT_WATER_PART_COEFF_INPUT Then
          octanol_form!Option1(4).Value = True
          PropertySourceToHighlight = 3
       Else
          octanol_form!Option1(4).Value = False
       End If

    End If

       For I = 0 To 3
           octanol_form!lblSource(I).ForeColor = &H80000008
       Next I

    '*** Highlight selected property source
    If ValueToDisplayIndex <> 0 Then
       octanol_form!lblSource(PropertySourceToHighlight).BackColor = &H800000
       octanol_form!lblSource(PropertySourceToHighlight).ForeColor = &H80000005
       hilight.OctWaterPartCoeff.PreviousIndex = PropertySourceToHighlight
    End If

' ***** END Displaying Octanol Water Partition Coefficient values
' ***** in octanol water partition coefficient form (octanol_form)

End Sub

Sub DisplayOctWaterPartCoeffMainScreen(ValueToDisplayIndex As Integer)
    Dim DisplayedValueOnMainScreen As Integer
    Dim ValueToDisplay As Double
    Dim SourceOfValueToDisplay As Long
    Dim SIValue As Double, EnglishValue As Double

    If phprop.OctWaterPartCoeff.CurrentSelection.choice = 0 Then
       DisplayedValueOnMainScreen = False
       ValueToDisplayIndex = 0
    Else
       DisplayedValueOnMainScreen = True
       ValueToDisplayIndex = phprop.OctWaterPartCoeff.CurrentSelection.choice
    End If
    
    If Not DisplayedValueOnMainScreen Then
       Call CheckOctWaterPartCoeff(1, ValueToDisplayIndex, DisplayedValueOnMainScreen)
    End If

    If Not DisplayedValueOnMainScreen Then
       Call CheckOctWaterPartCoeff(2, ValueToDisplayIndex, DisplayedValueOnMainScreen)
    End If

    If Not DisplayedValueOnMainScreen Then
       Call CheckOctWaterPartCoeff(3, ValueToDisplayIndex, DisplayedValueOnMainScreen)
    End If

    If Not DisplayedValueOnMainScreen Then
       Call CheckOctWaterPartCoeff(4, ValueToDisplayIndex, DisplayedValueOnMainScreen)
    End If

    If Not DisplayedValueOnMainScreen Then
       contam_prop_form!lblContaminantProperties(10).Caption = "Not Available"
       HaveProperty(OCT_WATER_PART_COEFF) = False
    Else
       Select Case ValueToDisplayIndex

          Case OCT_WATER_PART_COEFF_OPT_UNIFAC

             SIValue = phprop.OctWaterPartCoeff.operatingT.UNIFAC.Value
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call KOWCONV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If
            
             SourceOfValueToDisplay = phprop.OctWaterPartCoeff.operatingT.UNIFAC.source.short
             octanol_form!lblCurrentValues(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
             octanol_form!lblCurrentValues(1).Caption = octanol_form!lblSource(0).Caption

          Case OCT_WATER_PART_COEFF_DB

             SIValue = phprop.OctWaterPartCoeff.database.Value
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call KOWCONV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If
           
             SourceOfValueToDisplay = phprop.OctWaterPartCoeff.database.source.short
             octanol_form!lblCurrentValues(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
             octanol_form!lblCurrentValues(1).Caption = octanol_form!lblSource(1).Caption

          Case OCT_WATER_PART_COEFF_DBT_UNIFAC

             SIValue = phprop.OctWaterPartCoeff.databaseT.UNIFAC.Value
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call KOWCONV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If
            
             SourceOfValueToDisplay = phprop.OctWaterPartCoeff.databaseT.UNIFAC.source.short
             octanol_form!lblCurrentValues(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
             octanol_form!lblCurrentValues(1).Caption = octanol_form!lblSource(2).Caption

          Case OCT_WATER_PART_COEFF_INPUT

             SIValue = phprop.OctWaterPartCoeff.input.Value
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call KOWCONV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If
            
             SourceOfValueToDisplay = phprop.OctWaterPartCoeff.input.source.short
             octanol_form!lblCurrentValues(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
             octanol_form!lblCurrentValues(1).Caption = octanol_form!lblSource(3).Caption
       End Select

       HaveProperty(OCT_WATER_PART_COEFF) = True
       phprop.OctWaterPartCoeff.CurrentSelection.choice = ValueToDisplayIndex
       phprop.OctWaterPartCoeff.CurrentSelection.Value = SIValue
       phprop.OctWaterPartCoeff.CurrentSelection.source = SourceOfValueToDisplay

       contam_prop_form!lblContaminantProperties(10).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
    End If

End Sub

Sub DisplayRefractiveIndex()
    Dim ValueToDisplayIndex As Integer
    Dim PropertySourceToHighlight As Integer
    Dim I As Integer
    Dim SIValue As Double, EnglishValue As Double
    Dim ValueToDisplay As Double


' ***** END Calculations to determine which value is displayed on
' ***** main screen according to hierarchy

    rindex_form!lblCurrentValues(0).Caption = ""
    rindex_form!lblCurrentValues(1).Caption = ""

    Call DisplayRefractiveIndexMainScreen(ValueToDisplayIndex)

' ***** END Calculations to determine which value is displayed on
' ***** main screen according to hierarchy


' ***** BEGIN Displaying Refractive Index Values in
' ***** refractive index form (rindex_form)

'   *** Initialize all text and label boxes on rindex_form
'   *** to gray and empty and disabled
    For I = 0 To 0
        rindex_form!Option1(I + 1).BackColor = &HC0C0C0
        rindex_form!Option1(I + 1).Enabled = False
        rindex_form!Option1(I + 1).Value = False
        rindex_form!lblSource(I).BackColor = &HC0C0C0
        rindex_form!lblRefractiveIndexValue(I).Caption = "Not Available"
'        rindex_form!lblRefractiveIndexValue(I).Enabled = False
        rindex_form!lblRefractiveIndexValue(I).BackColor = &HC0C0C0
    Next I

        rindex_form!Option1(2).BackColor = &HC0C0C0
        rindex_form!Option1(2).Enabled = False
        rindex_form!Option1(2).Value = False
        rindex_form!lblSource(1).BackColor = &HC0C0C0
        rindex_form!txtRefractiveIndexValue(1).Text = ""
        rindex_form!txtRefractiveIndexValue(1).Enabled = False
        rindex_form!txtRefractiveIndexValue(1).BackColor = &HC0C0C0

    If PROPAVAILABLE(REFRACTIVE_INDEX_DATABASE) Then

       SIValue = phprop.RefractiveIndex.database.Value
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call RICONV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       rindex_form!lblRefractiveIndexValue(0).Caption = Format$(ValueToDisplay, REFRACTIVE_INDEX_FORMAT)

       '*** Set colors of available choices to white
       rindex_form!Option1(1).BackColor = &HFFFFFF
       rindex_form!Option1(1).Enabled = True
       rindex_form!lblSource(0).BackColor = &HFFFFFF
       rindex_form!lblRefractiveIndexValue(0).Enabled = True
       rindex_form!lblRefractiveIndexValue(0).BackColor = &HFFFFFF
             
       If ValueToDisplayIndex = REFRACTIVE_INDEX_DATABASE Then
          rindex_form!Option1(1).Value = True
          PropertySourceToHighlight = 0
       Else
          rindex_form!Option1(1).Value = False
       End If
    End If

'  *** User input always possible so set backcolor to white
    rindex_form!Option1(2).BackColor = &HFFFFFF
    rindex_form!Option1(2).Enabled = True
    rindex_form!lblSource(1).BackColor = &HFFFFFF
    rindex_form!txtRefractiveIndexValue(1).Enabled = True
    rindex_form!txtRefractiveIndexValue(1).BackColor = &HFFFFFF

    If PROPAVAILABLE(REFRACTIVE_INDEX_INPUT) Then

       SIValue = phprop.RefractiveIndex.input.Value
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call RICONV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       rindex_form!txtRefractiveIndexValue(1).Text = Format$(ValueToDisplay, REFRACTIVE_INDEX_FORMAT)

       If ValueToDisplayIndex = REFRACTIVE_INDEX_INPUT Then
          rindex_form!Option1(2).Value = True
          PropertySourceToHighlight = 1
       Else
          rindex_form!Option1(2).Value = False
       End If

    End If

       For I = 0 To 1
           rindex_form!lblSource(I).ForeColor = &H80000008
       Next I

    '*** Highlight selected property source
    If ValueToDisplayIndex <> 0 Then
       rindex_form!lblSource(PropertySourceToHighlight).BackColor = &H800000
       rindex_form!lblSource(PropertySourceToHighlight).ForeColor = &H80000005
       hilight.RefractiveIndex.PreviousIndex = PropertySourceToHighlight
    End If

' ***** END Displaying Refractive Index Values in
' ***** refractive index form (rindex_form)


End Sub

Sub DisplayRefractiveIndexMainScreen(ValueToDisplayIndex As Integer)
    Dim ValueToDisplay As Double
    Dim DisplayedValueOnMainScreen As Integer
    Dim SourceOfValueToDisplay As Long
    Dim EnglishValue As Double, SIValue As Double

    If phprop.RefractiveIndex.CurrentSelection.choice = 0 Then
       DisplayedValueOnMainScreen = False
       ValueToDisplayIndex = 0
    Else
       DisplayedValueOnMainScreen = True
       ValueToDisplayIndex = phprop.RefractiveIndex.CurrentSelection.choice
    End If
    
    If Not DisplayedValueOnMainScreen Then
       Call CheckRefractiveIndex(1, ValueToDisplayIndex, DisplayedValueOnMainScreen)
    End If

    If Not DisplayedValueOnMainScreen Then
       Call CheckRefractiveIndex(2, ValueToDisplayIndex, DisplayedValueOnMainScreen)
    End If

    If Not DisplayedValueOnMainScreen Then
       contam_prop_form!lblContaminantProperties(8).Caption = "Not Available"
       HaveProperty(REFRACTIVE_INDEX) = False
    Else
       Select Case ValueToDisplayIndex

          Case REFRACTIVE_INDEX_DATABASE

             SIValue = phprop.RefractiveIndex.database.Value
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call RICONV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If
            
             SourceOfValueToDisplay = phprop.RefractiveIndex.database.source.short
             rindex_form!lblCurrentValues(0).Caption = Format$(ValueToDisplay, REFRACTIVE_INDEX_FORMAT)
             rindex_form!lblCurrentValues(1).Caption = rindex_form!lblSource(0).Caption

          Case REFRACTIVE_INDEX_INPUT

             SIValue = phprop.RefractiveIndex.input.Value
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call RICONV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If
            
             SourceOfValueToDisplay = phprop.RefractiveIndex.input.source.short
             rindex_form!lblCurrentValues(0).Caption = Format$(ValueToDisplay, REFRACTIVE_INDEX_FORMAT)
             rindex_form!lblCurrentValues(1).Caption = rindex_form!lblSource(1).Caption
       End Select

       HaveProperty(REFRACTIVE_INDEX) = True
       phprop.RefractiveIndex.CurrentSelection.choice = ValueToDisplayIndex
       phprop.RefractiveIndex.CurrentSelection.Value = SIValue
       phprop.RefractiveIndex.CurrentSelection.source = SourceOfValueToDisplay

       contam_prop_form!lblContaminantProperties(8).Caption = Format$(ValueToDisplay, REFRACTIVE_INDEX_FORMAT)
    End If

End Sub

Sub DisplayVaporPressure()
    Dim ValueToDisplayIndex As Integer
    Dim SIValue As Double
    Dim EnglishValue As Double
    Dim ValueToDisplay As Double
    Dim I As Integer
    Dim PropertySourceToHighlight As Integer

' ***** BEGIN Calculations to determine which value is displayed on
' ***** main screen according to hierarchy

    vp_form!lblCurrentValues(0).Caption = ""
    vp_form!lblCurrentValues(1).Caption = ""

    Call DisplayVaporPressureMainScreen(ValueToDisplayIndex)

' ***** END Calculations to determine which value is displayed on
' ***** main screen according to hierarchy


' ***** BEGIN Displaying Vapor Pressure Values in vapor pressure
' ***** form (vp_form)

'   *** Initialize all text and label boxes on vp_form to gray and empty
    For I = 0 To 2
        vp_form!Option1(I + 1).BackColor = &HC0C0C0
        vp_form!Option1(I + 1).Enabled = False
        vp_form!Option1(I + 1).Value = False
        vp_form!lblSourceLabel(I).BackColor = &HC0C0C0
        vp_form!lblVaporPressureValue(I).Caption = "Not Available"
'        vp_form!lblVaporPressureValue(I).Enabled = False
        vp_form!lblVaporPressureValue(I).BackColor = &HC0C0C0
        vp_form!lblVPTemperature(I).Caption = ""
        vp_form!lblVPTemperature(I).Enabled = False
        vp_form!lblVPTemperature(I).BackColor = &HC0C0C0
        vp_form!lblVPminimumT(I).Caption = ""
        vp_form!lblVPminimumT(I).Enabled = False
        vp_form!lblVPminimumT(I).BackColor = &HC0C0C0
        vp_form!lblVPmaximumT(I).Caption = ""
        vp_form!lblVPmaximumT(I).Enabled = False
        vp_form!lblVPmaximumT(I).BackColor = &HC0C0C0
    Next I

        vp_form!Option1(4).BackColor = &HC0C0C0
        vp_form!Option1(4).Enabled = False
        vp_form!Option1(4).Value = False
        vp_form!lblSourceLabel(3).BackColor = &HC0C0C0
        vp_form!txtVaporPressureValue(3).Text = ""
        vp_form!txtVaporPressureValue(3).Enabled = False
        vp_form!txtVaporPressureValue(3).BackColor = &HC0C0C0
        vp_form!txtVPTemperature(3).Text = ""
        vp_form!txtVPTemperature(3).Enabled = False
        vp_form!txtVPTemperature(3).BackColor = &HC0C0C0
        vp_form!txtVPminimumT(3).Text = ""
        vp_form!txtVPminimumT(3).Enabled = False
        vp_form!txtVPminimumT(3).BackColor = &HC0C0C0
        vp_form!txtVPmaximumT(3).Text = ""
        vp_form!txtVPmaximumT(3).Enabled = False
        vp_form!txtVPmaximumT(3).BackColor = &HC0C0C0

    If PROPAVAILABLE(VAPOR_PRESSURE_DATABASE) Then
       Select Case phprop.VaporPressure.database.source.short
          Case 4   'DIPPR801

             SIValue = phprop.VaporPressure.database.Value
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call VPCONV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If
             vp_form!lblVaporPressureValue(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

             SIValue = phprop.VaporPressure.database.temperature
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call TEMPCNV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If
             vp_form!lblVPTemperature(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

             SIValue = phprop.VaporPressure.database.minimumT
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call TEMPCNV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If
             vp_form!lblVPminimumT(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

             SIValue = phprop.VaporPressure.database.maximumT
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call TEMPCNV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If
             vp_form!lblVPmaximumT(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

             '*** Set colors of available choices to white
             vp_form!Option1(1).BackColor = &HFFFFFF
             vp_form!Option1(1).Enabled = True
             vp_form!lblSourceLabel(0).BackColor = &HFFFFFF
             vp_form!lblVaporPressureValue(0).Enabled = True
             vp_form!lblVaporPressureValue(0).BackColor = &HFFFFFF
             vp_form!lblVPTemperature(0).Enabled = True
             vp_form!lblVPTemperature(0).BackColor = &HFFFFFF
             vp_form!lblVPminimumT(0).Enabled = True
             vp_form!lblVPminimumT(0).BackColor = &HFFFFFF
             vp_form!lblVPmaximumT(0).Enabled = True
             vp_form!lblVPmaximumT(0).BackColor = &HFFFFFF
             
             If ValueToDisplayIndex = VAPOR_PRESSURE_DATABASE Then
                vp_form!Option1(1).Value = True
                PropertySourceToHighlight = 0
             Else
                vp_form!Option1(1).Value = False
             End If

          Case 1   'Yaws (Antoine's)

             SIValue = phprop.VaporPressure.database.Value
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call VPCONV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If
             vp_form!lblVaporPressureValue(1).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

             SIValue = phprop.VaporPressure.database.temperature
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call TEMPCNV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If
             vp_form!lblVPTemperature(1).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

             SIValue = phprop.VaporPressure.database.minimumT
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call TEMPCNV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If
             vp_form!lblVPminimumT(1).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

             SIValue = phprop.VaporPressure.database.maximumT
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call TEMPCNV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If
             vp_form!lblVPmaximumT(1).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

             '*** Set colors of available choices to white
             vp_form!Option1(2).BackColor = &HFFFFFF
             vp_form!Option1(2).Enabled = True
             vp_form!lblSourceLabel(1).BackColor = &HFFFFFF
             vp_form!lblVaporPressureValue(1).Enabled = True
             vp_form!lblVaporPressureValue(1).BackColor = &HFFFFFF
             vp_form!lblVPTemperature(1).Enabled = True
             vp_form!lblVPTemperature(1).BackColor = &HFFFFFF
             vp_form!lblVPminimumT(1).Enabled = True
             vp_form!lblVPminimumT(1).BackColor = &HFFFFFF
             vp_form!lblVPmaximumT(1).Enabled = True
             vp_form!lblVPmaximumT(1).BackColor = &HFFFFFF

             If ValueToDisplayIndex = VAPOR_PRESSURE_DATABASE Then
                vp_form!Option1(2).Value = True
                PropertySourceToHighlight = 1
             Else
                vp_form!Option1(2).Value = False
             End If

          Case 2   'Superfund

             SIValue = phprop.VaporPressure.database.Value
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call VPCONV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If
             vp_form!lblVaporPressureValue(2).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

             SIValue = phprop.VaporPressure.database.temperature
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call TEMPCNV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If
             vp_form!lblVPTemperature(2).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

             vp_form!lblVPminimumT(2).Caption = "N/A"
             vp_form!lblVPmaximumT(2).Caption = "N/A"
             '*** Set colors of available choices to white
             vp_form!Option1(3).BackColor = &HFFFFFF
             vp_form!Option1(3).Enabled = True
             vp_form!lblSourceLabel(2).BackColor = &HFFFFFF
             vp_form!lblVaporPressureValue(2).Enabled = True
             vp_form!lblVaporPressureValue(2).BackColor = &HFFFFFF
             vp_form!lblVPTemperature(2).Enabled = True
             vp_form!lblVPTemperature(2).BackColor = &HFFFFFF
             vp_form!lblVPminimumT(2).Enabled = True
             vp_form!lblVPminimumT(2).BackColor = &HFFFFFF
             vp_form!lblVPmaximumT(2).Enabled = True
             vp_form!lblVPmaximumT(2).BackColor = &HFFFFFF

             If ValueToDisplayIndex = VAPOR_PRESSURE_DATABASE Then
                vp_form!Option1(3).Value = True
                PropertySourceToHighlight = 2
             Else
                vp_form!Option1(3).Value = False
             End If

       End Select
    End If

'  *** User input always possible so set backcolor to white
    vp_form!Option1(4).BackColor = &HFFFFFF
    vp_form!Option1(4).Enabled = True
    vp_form!lblSourceLabel(3).BackColor = &HFFFFFF
    vp_form!txtVaporPressureValue(3).Enabled = True
    vp_form!txtVaporPressureValue(3).BackColor = &HFFFFFF
    vp_form!txtVPTemperature(3).Enabled = True
    vp_form!txtVPTemperature(3).BackColor = &HFFFFFF
    vp_form!txtVPminimumT(3).Enabled = True
    vp_form!txtVPminimumT(3).BackColor = &HFFFFFF
    vp_form!txtVPmaximumT(3).Enabled = True
    vp_form!txtVPmaximumT(3).BackColor = &HFFFFFF

    If PROPAVAILABLE(VAPOR_PRESSURE_INPUT) Then

       SIValue = phprop.VaporPressure.input.Value
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call VPCONV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       vp_form!txtVaporPressureValue(3).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       If HaveTemp(phprop.VaporPressure.input.temperature) Then
          SIValue = phprop.VaporPressure.input.temperature
          If CurrentUnits = SIUnits Then
             ValueToDisplay = SIValue
          ElseIf CurrentUnits = EnglishUnits Then
             Call TEMPCNV(EnglishValue, SIValue)
             ValueToDisplay = EnglishValue
          End If
          vp_form!txtVPTemperature(3).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
       Else
          vp_form!txtVPTemperature(3).Text = ""
       End If
       vp_form!txtVPminimumT(3).Text = "N/A"
       vp_form!txtVPmaximumT(3).Text = "N/A"

       If ValueToDisplayIndex = VAPOR_PRESSURE_INPUT Then
          vp_form!Option1(4).Value = True
          PropertySourceToHighlight = 3
       Else
          vp_form!Option1(4).Value = False
       End If

    End If

       For I = 0 To 3
           vp_form!lblSourceLabel(I).ForeColor = &H80000008
       Next I

    '*** Highlight selected property source
    If ValueToDisplayIndex <> 0 Then
       vp_form!lblSourceLabel(PropertySourceToHighlight).BackColor = &H800000
       vp_form!lblSourceLabel(PropertySourceToHighlight).ForeColor = &H80000005
       hilight.VaporPressure.PreviousIndex = PropertySourceToHighlight
    End If

' ***** END Displaying Vapor Pressure Values in vapor pressure
' ***** form (vp_form)

End Sub

Sub DisplayVaporPressureMainScreen(ValueToDisplayIndex As Integer)
    Dim DisplayedValueOnMainScreen As Integer
    Dim SourceOfValueToDisplay As Long
    Dim ValueToDisplay As Double
    Dim SIValue As Double
    Dim EnglishValue As Double

    If phprop.VaporPressure.CurrentSelection.choice = 0 Then
       DisplayedValueOnMainScreen = False
       ValueToDisplayIndex = 0
    Else
       DisplayedValueOnMainScreen = True
       ValueToDisplayIndex = phprop.VaporPressure.CurrentSelection.choice
    End If
    
    If Not DisplayedValueOnMainScreen Then
       Call CheckVaporPressure(1, ValueToDisplayIndex, DisplayedValueOnMainScreen)
    End If

    If Not DisplayedValueOnMainScreen Then
       Call CheckVaporPressure(2, ValueToDisplayIndex, DisplayedValueOnMainScreen)
    End If

    If Not DisplayedValueOnMainScreen Then
       contam_prop_form!lblContaminantProperties(0).Caption = "Not Available"
       HaveProperty(VAPOR_PRESSURE) = False
    Else
       Select Case ValueToDisplayIndex
          Case VAPOR_PRESSURE_DATABASE

             SIValue = phprop.VaporPressure.database.Value
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call VPCONV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If

             SourceOfValueToDisplay = phprop.VaporPressure.database.source.short
             Select Case phprop.VaporPressure.database.source.short
                Case 4   'DIPPR801
                   vp_form!lblCurrentValues(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
                   vp_form!lblCurrentValues(1).Caption = vp_form!lblSourceLabel(0).Caption
                Case 1   'Yaws (Antoine's)
                   vp_form!lblCurrentValues(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
                   vp_form!lblCurrentValues(1).Caption = vp_form!lblSourceLabel(1).Caption
                Case 2   'Superfund
                   vp_form!lblCurrentValues(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
                   vp_form!lblCurrentValues(1).Caption = vp_form!lblSourceLabel(2).Caption
             End Select

          Case VAPOR_PRESSURE_INPUT

             SIValue = phprop.VaporPressure.input.Value
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call VPCONV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If
            
             SourceOfValueToDisplay = phprop.VaporPressure.input.source.short
             vp_form!lblCurrentValues(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
             vp_form!lblCurrentValues(1).Caption = vp_form!lblSourceLabel(3).Caption
       End Select

       HaveProperty(VAPOR_PRESSURE) = True
       phprop.VaporPressure.CurrentSelection.choice = ValueToDisplayIndex
       phprop.VaporPressure.CurrentSelection.Value = SIValue
       phprop.VaporPressure.CurrentSelection.source = SourceOfValueToDisplay

       contam_prop_form!lblContaminantProperties(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
    End If

End Sub

Sub DisplayWaterDensity()
    Dim ValueToDisplayIndex As Integer
    Dim I As Integer
    Dim PropertySourceToHighlight As Integer
    Dim EnglishValue As Double, SIValue As Double
    Dim ValueToDisplay As Double

' ***** BEGIN Calculations to determine which value is displayed on
' ***** main screen according to hierarchy

    frmWaterDensity!lblCurrentValues(0).Caption = ""
    frmWaterDensity!lblCurrentValues(1).Caption = ""

    Call DisplayWaterDensityMainScreen(ValueToDisplayIndex)

' ***** END Calculations to determine which value is displayed on
' ***** main screen according to hierarchy


' ***** BEGIN Displaying water density values in water density
' ***** form (frmWaterDensity)

'   *** Initialize all text and label boxes on frmWaterDensity to gray and empty
    For I = 0 To 0
        frmWaterDensity!Option1(I + 1).BackColor = &HC0C0C0
        frmWaterDensity!Option1(I + 1).Enabled = False
        frmWaterDensity!Option1(I + 1).Value = False
        frmWaterDensity!lblSource(I).BackColor = &HC0C0C0
        frmWaterDensity!lblWaterDensityValue(I).Caption = "Not Available"
'        frmWaterDensity!lblWaterDensityValue(I).Enabled = False
        frmWaterDensity!lblWaterDensityValue(I).BackColor = &HC0C0C0
        frmWaterDensity!lblH2ODensityTemperature(I).Caption = ""
        frmWaterDensity!lblH2ODensityTemperature(I).Enabled = False
        frmWaterDensity!lblH2ODensityTemperature(I).BackColor = &HC0C0C0
        frmWaterDensity!lblH2ODensityminimumT(I).Caption = ""
        frmWaterDensity!lblH2ODensityminimumT(I).Enabled = False
        frmWaterDensity!lblH2ODensityminimumT(I).BackColor = &HC0C0C0
        frmWaterDensity!lblH2ODensitymaximumT(I).Caption = ""
        frmWaterDensity!lblH2ODensitymaximumT(I).Enabled = False
        frmWaterDensity!lblH2ODensitymaximumT(I).BackColor = &HC0C0C0
    Next I

        frmWaterDensity!Option1(2).BackColor = &HC0C0C0
        frmWaterDensity!Option1(2).Enabled = False
        frmWaterDensity!Option1(2).Value = False
        frmWaterDensity!lblSource(1).BackColor = &HC0C0C0
        frmWaterDensity!txtWaterDensityValue(1).Text = ""
        frmWaterDensity!txtWaterDensityValue(1).Enabled = False
        frmWaterDensity!txtWaterDensityValue(1).BackColor = &HC0C0C0
        frmWaterDensity!txtH2ODensityTemperature(1).Text = ""
        frmWaterDensity!txtH2ODensityTemperature(1).Enabled = False
        frmWaterDensity!txtH2ODensityTemperature(1).BackColor = &HC0C0C0
        frmWaterDensity!txtH2ODensityminimumT(1).Text = ""
        frmWaterDensity!txtH2ODensityminimumT(1).Enabled = False
        frmWaterDensity!txtH2ODensityminimumT(1).BackColor = &HC0C0C0
        frmWaterDensity!txtH2ODensitymaximumT(1).Text = ""
        frmWaterDensity!txtH2ODensitymaximumT(1).Enabled = False
        frmWaterDensity!txtH2ODensitymaximumT(1).BackColor = &HC0C0C0

    If PROPAVAILABLE(WATER_DENSITY_CORRELATION) Then

       SIValue = phprop.WaterDensity.correlation.Value
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call WDENSCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       frmWaterDensity!lblWaterDensityValue(0).Caption = Format$(ValueToDisplay, WATER_DENSITY_FORMAT)

       SIValue = phprop.WaterDensity.correlation.temperature
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call TEMPCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       frmWaterDensity!lblH2ODensityTemperature(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       If CurrentUnits = SIUnits Then
          frmWaterDensity!lblH2ODensityminimumT(0).Caption = "0.0"
          frmWaterDensity!lblH2ODensitymaximumT(0).Caption = "100.0"
       Else
          frmWaterDensity!lblH2ODensityminimumT(0).Caption = "32.0"
          frmWaterDensity!lblH2ODensitymaximumT(0).Caption = "212.0"
       End If

       '*** Set colors of available choices to white
       frmWaterDensity!Option1(1).BackColor = &HFFFFFF
       frmWaterDensity!Option1(1).Enabled = True
       frmWaterDensity!lblSource(0).BackColor = &HFFFFFF
       frmWaterDensity!lblWaterDensityValue(0).Enabled = True
       frmWaterDensity!lblWaterDensityValue(0).BackColor = &HFFFFFF
       frmWaterDensity!lblH2ODensityTemperature(0).Enabled = True
       frmWaterDensity!lblH2ODensityTemperature(0).BackColor = &HFFFFFF
       frmWaterDensity!lblH2ODensityminimumT(0).Enabled = True
       frmWaterDensity!lblH2ODensityminimumT(0).BackColor = &HFFFFFF
       frmWaterDensity!lblH2ODensitymaximumT(0).Enabled = True
       frmWaterDensity!lblH2ODensitymaximumT(0).BackColor = &HFFFFFF
       
       If ValueToDisplayIndex = WATER_DENSITY_CORRELATION Then
          frmWaterDensity!Option1(1).Value = True
          PropertySourceToHighlight = 0
       Else
          frmWaterDensity!Option1(1).Value = False
       End If
    End If

'  *** User input always possible so set backcolor to white
    frmWaterDensity!Option1(2).BackColor = &HFFFFFF
    frmWaterDensity!Option1(2).Enabled = True
    frmWaterDensity!lblSource(1).BackColor = &HFFFFFF
    frmWaterDensity!txtWaterDensityValue(1).Enabled = True
    frmWaterDensity!txtWaterDensityValue(1).BackColor = &HFFFFFF
    frmWaterDensity!txtH2ODensityTemperature(1).Enabled = True
    frmWaterDensity!txtH2ODensityTemperature(1).BackColor = &HFFFFFF
    frmWaterDensity!txtH2ODensityminimumT(1).Enabled = True
    frmWaterDensity!txtH2ODensityminimumT(1).BackColor = &HFFFFFF
    frmWaterDensity!txtH2ODensitymaximumT(1).Enabled = True
    frmWaterDensity!txtH2ODensitymaximumT(1).BackColor = &HFFFFFF

    If PROPAVAILABLE(WATER_DENSITY_INPUT) Then

       SIValue = phprop.WaterDensity.input.Value
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call WDENSCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       frmWaterDensity!txtWaterDensityValue(1).Text = Format$(ValueToDisplay, WATER_DENSITY_FORMAT)

       If HaveTemp(phprop.WaterDensity.input.temperature) Then
          SIValue = phprop.WaterDensity.input.temperature
          If CurrentUnits = SIUnits Then
             ValueToDisplay = SIValue
          ElseIf CurrentUnits = EnglishUnits Then
             Call TEMPCNV(EnglishValue, SIValue)
             ValueToDisplay = EnglishValue
          End If
          frmWaterDensity!txtH2ODensityTemperature(1).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
       Else
          frmWaterDensity!txtH2ODensityTemperature(1).Text = ""
       End If

       frmWaterDensity!txtH2ODensityminimumT(1).Text = ""
       frmWaterDensity!txtH2ODensitymaximumT(1).Text = ""

       If ValueToDisplayIndex = WATER_DENSITY_INPUT Then
          frmWaterDensity!Option1(2).Value = True
          PropertySourceToHighlight = 1
       Else
          frmWaterDensity!Option1(2).Value = False
       End If

    End If

       For I = 0 To 1
           frmWaterDensity!lblSource(I).ForeColor = &H80000008
       Next I

    '*** Highlight selected property source
    If ValueToDisplayIndex <> 0 Then
       frmWaterDensity!lblSource(PropertySourceToHighlight).BackColor = &H800000
       frmWaterDensity!lblSource(PropertySourceToHighlight).ForeColor = &H80000005
       hilight.WaterDensity.PreviousIndex = PropertySourceToHighlight
    End If

' ***** END Displaying water density values in water density
' ***** form (frmWaterDensity)

End Sub

Sub DisplayWaterDensityMainScreen(ValueToDisplayIndex As Integer)
    Dim DisplayedValueOnMainScreen As Integer
    Dim ValueToDisplay As Double
    Dim SourceOfValueToDisplay As Long
    Dim SIValue As Double, EnglishValue As Double

    If phprop.WaterDensity.CurrentSelection.choice = 0 Then
       DisplayedValueOnMainScreen = False
       ValueToDisplayIndex = 0
    Else
       DisplayedValueOnMainScreen = True
       ValueToDisplayIndex = phprop.WaterDensity.CurrentSelection.choice
    End If
    
    If Not DisplayedValueOnMainScreen Then
       Call CheckWaterDensity(1, ValueToDisplayIndex, DisplayedValueOnMainScreen)
    End If

    If Not DisplayedValueOnMainScreen Then
       Call CheckWaterDensity(2, ValueToDisplayIndex, DisplayedValueOnMainScreen)
    End If

    If Not DisplayedValueOnMainScreen Then
       contam_prop_form!lblAirWaterProperties(0).Caption = "Not Available"
       HaveProperty(WATER_DENSITY) = False
    Else
       Select Case ValueToDisplayIndex

          Case WATER_DENSITY_CORRELATION

             SIValue = phprop.WaterDensity.correlation.Value
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call WDENSCNV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If
            
             SourceOfValueToDisplay = phprop.WaterDensity.correlation.source.short
             frmWaterDensity!lblCurrentValues(0).Caption = Format$(ValueToDisplay, WATER_DENSITY_FORMAT)
             frmWaterDensity!lblCurrentValues(1).Caption = frmWaterDensity!lblSource(0).Caption

          Case WATER_DENSITY_INPUT

             SIValue = phprop.WaterDensity.input.Value
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call WDENSCNV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If
             
             SourceOfValueToDisplay = phprop.WaterDensity.input.source.short
             frmWaterDensity!lblCurrentValues(0).Caption = Format$(ValueToDisplay, WATER_DENSITY_FORMAT)
             frmWaterDensity!lblCurrentValues(1).Caption = frmWaterDensity!lblSource(1).Caption
       End Select

       HaveProperty(WATER_DENSITY) = True
       phprop.WaterDensity.CurrentSelection.choice = ValueToDisplayIndex
       phprop.WaterDensity.CurrentSelection.Value = SIValue
       phprop.WaterDensity.CurrentSelection.source = SourceOfValueToDisplay

       contam_prop_form!lblAirWaterProperties(0).Caption = Format$(ValueToDisplay, WATER_DENSITY_FORMAT)
    End If

End Sub

Sub DisplayWaterSurfaceTension()
    Dim ValueToDisplayIndex As Integer
    Dim I As Integer
    Dim PropertySourceToHighlight As Integer
    Dim SIValue As Double, EnglishValue As Double
    Dim ValueToDisplay As Double


' ***** BEGIN Calculations to determine which value is displayed on
' ***** main screen according to hierarchy

    frmWaterSurfaceTension!lblCurrentValues(0).Caption = ""
    frmWaterSurfaceTension!lblCurrentValues(1).Caption = ""

    Call DisplayWaterSurfaceTensionMainScreen(ValueToDisplayIndex)

' ***** END Calculations to determine which value is displayed on
' ***** main screen according to hierarchy


' ***** BEGIN Displaying water surface tension values in water surface tension
' ***** form (frmWaterSurfaceTension)

'   *** Initialize all text and label boxes on frmWaterViscosity to gray and empty
    For I = 0 To 0
        frmWaterSurfaceTension!Option1(I + 1).BackColor = &HC0C0C0
        frmWaterSurfaceTension!Option1(I + 1).Enabled = False
        frmWaterSurfaceTension!Option1(I + 1).Value = False
        frmWaterSurfaceTension!lblSource(I).BackColor = &HC0C0C0
        frmWaterSurfaceTension!lblWaterSurfaceTensionValue(I).Caption = "Not Available"
'        frmWaterSurfaceTension!lblWaterSurfaceTensionValue(I).Enabled = False
        frmWaterSurfaceTension!lblWaterSurfaceTensionValue(I).BackColor = &HC0C0C0
        frmWaterSurfaceTension!lblH2OSurfTensTemperature(I).Caption = ""
        frmWaterSurfaceTension!lblH2OSurfTensTemperature(I).Enabled = False
        frmWaterSurfaceTension!lblH2OSurfTensTemperature(I).BackColor = &HC0C0C0
        frmWaterSurfaceTension!lblH2OSurfTensminimumT(I).Caption = ""
        frmWaterSurfaceTension!lblH2OSurfTensminimumT(I).Enabled = False
        frmWaterSurfaceTension!lblH2OSurfTensminimumT(I).BackColor = &HC0C0C0
        frmWaterSurfaceTension!lblH2OSurfTensmaximumT(I).Caption = ""
        frmWaterSurfaceTension!lblH2OSurfTensmaximumT(I).Enabled = False
        frmWaterSurfaceTension!lblH2OSurfTensmaximumT(I).BackColor = &HC0C0C0
    Next I

        frmWaterSurfaceTension!Option1(2).BackColor = &HC0C0C0
        frmWaterSurfaceTension!Option1(2).Enabled = False
        frmWaterSurfaceTension!Option1(2).Value = False
        frmWaterSurfaceTension!lblSource(1).BackColor = &HC0C0C0
        frmWaterSurfaceTension!txtWaterSurfaceTensionValue(1).Text = ""
        frmWaterSurfaceTension!txtWaterSurfaceTensionValue(1).Enabled = False
        frmWaterSurfaceTension!txtWaterSurfaceTensionValue(1).BackColor = &HC0C0C0
        frmWaterSurfaceTension!txtH2OSurfTensTemperature(1).Text = ""
        frmWaterSurfaceTension!txtH2OSurfTensTemperature(1).Enabled = False
        frmWaterSurfaceTension!txtH2OSurfTensTemperature(1).BackColor = &HC0C0C0
        frmWaterSurfaceTension!txtH2OSurfTensminimumT(1).Text = ""
        frmWaterSurfaceTension!txtH2OSurfTensminimumT(1).Enabled = False
        frmWaterSurfaceTension!txtH2OSurfTensminimumT(1).BackColor = &HC0C0C0
        frmWaterSurfaceTension!txtH2OSurfTensmaximumT(1).Text = ""
        frmWaterSurfaceTension!txtH2OSurfTensmaximumT(1).Enabled = False
        frmWaterSurfaceTension!txtH2OSurfTensmaximumT(1).BackColor = &HC0C0C0

    If PROPAVAILABLE(WATER_SURF_TENSION_CORRELATION) Then

       SIValue = phprop.WaterSurfaceTension.correlation.Value
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call H2OSTCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       frmWaterSurfaceTension!lblWaterSurfaceTensionValue(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       SIValue = phprop.WaterSurfaceTension.correlation.temperature
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call TEMPCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       frmWaterSurfaceTension!lblH2OSurfTensTemperature(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       frmWaterSurfaceTension!lblH2OSurfTensminimumT(0).Caption = "N/A"
       frmWaterSurfaceTension!lblH2OSurfTensmaximumT(0).Caption = "N/A"
       '*** Set colors of available choices to white
       frmWaterSurfaceTension!Option1(1).BackColor = &HFFFFFF
       frmWaterSurfaceTension!Option1(1).Enabled = True
       frmWaterSurfaceTension!lblSource(0).BackColor = &HFFFFFF
       frmWaterSurfaceTension!lblWaterSurfaceTensionValue(0).Enabled = True
       frmWaterSurfaceTension!lblWaterSurfaceTensionValue(0).BackColor = &HFFFFFF
       frmWaterSurfaceTension!lblH2OSurfTensTemperature(0).Enabled = True
       frmWaterSurfaceTension!lblH2OSurfTensTemperature(0).BackColor = &HFFFFFF
       frmWaterSurfaceTension!lblH2OSurfTensminimumT(0).Enabled = True
       frmWaterSurfaceTension!lblH2OSurfTensminimumT(0).BackColor = &HFFFFFF
       frmWaterSurfaceTension!lblH2OSurfTensmaximumT(0).Enabled = True
       frmWaterSurfaceTension!lblH2OSurfTensmaximumT(0).BackColor = &HFFFFFF
       
       If ValueToDisplayIndex = WATER_SURF_TENSION_CORRELATION Then
          frmWaterSurfaceTension!Option1(1).Value = True
          PropertySourceToHighlight = 0
       Else
          frmWaterSurfaceTension!Option1(1).Value = False
       End If
    End If

'  *** User input always possible so set backcolor to white
    frmWaterSurfaceTension!Option1(2).BackColor = &HFFFFFF
    frmWaterSurfaceTension!Option1(2).Enabled = True
    frmWaterSurfaceTension!lblSource(1).BackColor = &HFFFFFF
    frmWaterSurfaceTension!txtWaterSurfaceTensionValue(1).Enabled = True
    frmWaterSurfaceTension!txtWaterSurfaceTensionValue(1).BackColor = &HFFFFFF
    frmWaterSurfaceTension!txtH2OSurfTensTemperature(1).Enabled = True
    frmWaterSurfaceTension!txtH2OSurfTensTemperature(1).BackColor = &HFFFFFF
    frmWaterSurfaceTension!txtH2OSurfTensminimumT(1).Enabled = True
    frmWaterSurfaceTension!txtH2OSurfTensminimumT(1).BackColor = &HFFFFFF
    frmWaterSurfaceTension!txtH2OSurfTensmaximumT(1).Enabled = True
    frmWaterSurfaceTension!txtH2OSurfTensmaximumT(1).BackColor = &HFFFFFF

    If PROPAVAILABLE(WATER_SURF_TENSION_INPUT) Then

       SIValue = phprop.WaterSurfaceTension.input.Value
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call H2OSTCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       frmWaterSurfaceTension!txtWaterSurfaceTensionValue(1).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       If HaveTemp(phprop.WaterSurfaceTension.input.temperature) Then
          SIValue = phprop.WaterSurfaceTension.input.temperature
          If CurrentUnits = SIUnits Then
             ValueToDisplay = SIValue
          ElseIf CurrentUnits = EnglishUnits Then
             Call TEMPCNV(EnglishValue, SIValue)
             ValueToDisplay = EnglishValue
          End If
          frmWaterSurfaceTension!txtH2OSurfTensTemperature(1).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
       Else
          frmWaterSurfaceTension!txtH2OSurfTensTemperature(1).Text = ""
       End If
       frmWaterSurfaceTension!txtH2OSurfTensminimumT(1).Text = ""
       frmWaterSurfaceTension!txtH2OSurfTensmaximumT(1).Text = ""

       If ValueToDisplayIndex = WATER_SURF_TENSION_INPUT Then
          frmWaterSurfaceTension!Option1(2).Value = True
          PropertySourceToHighlight = 1
       Else
          frmWaterSurfaceTension!Option1(2).Value = False
       End If

    End If

       For I = 0 To 1
           frmWaterSurfaceTension!lblSource(I).ForeColor = &H80000008
       Next I

    '*** Highlight selected property source
    If ValueToDisplayIndex <> 0 Then
       frmWaterSurfaceTension!lblSource(PropertySourceToHighlight).BackColor = &H800000
       frmWaterSurfaceTension!lblSource(PropertySourceToHighlight).ForeColor = &H80000005
       hilight.WaterSurfaceTension.PreviousIndex = PropertySourceToHighlight
    End If

' ***** END Displaying water surface tension values in water surface tension
' ***** form (frmWaterSurfaceTension)

End Sub

Sub DisplayWaterSurfaceTensionMainScreen(ValueToDisplayIndex As Integer)
    Dim DisplayedValueOnMainScreen As Integer
    Dim ValueToDisplay As Double
    Dim SourceOfValueToDisplay As Long
    Dim EnglishValue As Double, SIValue As Double

    If phprop.WaterSurfaceTension.CurrentSelection.choice = 0 Then
       DisplayedValueOnMainScreen = False
       ValueToDisplayIndex = 0
    Else
       DisplayedValueOnMainScreen = True
       ValueToDisplayIndex = phprop.WaterSurfaceTension.CurrentSelection.choice
    End If
    
    If Not DisplayedValueOnMainScreen Then
       Call CheckWaterSurfaceTension(1, ValueToDisplayIndex, DisplayedValueOnMainScreen)
    End If

    If Not DisplayedValueOnMainScreen Then
       Call CheckWaterSurfaceTension(2, ValueToDisplayIndex, DisplayedValueOnMainScreen)
    End If

    If Not DisplayedValueOnMainScreen Then
       contam_prop_form!lblAirWaterProperties(2).Caption = "Not Available"
       HaveProperty(WATER_SURFACE_TENSION) = False
    Else
       Select Case ValueToDisplayIndex

          Case WATER_SURF_TENSION_CORRELATION

             SIValue = phprop.WaterSurfaceTension.correlation.Value
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call H2OSTCNV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If
             
             SourceOfValueToDisplay = phprop.WaterSurfaceTension.correlation.source.short
             frmWaterSurfaceTension!lblCurrentValues(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
             frmWaterSurfaceTension!lblCurrentValues(1).Caption = frmWaterSurfaceTension!lblSource(0).Caption

          Case WATER_SURF_TENSION_INPUT

             SIValue = phprop.WaterSurfaceTension.input.Value
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call H2OSTCNV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If
             
             SourceOfValueToDisplay = phprop.WaterSurfaceTension.input.source.short
             frmWaterSurfaceTension!lblCurrentValues(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
             frmWaterSurfaceTension!lblCurrentValues(1).Caption = frmWaterSurfaceTension!lblSource(1).Caption
       End Select

       HaveProperty(WATER_SURFACE_TENSION) = True
       phprop.WaterSurfaceTension.CurrentSelection.choice = ValueToDisplayIndex
       phprop.WaterSurfaceTension.CurrentSelection.Value = SIValue
       phprop.WaterSurfaceTension.CurrentSelection.source = SourceOfValueToDisplay

       contam_prop_form!lblAirWaterProperties(2).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
    End If

End Sub

Sub DisplayWaterViscosity()
    Dim ValueToDisplayIndex As Integer
    Dim I As Integer
    Dim PropertySourceToHighlight As Integer
    Dim EnglishValue As Double, SIValue As Double
    Dim ValueToDisplay As Double


' ***** BEGIN Calculations to determine which value is displayed on
' ***** main screen according to hierarchy

    frmWaterViscosity!lblCurrentValues(0).Caption = ""
    frmWaterViscosity!lblCurrentValues(1).Caption = ""

    Call DisplayWaterViscosityMainScreen(ValueToDisplayIndex)

' ***** END Calculations to determine which value is displayed on
' ***** main screen according to hierarchy


' ***** BEGIN Displaying water viscosity values in water viscosity
' ***** form (frmWaterViscosity)

'   *** Initialize all text and label boxes on frmWaterViscosity to gray and empty
    For I = 0 To 0
        frmWaterViscosity!Option1(I + 1).BackColor = &HC0C0C0
        frmWaterViscosity!Option1(I + 1).Enabled = False
        frmWaterViscosity!Option1(I + 1).Value = False
        frmWaterViscosity!lblSource(I).BackColor = &HC0C0C0
        frmWaterViscosity!lblWaterViscosityValue(I).Caption = "Not Available"
'        frmWaterViscosity!lblWaterViscosityValue(I).Enabled = False
        frmWaterViscosity!lblWaterViscosityValue(I).BackColor = &HC0C0C0
        frmWaterViscosity!lblH2OViscosityTemperature(I).Caption = ""
        frmWaterViscosity!lblH2OViscosityTemperature(I).Enabled = False
        frmWaterViscosity!lblH2OViscosityTemperature(I).BackColor = &HC0C0C0
        frmWaterViscosity!lblH2OViscosityminimumT(I).Caption = ""
        frmWaterViscosity!lblH2OViscosityminimumT(I).Enabled = False
        frmWaterViscosity!lblH2OViscosityminimumT(I).BackColor = &HC0C0C0
        frmWaterViscosity!lblH2OViscositymaximumT(I).Caption = ""
        frmWaterViscosity!lblH2OViscositymaximumT(I).Enabled = False
        frmWaterViscosity!lblH2OViscositymaximumT(I).BackColor = &HC0C0C0
    Next I

        frmWaterViscosity!Option1(2).BackColor = &HC0C0C0
        frmWaterViscosity!Option1(2).Enabled = False
        frmWaterViscosity!Option1(2).Value = False
        frmWaterViscosity!lblSource(1).BackColor = &HC0C0C0
        frmWaterViscosity!txtWaterViscosityValue(1).Text = ""
        frmWaterViscosity!txtWaterViscosityValue(1).Enabled = False
        frmWaterViscosity!txtWaterViscosityValue(1).BackColor = &HC0C0C0
        frmWaterViscosity!txtH2OViscosityTemperature(1).Text = ""
        frmWaterViscosity!txtH2OViscosityTemperature(1).Enabled = False
        frmWaterViscosity!txtH2OViscosityTemperature(1).BackColor = &HC0C0C0
        frmWaterViscosity!txtH2OViscosityminimumT(1).Text = ""
        frmWaterViscosity!txtH2OViscosityminimumT(1).Enabled = False
        frmWaterViscosity!txtH2OViscosityminimumT(1).BackColor = &HC0C0C0
        frmWaterViscosity!txtH2OViscositymaximumT(1).Text = ""
        frmWaterViscosity!txtH2OViscositymaximumT(1).Enabled = False
        frmWaterViscosity!txtH2OViscositymaximumT(1).BackColor = &HC0C0C0

    If PROPAVAILABLE(WATER_VISCOSITY_CORRELATION) Then

       SIValue = phprop.WaterViscosity.correlation.Value
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call WVISCCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       frmWaterViscosity!lblWaterViscosityValue(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       SIValue = phprop.WaterViscosity.correlation.temperature
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call TEMPCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       frmWaterViscosity!lblH2OViscosityTemperature(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       If CurrentUnits = SIUnits Then
          frmWaterViscosity!lblH2OViscosityminimumT(0).Caption = "0.0"
          frmWaterViscosity!lblH2OViscositymaximumT(0).Caption = "370.0"
       Else
          frmWaterViscosity!lblH2OViscosityminimumT(0).Caption = "32.0"
          frmWaterViscosity!lblH2OViscositymaximumT(0).Caption = "698.0"
       End If
       '*** Set colors of available choices to white
       frmWaterViscosity!Option1(1).BackColor = &HFFFFFF
       frmWaterViscosity!Option1(1).Enabled = True
       frmWaterViscosity!lblSource(0).BackColor = &HFFFFFF
       frmWaterViscosity!lblWaterViscosityValue(0).Enabled = True
       frmWaterViscosity!lblWaterViscosityValue(0).BackColor = &HFFFFFF
       frmWaterViscosity!lblH2OViscosityTemperature(0).Enabled = True
       frmWaterViscosity!lblH2OViscosityTemperature(0).BackColor = &HFFFFFF
       frmWaterViscosity!lblH2OViscosityminimumT(0).Enabled = True
       frmWaterViscosity!lblH2OViscosityminimumT(0).BackColor = &HFFFFFF
       frmWaterViscosity!lblH2OViscositymaximumT(0).Enabled = True
       frmWaterViscosity!lblH2OViscositymaximumT(0).BackColor = &HFFFFFF
       
       If ValueToDisplayIndex = WATER_VISCOSITY_CORRELATION Then
          frmWaterViscosity!Option1(1).Value = True
          PropertySourceToHighlight = 0
       Else
          frmWaterViscosity!Option1(1).Value = False
       End If
    End If

'  *** User input always possible so set backcolor to white
    frmWaterViscosity!Option1(2).BackColor = &HFFFFFF
    frmWaterViscosity!Option1(2).Enabled = True
    frmWaterViscosity!lblSource(1).BackColor = &HFFFFFF
    frmWaterViscosity!txtWaterViscosityValue(1).Enabled = True
    frmWaterViscosity!txtWaterViscosityValue(1).BackColor = &HFFFFFF
    frmWaterViscosity!txtH2OViscosityTemperature(1).Enabled = True
    frmWaterViscosity!txtH2OViscosityTemperature(1).BackColor = &HFFFFFF
    frmWaterViscosity!txtH2OViscosityminimumT(1).Enabled = True
    frmWaterViscosity!txtH2OViscosityminimumT(1).BackColor = &HFFFFFF
    frmWaterViscosity!txtH2OViscositymaximumT(1).Enabled = True
    frmWaterViscosity!txtH2OViscositymaximumT(1).BackColor = &HFFFFFF

    If PROPAVAILABLE(WATER_VISCOSITY_INPUT) Then

       SIValue = phprop.WaterViscosity.input.Value
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call WVISCCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       frmWaterViscosity!txtWaterViscosityValue(1).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       If HaveTemp(phprop.WaterViscosity.input.temperature) Then
          SIValue = phprop.WaterViscosity.input.temperature
          If CurrentUnits = SIUnits Then
             ValueToDisplay = SIValue
          ElseIf CurrentUnits = EnglishUnits Then
             Call TEMPCNV(EnglishValue, SIValue)
             ValueToDisplay = EnglishValue
          End If
          frmWaterViscosity!txtH2OViscosityTemperature(1).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
       Else
          frmWaterViscosity!txtH2OViscosityTemperature(1).Text = ""
       End If
       frmWaterViscosity!txtH2OViscosityminimumT(1).Text = ""
       frmWaterViscosity!txtH2OViscositymaximumT(1).Text = ""

       If ValueToDisplayIndex = WATER_VISCOSITY_INPUT Then
          frmWaterViscosity!Option1(2).Value = True
          PropertySourceToHighlight = 1
       Else
          frmWaterViscosity!Option1(2).Value = False
       End If

    End If

       For I = 0 To 1
           frmWaterViscosity!lblSource(I).ForeColor = &H80000008
       Next I

    '*** Highlight selected property source
    If ValueToDisplayIndex <> 0 Then
       frmWaterViscosity!lblSource(PropertySourceToHighlight).BackColor = &H800000
       frmWaterViscosity!lblSource(PropertySourceToHighlight).ForeColor = &H80000005
       hilight.WaterViscosity.PreviousIndex = PropertySourceToHighlight
    End If


' ***** END Displaying water viscosity values in water viscosity
' ***** form (frmWaterViscosity)

End Sub

Sub DisplayWaterViscosityMainScreen(ValueToDisplayIndex As Integer)
    Dim DisplayedValueOnMainScreen As Integer
    Dim ValueToDisplay As Double
    Dim SourceOfValueToDisplay As Long
    Dim SIValue As Double, EnglishValue As Double

    If phprop.WaterViscosity.CurrentSelection.choice = 0 Then
       DisplayedValueOnMainScreen = False
       ValueToDisplayIndex = 0
    Else
       DisplayedValueOnMainScreen = True
       ValueToDisplayIndex = phprop.WaterViscosity.CurrentSelection.choice
    End If
    
    If Not DisplayedValueOnMainScreen Then
       Call CheckWaterViscosity(1, ValueToDisplayIndex, DisplayedValueOnMainScreen)
    End If

    If Not DisplayedValueOnMainScreen Then
       Call CheckWaterViscosity(2, ValueToDisplayIndex, DisplayedValueOnMainScreen)
    End If

    If Not DisplayedValueOnMainScreen Then
       contam_prop_form!lblAirWaterProperties(1).Caption = "Not Available"
       HaveProperty(WATER_VISCOSITY) = False
    Else
       Select Case ValueToDisplayIndex
          Case WATER_VISCOSITY_CORRELATION

             SIValue = phprop.WaterViscosity.correlation.Value
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call WVISCCNV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If
           
             SourceOfValueToDisplay = phprop.WaterViscosity.correlation.source.short
             frmWaterViscosity!lblCurrentValues(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
             frmWaterViscosity!lblCurrentValues(1).Caption = frmWaterViscosity!lblSource(0).Caption

          Case WATER_VISCOSITY_INPUT

             SIValue = phprop.WaterViscosity.input.Value
             If CurrentUnits = SIUnits Then
                ValueToDisplay = SIValue
             ElseIf CurrentUnits = EnglishUnits Then
                Call WVISCCNV(EnglishValue, SIValue)
                ValueToDisplay = EnglishValue
             End If
            
             SourceOfValueToDisplay = phprop.WaterViscosity.input.source.short
             frmWaterViscosity!lblCurrentValues(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
             frmWaterViscosity!lblCurrentValues(1).Caption = frmWaterViscosity!lblSource(1).Caption
       End Select

       HaveProperty(WATER_VISCOSITY) = True
       phprop.WaterViscosity.CurrentSelection.choice = ValueToDisplayIndex
       phprop.WaterViscosity.CurrentSelection.Value = SIValue
       phprop.WaterViscosity.CurrentSelection.source = SourceOfValueToDisplay

       contam_prop_form!lblAirWaterProperties(1).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
    End If

End Sub

Function GetBIP_DB_String(BinaryInteractionParameterDatabase As Long) As String

   Select Case BinaryInteractionParameterDatabase
      Case 1    'Original UNIFAC VLE
           GetBIP_DB_String = "Original UNIFAC VLE"
      Case 2    'UNIFAC LLE
           GetBIP_DB_String = "UNIFAC LLE"
      Case 3    'Environmental VLE
           GetBIP_DB_String = "Environmental VLE"
      Case 0    'UNIFAC calculation not possible
           GetBIP_DB_String = "UNIFAC calculation not possible"
   End Select

End Function

Function GetTheFormat(Value As Double) As String
   Dim AbsValue As Double

   AbsValue = Abs(Value)

   If AbsValue < 0.001 Then
      GetTheFormat = "0.00E+00"
   ElseIf AbsValue < 0.01 Then
      GetTheFormat = "0.00E+00"
   ElseIf AbsValue < 0.1 Then
      GetTheFormat = "0.0000"
   ElseIf AbsValue < 1 Then
      GetTheFormat = "0.000"
   ElseIf AbsValue < 10 Then
      GetTheFormat = "0.00"
   ElseIf AbsValue < 100 Then
      GetTheFormat = "0.0"
   ElseIf AbsValue < 1000 Then
      GetTheFormat = "0"
   Else
      GetTheFormat = "0.00E+00"
   End If

End Function

