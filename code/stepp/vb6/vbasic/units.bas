Attribute VB_Name = "UnitsMod"
'*** This module will store routines for switching between
'*** SI and English Units

Global Const SIUnits = 0
Global Const EnglishUnits = 1
Global CurrentUnits As Integer

'Array to hold current Units for each property
Global Units(1 To NUMBER_OF_PROPERTIES) As String

Sub BuildEnglishLabels()
'*** This subroutine will place labels appropriate to English Units
'*** on to the forms used in the program and place the units of each property into
'*** an array called UNITS

     '*** Main Form:  contam_prop_form
     contam_prop_form!lblOperatingConditions(0).Caption = "Pressure (psi)"
     contam_prop_form!lblOperatingConditions(1).Caption = "Temperature (F)"
     contam_prop_form.lblContaminantPropertiesLabel(0).Caption = "Vapor Pressure (psi)"
     contam_prop_form.lblContaminantPropertiesLabel(1).Caption = "Infinite Dilution Activity Coeff. (-)"
     contam_prop_form.lblContaminantPropertiesLabel(2).Caption = "Henry's Constant (-)"
     contam_prop_form.lblContaminantPropertiesLabel(3).Caption = "Molecular Weight (lb/lb-mol)"
     contam_prop_form.lblContaminantPropertiesLabel(4).Caption = "Normal Boiling Point (NBP) (F)"
     contam_prop_form.lblContaminantPropertiesLabel(5).Caption = "Liquid Density (lb/ft3)"
     contam_prop_form.lblContaminantPropertiesLabel(6).Caption = "Molar Volume at Op. T (ft3/lb-mol)"
     contam_prop_form.lblContaminantPropertiesLabel(7).Caption = "Molar Volume at NBP (ft3/lb-mol)"
     contam_prop_form.lblContaminantPropertiesLabel(8).Caption = "Refractive Index @77 F"
     contam_prop_form.lblContaminantPropertiesLabel(9).Caption = "Aqueous Solubility (PPMw)"
     contam_prop_form.lblContaminantPropertiesLabel(10).Caption = "log Octanol Water Part. Coeff. (-)"
     contam_prop_form.lblContaminantPropertiesLabel(11).Caption = "Liquid Diffusivity (ft2/s)"
     contam_prop_form.lblContaminantPropertiesLabel(12).Caption = "Gas Diffusivity (ft2/s)"
     contam_prop_form.lblAirWaterPropertiesLabel(0).Caption = "Water Density (lb/ft3)"
     contam_prop_form.lblAirWaterPropertiesLabel(1).Caption = "Water Viscosity (lb/ft/s)"
     contam_prop_form.lblAirWaterPropertiesLabel(2).Caption = "Water Surface Tension (lbf/ft)"
     contam_prop_form.lblAirWaterPropertiesLabel(3).Caption = "Air Density (lb/ft3)"
     contam_prop_form.lblAirWaterPropertiesLabel(4).Caption = "Air Viscosity (lb/ft/s)"

     '*** Vapor Pressure Form (vp_form)
     vp_form!lblCurrentInformation(0).Caption = "Value (psi)"
     vp_form!lblVPLabel.Caption = "Vapor Pressure (psi)"
     vp_form!lblVPTempLabel.Caption = "Temp. (F)"
     vp_form!lblVPminTLabel.Caption = "Tmin (F)"
     vp_form!lblVPmaxTLabel.Caption = "Tmax (F)"

     '*** Infinite Dilution Activity Coefficient form (vp_form)
     infinite_dilution_form!lblCurrentInformation(0).Caption = "Value (-)"
     infinite_dilution_form!lblACLabel.Caption = "Activity Coefficient (-)"
     infinite_dilution_form!lblACTempLabel.Caption = "Temp. (F)"

     '*** Henry's Constant form (hc_form)
     hc_form!lblCurrentInformation(0).Caption = "Value (-)"
     hc_form!lblHCLabel.Caption = "Henry's Constant (-)"
     hc_form!lblHCTempLabel.Caption = "Temp. (F)"
     hc_form!lblHCminTLabel.Caption = "Tmin (F)"
     hc_form!lblHCmaxTLabel.Caption = "Tmax (F)"

     '*** Molecular Weight Form
     mwt_form!lblCurrentInformation(0).Caption = "Value (lb/lb-mol)"
     mwt_form!lblMWTLabel.Caption = "Molecular Weight (lb/lb-mol)"

     '*** Normal Boiling Point form (nbp_form)
     nbp_form!lblCurrentInformation(0).Caption = "Value (F)"
     nbp_form!lblNBPLabel.Caption = "Normal Boiling Point (F)"

     '*** Liquid Density form (ldens_form)
     ldens_form!lblCurrentInformation(0).Caption = "Value (lb/ft3)"
     ldens_form!lblLDensLabel.Caption = "Liq. Dens (lb/ft3)"
     ldens_form!lblLDensTempLabel.Caption = "Temp (F)"
     ldens_form!lblLDensminTLabel.Caption = "Tmin (F)"
     ldens_form!lblLDensmaxTLabel.Caption = "Tmax (F)"

     '*** Molar Volume at Operating Temperature form (molar_vol_form)
     molar_vol_form!lblCurrentInformation(0).Caption = "Value (ft3/lb-mol)"
     molar_vol_form!lblMVOpTLabel.Caption = "Molar Vol." + Chr$(13) + "(ft3/lb-mol)"
     molar_vol_form!lblMVOpTTempLabel.Caption = "Temp. (F)"
     molar_vol_form!lblMVOpTminTLabel.Caption = "Tmin (F)"
     molar_vol_form!lblMVOpTMaxTLabel.Caption = "Tmax (F)"

     '*** Molar Volume at Normal Boiling Point form (mv_nbp_form)
     mv_nbp_form!lblCurrentInformation(0).Caption = "Value (ft3/lb-mol)"
     mv_nbp_form!lblMVNBPLabel.Caption = "Molar Vol. (ft3/lb-mol)"
     mv_nbp_form!lblMVNBPTempLabel = "Temp. (F)"

     '*** Refractive Index form (rindex_form)
     rindex_form!lblCurrentInformation(0).Caption = "Value (-)"
     rindex_form!lblRefIndexLabel.Caption = "Refractive Index @77 F"

     '*** Aqueous Solubility Form (aqsol_form)
     aqsol_form!lblCurrentInformation(0).Caption = "Value (PPMw)"
     aqsol_form!lblAqSolLabel.Caption = "Aqueous Sol. (PPMw)"
     aqsol_form!lblAqSolTempLabel.Caption = "Temp. (F)"

    '*** Octanol Water Partition Coefficient form (octanol_form)
    octanol_form!lblCurrentInformation(0).Caption = "Value of log Kow (-)"
    octanol_form!lblKowLabel.Caption = "log Kow (-)"
    octanol_form!lblKowTempLabel.Caption = "Temp. (F)"

    '*** Liquid Diffusivity Form (liquid_diff_form)
    liquid_diff_form!lblCurrentInformation(0).Caption = "Value (ft2/s)"
    liquid_diff_form!lblLiqDiffLabel.Caption = "Liquid Diff. (ft2/s)"
    liquid_diff_form!lblLiqDiffTempLabel.Caption = "Temp. (F)"

    '*** Gas Diffusivity Form (gas_diff_form)
    gas_diff_form!lblCurrentInformation(0).Caption = "Value (ft2/s)"
    gas_diff_form!lblGasDiffLabel.Caption = "Gas Diffusivity" + Chr$(13) + "(ft2/s)"
    gas_diff_form!lblGasDiffTempLabel.Caption = "Temp. (F)"

    '*** Water Density form (frmWaterDensity)
    frmWaterDensity!lblCurrentInformation(0).Caption = "Value (lb/ft3)"
    frmWaterDensity!lblH2ODensLabel.Caption = "H2O Dens. (lb/ft3)"
    frmWaterDensity!lblH2ODensTempLabel.Caption = "Temp. (F)"
    frmWaterDensity!lblH2ODensminTLabel.Caption = "Tmin (F)"
    frmWaterDensity!lblH2ODensmaxTLabel.Caption = "Tmax (F)"

    '*** Water Viscosity form (frmWaterViscosity)
    frmWaterViscosity!lblCurrentInformation(0).Caption = "Value (lb/ft/s)"
    frmWaterViscosity!lblH2OViscLabel.Caption = "H2O Visc." + Chr$(13) + "(lb/ft/s)"
    frmWaterViscosity!lblH2OViscTempLabel.Caption = "Temp. (F)"
    frmWaterViscosity!lblH2OViscminTLabel.Caption = "Tmin (F)"
    frmWaterViscosity!lblH2OViscmaxTLabel.Caption = "Tmax (F)"

    '*** Water Surface Tension form (frmWaterSurfaceTension)
    frmWaterSurfaceTension!lblCurrentInformation(0).Caption = "Value (lbf/ft)"
    frmWaterSurfaceTension!lblH2OSTLabel.Caption = "Surf. Tens." + Chr$(13) + "(lbf/ft)"
    frmWaterSurfaceTension!lblH2OSTTempLabel.Caption = "Temp. (F)"
    frmWaterSurfaceTension!lblH2OSTminTLabel.Caption = "Tmin (F)"
    frmWaterSurfaceTension!lblH2OSTmaxTLabel.Caption = "Tmax (F)"

    '*** Air Density form (frmAirDensity)
    frmAirDensity!lblCurrentInformation(0).Caption = "Value (lb/ft3)"
    frmAirDensity!lblAirDensLabel.Caption = "Air Dens." + Chr$(13) + "(lb/ft3)"
    frmAirDensity!lblAirDensTempLabel.Caption = "Temp. (F)"
    frmAirDensity!lblAirDensminTLabel.Caption = "Tmin (F)"
    frmAirDensity!lblAirDensmaxTLabel.Caption = "Tmax (F)"

    '*** Air Viscosity form (frmAirViscosity)
    frmAirViscosity!lblCurrentInformation(0).Caption = "Value (lb/ft/s)"
    frmAirViscosity!lblAirViscLabel.Caption = "Air Visc." + Chr$(13) + "(lb/ft/s)"
    frmAirViscosity!lblAirViscTempLabel.Caption = "Temp. (F)"
    frmAirViscosity!lblAirViscminTLabel.Caption = "Tmin (F)"
    frmAirViscosity!lblAirViscmaxTLabel.Caption = "Tmax (F)"

    Call CreateUnitsArrayEnglish

End Sub

Sub BuildSILabels()

'*** This subroutine will place labels on all the forms
'*** corresponding to SI Units

     '*** Main Form:  contam_prop_form
     contam_prop_form!lblOperatingConditions(0).Caption = "Pressure (Pa)"
     contam_prop_form!lblOperatingConditions(1).Caption = "Temperature (C)"
     contam_prop_form.lblContaminantPropertiesLabel(0).Caption = "Vapor Pressure (Pa)"
     contam_prop_form.lblContaminantPropertiesLabel(1).Caption = "Infinite Dilution Activity Coeff. (-)"
     contam_prop_form.lblContaminantPropertiesLabel(2).Caption = "Henry's Constant (-)"
     contam_prop_form.lblContaminantPropertiesLabel(3).Caption = "Molecular Weight (kg/kmol)"
     contam_prop_form.lblContaminantPropertiesLabel(4).Caption = "Normal Boiling Point (NBP) (C)"
     contam_prop_form.lblContaminantPropertiesLabel(5).Caption = "Liquid Density (kg/m3)"
     contam_prop_form.lblContaminantPropertiesLabel(6).Caption = "Molar Volume at Op. T (m3/kmol)"
     contam_prop_form.lblContaminantPropertiesLabel(7).Caption = "Molar Volume at NBP (m3/kmol)"
     contam_prop_form.lblContaminantPropertiesLabel(8).Caption = "Refractive Index @25 C"
     contam_prop_form.lblContaminantPropertiesLabel(9).Caption = "Aqueous Solubility (PPMw)"
     contam_prop_form.lblContaminantPropertiesLabel(10).Caption = "log Octanol Water Part. Coeff. (-)"
     contam_prop_form.lblContaminantPropertiesLabel(11).Caption = "Liquid Diffusivity (m2/s)"
     contam_prop_form.lblContaminantPropertiesLabel(12).Caption = "Gas Diffusivity (m2/s)"
     contam_prop_form.lblAirWaterPropertiesLabel(0).Caption = "Water Density (kg/m3)"
     contam_prop_form.lblAirWaterPropertiesLabel(1).Caption = "Water Viscosity (kg/m/s)"
     contam_prop_form.lblAirWaterPropertiesLabel(2).Caption = "Water Surface Tension (N/m)"
     contam_prop_form.lblAirWaterPropertiesLabel(3).Caption = "Air Density (kg/m3)"
     contam_prop_form.lblAirWaterPropertiesLabel(4).Caption = "Air Viscosity (kg/m/s)"

     '*** Vapor Pressure Form (vp_form)
     vp_form!lblCurrentInformation(0).Caption = "Value (Pa)"
     vp_form!lblVPLabel.Caption = "Vapor Pressure (Pa)"
     vp_form!lblVPTempLabel.Caption = "Temp. (C)"
     vp_form!lblVPminTLabel.Caption = "Tmin (C)"
     vp_form!lblVPmaxTLabel.Caption = "Tmax (C)"

     '*** Infinite Dilution Activity Coefficient form (vp_form)
     infinite_dilution_form!lblCurrentInformation(0).Caption = "Value (-)"
     infinite_dilution_form!lblACLabel.Caption = "Activity Coefficient (-)"
     infinite_dilution_form!lblACTempLabel.Caption = "Temp. (C)"

     '*** Henry's Constant form (hc_form)
     hc_form!lblCurrentInformation(0).Caption = "Value (-)"
     hc_form!lblHCLabel.Caption = "Henry's Constant (-)"
     hc_form!lblHCTempLabel.Caption = "Temp. (C)"
     hc_form!lblHCminTLabel.Caption = "Tmin (C)"
     hc_form!lblHCmaxTLabel.Caption = "Tmax (C)"

     '*** Molecular Weight Form
     mwt_form!lblCurrentInformation(0).Caption = "Value (kg/kmol)"
     mwt_form!lblMWTLabel.Caption = "Molecular Weight (kg/kmol)"

     '*** Normal Boiling Point form (nbp_form)
     nbp_form!lblCurrentInformation(0).Caption = "Value (C)"
     nbp_form!lblNBPLabel.Caption = "Normal Boiling Point (C)"

     '*** Liquid Density form (ldens_form)
     ldens_form!lblCurrentInformation(0).Caption = "Value (kg/m3)"
     ldens_form!lblLDensLabel.Caption = "Liq. Dens (kg/m3)"
     ldens_form!lblLDensTempLabel.Caption = "Temp (C)"
     ldens_form!lblLDensminTLabel.Caption = "Tmin (C)"
     ldens_form!lblLDensmaxTLabel.Caption = "Tmax (C)"

     '*** Molar Volume at Operating Temperature form (molar_vol_form)
     molar_vol_form!lblCurrentInformation(0).Caption = "Value (m3/kmol)"
     molar_vol_form!lblMVOpTLabel.Caption = "Molar Vol." + Chr$(13) + "(m3/kmol)"
     molar_vol_form!lblMVOpTTempLabel.Caption = "Temp. (C)"
     molar_vol_form!lblMVOpTminTLabel.Caption = "Tmin (C)"
     molar_vol_form!lblMVOpTMaxTLabel.Caption = "Tmax (C)"

     '*** Molar Volume at Normal Boiling Point form (mv_nbp_form)
     mv_nbp_form!lblCurrentInformation(0).Caption = "Value (m3/kmol)"
     mv_nbp_form!lblMVNBPLabel.Caption = "Molar Vol. (m3/kmol)"
     mv_nbp_form!lblMVNBPTempLabel = "Temp. (C)"

     '*** Refractive Index form (rindex_form)
     rindex_form!lblCurrentInformation(0).Caption = "Value (-)"
     rindex_form!lblRefIndexLabel.Caption = "Refractive Index @25 C"

     '*** Aqueous Solubility Form (aqsol_form)
     aqsol_form!lblCurrentInformation(0).Caption = "Value (PPMw)"
     aqsol_form!lblAqSolLabel.Caption = "Aqueous Sol. (PPMw)"
     aqsol_form!lblAqSolTempLabel.Caption = "Temp. (C)"

    '*** Octanol Water Partition Coefficient form (octanol_form)
    octanol_form!lblCurrentInformation(0).Caption = "Value of log Kow (-)"
    octanol_form!lblKowLabel.Caption = "log Kow (-)"
    octanol_form!lblKowTempLabel.Caption = "Temp. (C)"

    '*** Liquid Diffusivity Form (liquid_diff_form)
    liquid_diff_form!lblCurrentInformation(0).Caption = "Value (m2/s)"
    liquid_diff_form!lblLiqDiffLabel.Caption = "Liquid Diff. (m2/s)"
    liquid_diff_form!lblLiqDiffTempLabel.Caption = "Temp. (C)"

    '*** Gas Diffusivity Form (gas_diff_form)
    gas_diff_form!lblCurrentInformation(0).Caption = "Value (m2/s)"
    gas_diff_form!lblGasDiffLabel.Caption = "Gas Diffusivity" + Chr$(13) + "(m2/s)"
    gas_diff_form!lblGasDiffTempLabel.Caption = "Temp. (C)"

    '*** Water Density form (frmWaterDensity)
    frmWaterDensity!lblCurrentInformation(0).Caption = "Value (kg/m3)"
    frmWaterDensity!lblH2ODensLabel.Caption = "H2O Dens. (kg/m3)"
    frmWaterDensity!lblH2ODensTempLabel.Caption = "Temp. (C)"
    frmWaterDensity!lblH2ODensminTLabel.Caption = "Tmin (C)"
    frmWaterDensity!lblH2ODensmaxTLabel.Caption = "Tmax (C)"

    '*** Water Viscosity form (frmWaterViscosity)
    frmWaterViscosity!lblCurrentInformation(0).Caption = "Value (kg/m/s)"
    frmWaterViscosity!lblH2OViscLabel.Caption = "H2O Visc." + Chr$(13) + "(kg/m/s)"
    frmWaterViscosity!lblH2OViscTempLabel.Caption = "Temp. (C)"
    frmWaterViscosity!lblH2OViscminTLabel.Caption = "Tmin (C)"
    frmWaterViscosity!lblH2OViscmaxTLabel.Caption = "Tmax (C)"

    '*** Water Surface Tension form (frmWaterSurfaceTension)
    frmWaterSurfaceTension!lblCurrentInformation(0).Caption = "Value (N/m)"
    frmWaterSurfaceTension!lblH2OSTLabel.Caption = "Surf. Tens." + Chr$(13) + "(N/m)"
    frmWaterSurfaceTension!lblH2OSTTempLabel.Caption = "Temp. (C)"
    frmWaterSurfaceTension!lblH2OSTminTLabel.Caption = "Tmin (C)"
    frmWaterSurfaceTension!lblH2OSTmaxTLabel.Caption = "Tmax (C)"

    '*** Air Density form (frmAirDensity)
    frmAirDensity!lblCurrentInformation(0).Caption = "Value (kg/m3)"
    frmAirDensity!lblAirDensLabel.Caption = "Air Dens." + Chr$(13) + "(kg/m3)"
    frmAirDensity!lblAirDensTempLabel.Caption = "Temp. (C)"
    frmAirDensity!lblAirDensminTLabel.Caption = "Tmin (C)"
    frmAirDensity!lblAirDensmaxTLabel.Caption = "Tmax (C)"

    '*** Air Viscosity form (frmAirViscosity)
    frmAirViscosity!lblCurrentInformation(0).Caption = "Value (kg/m/s)"
    frmAirViscosity!lblAirViscLabel.Caption = "Air Visc." + Chr$(13) + "(kg/m/s)"
    frmAirViscosity!lblAirViscTempLabel.Caption = "Temp. (C)"
    frmAirViscosity!lblAirViscminTLabel.Caption = "Tmin (C)"
    frmAirViscosity!lblAirViscmaxTLabel.Caption = "Tmax (C)"

    Call CreateUnitsArraySI

End Sub

Sub CreateUnitsArrayEnglish()

'Place the Units into an array called units
    Units(OPERATING_PRESSURE) = "psi"
    Units(OPERATING_TEMPERATURE) = "F"
    Units(VAPOR_PRESSURE) = "psi"
    Units(ACTIVITY_COEFFICIENT) = "(-)"
    Units(HENRYS_CONSTANT) = "(-)"
    Units(MOLECULAR_WEIGHT) = "lbm/lbm-mol"
    Units(BOILING_POINT) = "F"
    Units(LIQUID_DENSITY) = "lbm/ft3"
    Units(MOLAR_VOLUME_BOILING_POINT) = "ft3/lbm-mol"
    Units(MOLAR_VOLUME_OPT) = "ft3/lbm-mol"
    Units(REFRACTIVE_INDEX) = "(-)"
    Units(AQUEOUS_SOLUBILITY) = "PPMw"
    Units(OCT_WATER_PART_COEFF) = "(-)"
    Units(LIQUID_DIFFUSIVITY) = "ft2/sec"
    Units(GAS_DIFFUSIVITY) = "ft2/sec"
    Units(WATER_DENSITY) = "lbm/ft3"
    Units(WATER_VISCOSITY) = "lbm/ft/sec"
    Units(WATER_SURFACE_TENSION) = "lbf/ft"
    Units(AIR_DENSITY) = "lbm/ft3"
    Units(AIR_VISCOSITY) = "lbm/ft/sec"

End Sub

Sub CreateUnitsArraySI()

'Place the Units into an array called units
    Units(OPERATING_PRESSURE) = "Pa"
    Units(OPERATING_TEMPERATURE) = "C"
    Units(VAPOR_PRESSURE) = "Pa"
    Units(ACTIVITY_COEFFICIENT) = "(-)"
    Units(HENRYS_CONSTANT) = "(-)"
    Units(MOLECULAR_WEIGHT) = "kg/kmol"
    Units(BOILING_POINT) = "C"
    Units(LIQUID_DENSITY) = "kg/m3"
    Units(MOLAR_VOLUME_BOILING_POINT) = "m3/kmol"
    Units(MOLAR_VOLUME_OPT) = "m3/kmol"
    Units(REFRACTIVE_INDEX) = "(-)"
    Units(AQUEOUS_SOLUBILITY) = "PPMw"
    Units(OCT_WATER_PART_COEFF) = "(-)"
    Units(LIQUID_DIFFUSIVITY) = "m2/sec"
    Units(GAS_DIFFUSIVITY) = "m2/sec"
    Units(WATER_DENSITY) = "kg/m3"
    Units(WATER_VISCOSITY) = "kg/m/sec"
    Units(WATER_SURFACE_TENSION) = "N/m"
    Units(AIR_DENSITY) = "kg/m3"
    Units(AIR_VISCOSITY) = "kg/m/sec"

End Sub

Sub GetUnits()

    '*** Place appropriate labels on to forms
    If CurrentUnits = SIUnits Then
       Call BuildSILabels
    ElseIf CurrentUnits = EnglishUnits Then
       Call BuildEnglishLabels
    End If

End Sub

