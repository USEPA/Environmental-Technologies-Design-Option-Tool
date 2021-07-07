Attribute VB_Name = "Refresh"
Option Explicit

Global ThisGraph As Object


Const Refresh_declarations_end = True


Sub frmIonExchangeMain_Repopulate_Values()
Dim Frm As Form
Set Frm = frmIonExchangeMain

  Call unitsys_set_number_in_base_units(frmIonExchangeMain.txtOperatingConditions(0), NowProj.Operating.Pressure)
  Call unitsys_set_number_in_base_units(frmIonExchangeMain.txtOperatingConditions(1), NowProj.Operating.Temperature)
  Call unitsys_set_number_in_base_units(frmIonExchangeMain.txtAdsorbentProperties(1), NowProj.Resin.ApparentDensity)
  Call unitsys_set_number_in_base_units(frmIonExchangeMain.txtAdsorbentProperties(2), NowProj.Resin.ParticleRadius)
  Call unitsys_set_number_in_base_units(frmIonExchangeMain.txtAdsorbentProperties(5), NowProj.Resin.TotalCapacity)
  Call unitsys_set_number_in_base_units(frmIonExchangeMain.txtBedData(0), NowProj.Bed.length)
  Call unitsys_set_number_in_base_units(frmIonExchangeMain.txtBedData(1), NowProj.Bed.Diameter)
  Call unitsys_set_number_in_base_units(frmIonExchangeMain.txtBedData(2), NowProj.Bed.Weight)
  Call unitsys_set_number_in_base_units(frmIonExchangeMain.txtBedData(3), NowProj.Bed.Flowrate.Value)
  Call unitsys_set_number_in_base_units(frmIonExchangeMain.txtBedData(4), NowProj.Bed.EBCT.Value)
End Sub
Sub frmIonExchangeMain_Refresh()
Dim i As Integer

  Call Populate_frmIonExchangeMain_Units
  Call frmIonExchangeMain_Repopulate_Values
  Call AssignTextAndTag(frmIonExchangeMain.txtOperatingConditions(0), _
    NowProj.Operating.Pressure)
  Call AssignTextAndTag(frmIonExchangeMain.txtOperatingConditions(1), _
    NowProj.Operating.Temperature)
  Call AssignTextAndTag(frmIonExchangeMain.txtBedData(0), _
    NowProj.Bed.length)
  Call AssignTextAndTag(frmIonExchangeMain.txtBedData(1), _
    NowProj.Bed.Diameter)
  Call AssignTextAndTag(frmIonExchangeMain.txtBedData(2), _
    NowProj.Bed.Weight)
  Call AssignTextAndTag(frmIonExchangeMain.txtBedData(3), _
    NowProj.Bed.Flowrate.Value)
  Call AssignTextAndTag(frmIonExchangeMain.txtBedData(4), _
    NowProj.Bed.EBCT.Value)
  
  Call AssignTextAndTag(frmIonExchangeMain.txtAdsorbentProperties(1), _
    NowProj.Resin.ApparentDensity)
  Call AssignTextAndTag(frmIonExchangeMain.txtAdsorbentProperties(2), _
    NowProj.Resin.ParticleRadius)
  Call AssignTextAndTag(frmIonExchangeMain.txtAdsorbentProperties(3), _
    NowProj.Resin.ParticlePorosity)
  Call AssignTextAndTag(frmIonExchangeMain.txtAdsorbentProperties(4), _
    NowProj.Resin.Tortuosity)
  Call AssignTextAndTag(frmIonExchangeMain.txtAdsorbentProperties(5), _
    NowProj.Resin.TotalCapacity)
  
  frmIonExchangeMain!cboIons(0).Clear
  frmIonExchangeMain!cboIons(2).Clear
  frmIonExchangeMain!lstIons(0).Clear
  frmInputKineticParameters!cboIon.Clear
  For i = 1 To NowProj.NumberOfCations
    If i <> NowProj.PresaturantCation Then
        frmIonExchangeMain!lstIons(0).AddItem NowProj.Cation(i).Name
    End If
    frmIonExchangeMain!cboIons(0).AddItem NowProj.Cation(i).Name
    frmIonExchangeMain!cboIons(2).AddItem NowProj.Cation(i).Name
    frmInputKineticParameters!cboIon.AddItem NowProj.Cation(i).Name
    Call AssignTextAndTag(frmAddComponent.txtAddIon(1), _
        NowProj.Cation(i).MolecularWeight)
  Next i
  'this activates the click event for cboIons(0)
  frmIonExchangeMain!cboIons(0).ListIndex = NowProj.PresaturantCation - 1

  frmIonExchangeMain!cboIons(1).Clear
  frmIonExchangeMain!lstIons(1).Clear
  For i = 1 To NowProj.NumberOfAnions
    If i <> NowProj.PresaturantAnion Then
        frmIonExchangeMain!lstIons(1).AddItem NowProj.Anion(i).Name
    End If
    frmIonExchangeMain!cboIons(1).AddItem NowProj.Anion(i).Name
    frmIonExchangeMain!cboIons(2).AddItem NowProj.Anion(i).Name
    frmInputKineticParameters!cboIon.AddItem NowProj.Anion(i).Name
    Call AssignTextAndTag(frmAddComponent.txtAddIon(1), _
        NowProj.Anion(i).MolecularWeight)
  Next i
  frmIonExchangeMain!cboIons(1).ListIndex = NowProj.PresaturantAnion - 1

  Call CalculateLiquidDensity
  Call CalculateLiquidViscosity
  Call CalculateBedArea
  Call CalculateBedVolume
  Call CalculateBedDensity
  Call CalculateBedPorosity
  Call CalculateSuperficialVelocity
  Call CalculateInterstitialVelocity
  Call CalculateEffectiveContactTime
  Call CalculateParticleDiameter
  Call UpdateKineticParametersAllIons

End Sub

Sub Populate_frmIonExchangeMain_Units()
  Call unitsys_register(frmIonExchangeMain, _
    frmIonExchangeMain!lblOperatingConditions(0), _
    frmIonExchangeMain!txtOperatingConditions(0), _
    frmIonExchangeMain!cboOperatingConditionsUnits(0), "pressure", _
     "Pa", "Pa", "", "", 100#, True)
  Call unitsys_register(frmIonExchangeMain, _
    frmIonExchangeMain!lblOperatingConditions(1), _
    frmIonExchangeMain!txtOperatingConditions(1), _
    frmIonExchangeMain!cboOperatingConditionsUnits(1), "temperature", _
      "k", "k", "", "", 100#, True)
  Call unitsys_register(frmIonExchangeMain, _
    frmIonExchangeMain!lblAdsorbentProperties(1), _
    frmIonExchangeMain!txtAdsorbentProperties(1), _
    frmIonExchangeMain!cboAdsorbentPropertyUnits(1), "density", _
      "g/ml", "g/ml", "", "", 100#, True)
  Call unitsys_register(frmIonExchangeMain, _
    frmIonExchangeMain!lblAdsorbentProperties(2), _
    frmIonExchangeMain!txtAdsorbentProperties(2), _
    frmIonExchangeMain!cboAdsorbentPropertyUnits(2), "length", _
      "m", "m", "", "", 100#, True)
  Call unitsys_register(frmIonExchangeMain, _
    frmIonExchangeMain!lblAdsorbentProperties(5), _
    frmIonExchangeMain!txtAdsorbentProperties(5), _
    frmIonExchangeMain!cboAdsorbentPropertyUnits(5), "resin_capacity", _
      "meq/g", "meq/g", "", "", 100#, True)
  Call unitsys_register(frmIonExchangeMain, _
    frmIonExchangeMain!lblBedData(0), _
    frmIonExchangeMain!txtBedData(0), _
    frmIonExchangeMain!cboBedDataUnits(0), "length", _
      "m", "m", "", "", 100#, True)
  Call unitsys_register(frmIonExchangeMain, _
    frmIonExchangeMain!lblBedData(1), _
    frmIonExchangeMain!txtBedData(1), _
    frmIonExchangeMain!cboBedDataUnits(1), "length", _
      "m", "m", "", "", 100#, True)
  Call unitsys_register(frmIonExchangeMain, _
    frmIonExchangeMain!lblBedData(2), _
    frmIonExchangeMain!txtBedData(2), _
    frmIonExchangeMain!cboBedDataUnits(2), "mass", _
      "kg", "kg", "", "", 100#, True)
  Call unitsys_register(frmIonExchangeMain, _
    frmIonExchangeMain!lblBedData(3), _
    frmIonExchangeMain!txtBedData(3), _
    frmIonExchangeMain!cboBedDataUnits(3), "flow_volumetric", _
      "m³/s", "m³/s", "", "", 100#, True)
  Call unitsys_register(frmIonExchangeMain, _
    frmIonExchangeMain!lblBedData(4), _
    frmIonExchangeMain!txtBedData(4), _
    frmIonExchangeMain!cboBedDataUnits(4), "time", _
      "min", "min", "", "", 100#, True)
      
'  Call unitsys_register(frmIonExchangeMain, lbldesc(3), txtData(3), cboUnits(3), "time", _
'      "min", "min", "", "", 100#, True)

End Sub


Sub Populate_frmAddComponent_Units()

  Call unitsys_register(frmAddComponent, _
    frmAddComponent!lblAddIon(4), _
    frmAddComponent!txtAddIon(1), _
    frmAddComponent!cboAddIonUnits(0), "molecular_weight", _
    "mg/mmol", "mg/mmol", "", "", 100#, True)
  Call unitsys_register(frmAddComponent, _
    frmAddComponent!lblAddIon(5), _
    frmAddComponent!txtAddIon(2), _
    frmAddComponent!cboAddIonUnits(1), "concentration", _
    "mg/L", "mg/L", "", "", 100#, True)
    
End Sub

Sub frmInputKineticParameters_Refresh()
Dim i As Integer

  Call Populate_frmInputKineticParameters_Units
  Call frmInputKineticParameters_Repopulate_Values
  For i = 1 To NowProj.NumberOfCations
    Call AssignTextAndTag(frmInputKineticParameters.txtLiquidDiffCorrelation, _
      NowProj.Cation(i).Kinetic.LiquidDiffusivityCorrelation)
    Call AssignTextAndTag(frmInputKineticParameters.txtLiquidDiffUserInput, _
      NowProj.Cation(i).Kinetic.LiquidDiffusivityUserInput)
    Call AssignTextAndTag(frmInputKineticParameters.txtIonicTransportCoeffCorr, _
      NowProj.Cation(i).Kinetic.IonicTransportCoeffCorrelation)
    Call AssignTextAndTag(frmInputKineticParameters.txtIonicTransCoeffUser, _
      NowProj.Cation(i).Kinetic.IonicTransportCoeffUserInput)
    Call AssignTextAndTag(frmInputKineticParameters.txtPoreDiffusivityCorr, _
      NowProj.Cation(i).Kinetic.PoreDiffusivityCorrelation)
    Call AssignTextAndTag(frmInputKineticParameters.txtPoreDiffusivityUser, _
      NowProj.Cation(i).Kinetic.PoreDiffusivityUserInput)
  Next i
  For i = 1 To NowProj.NumberOfAnions
    Call AssignTextAndTag(frmInputKineticParameters.txtLiquidDiffCorrelation, _
      NowProj.Anion(i).Kinetic.LiquidDiffusivityCorrelation)
    Call AssignTextAndTag(frmInputKineticParameters.txtLiquidDiffUserInput, _
      NowProj.Anion(i).Kinetic.LiquidDiffusivityUserInput)
    Call AssignTextAndTag(frmInputKineticParameters.txtIonicTransportCoeffCorr, _
      NowProj.Anion(i).Kinetic.IonicTransportCoeffCorrelation)
    Call AssignTextAndTag(frmInputKineticParameters.txtIonicTransCoeffUser, _
      NowProj.Anion(i).Kinetic.IonicTransportCoeffUserInput)
    Call AssignTextAndTag(frmInputKineticParameters.txtPoreDiffusivityCorr, _
      NowProj.Anion(i).Kinetic.PoreDiffusivityCorrelation)
    Call AssignTextAndTag(frmInputKineticParameters.txtPoreDiffusivityUser, _
      NowProj.Anion(i).Kinetic.PoreDiffusivityUserInput)
  Next i
  
  If NowProj.IonicTransportCoeffCorrName = IONIC_TRANSPORT_COEFFICIENT_1 Then
     frmInputKineticParameters!cboIonicTransport.ListIndex = 0
  ElseIf NowProj.IonicTransportCoeffCorrName = IONIC_TRANSPORT_COEFFICIENT_2 Then
     frmInputKineticParameters!cboIonicTransport.ListIndex = 1
  End If

  frmInputKineticParameters!cboIon.ListIndex = frmIonExchangeMain.cboKinDimComponent.ListIndex

End Sub

Sub Populate_frmInputKineticParameters_Units()
Dim H As Integer

  Call unitsys_register(frmInputKineticParameters, _
    frmInputKineticParameters!lblLiquidDiffCorrelation, _
    frmInputKineticParameters!txtLiquidDiffCorrelation, _
    frmInputKineticParameters!cboLiquidDiffusivityUnits(0), _
    "diffusivity", "cm²/s", "cm²/s", "", "", 100#, True)
  Call unitsys_register(frmInputKineticParameters, _
    frmInputKineticParameters!txtLiquidDiffUserInput, _
    frmInputKineticParameters!txtLiquidDiffUserInput, _
    frmInputKineticParameters!cboLiquidDiffusivityUnits(1), _
    "diffusivity", "cm²/s", "cm²/s", "", "", 100#, True)
    
  Call unitsys_register(frmInputKineticParameters, _
    frmInputKineticParameters!lblIonicTransportCoeffCorr, _
    frmInputKineticParameters!txtIonicTransportCoeffCorr, _
    frmInputKineticParameters!cboIonicTransportUnits(0), _
    "diffusivity", "cm²/s", "cm²/s", "", "", 100#, True)
  Call unitsys_register(frmInputKineticParameters, _
    frmInputKineticParameters!lblIonicTransCoeffUser, _
    frmInputKineticParameters!txtIonicTransCoeffUser, _
    frmInputKineticParameters!cboIonicTransportUnits(1), _
    "diffusivity", "cm²/s", "cm²/s", "", "", 100#, True)
  Call unitsys_register(frmInputKineticParameters, _
    frmInputKineticParameters!lblPoreDiffusivityCorr, _
    frmInputKineticParameters!txtPoreDiffusivityCorr, _
    frmInputKineticParameters!cboPoreDiffusivityUnits(0), _
    "diffusivity", "cm²/s", "cm²/s", "", "", 100#, True)
  Call unitsys_register(frmInputKineticParameters, _
    frmInputKineticParameters!lblPoreDiffusivityUser, _
    frmInputKineticParameters!txtPoreDiffusivityUser, _
    frmInputKineticParameters!cboPoreDiffusivityUnits(1), _
    "diffusivity", "cm²/s", "cm²/s", "", "", 100#, True)
  
    
End Sub

Sub frmInputKineticParameters_Repopulate_Values()

Dim Frm As Form
Dim i As Integer
Set Frm = frmInputKineticParameters

  For i = 1 To NowProj.NumberOfCations
    Call unitsys_set_number_in_base_units(frmInputKineticParameters.txtLiquidDiffCorrelation, _
      NowProj.Cation(i).Kinetic.LiquidDiffusivityCorrelation)
    Call unitsys_set_number_in_base_units(frmInputKineticParameters.txtLiquidDiffUserInput, _
      NowProj.Cation(i).Kinetic.LiquidDiffusivityUserInput)
    Call unitsys_set_number_in_base_units(frmInputKineticParameters.txtIonicTransportCoeffCorr, _
      NowProj.Cation(i).Kinetic.IonicTransportCoeffCorrelation)
    Call unitsys_set_number_in_base_units(frmInputKineticParameters.txtIonicTransCoeffUser, _
      NowProj.Cation(i).Kinetic.IonicTransportCoeffUserInput)
    Call unitsys_set_number_in_base_units(frmInputKineticParameters.txtPoreDiffusivityCorr, _
      NowProj.Cation(i).Kinetic.PoreDiffusivityCorrelation)
    Call unitsys_set_number_in_base_units(frmInputKineticParameters.txtPoreDiffusivityUser, _
      NowProj.Cation(i).Kinetic.PoreDiffusivityUserInput)
  Next i
  For i = 1 To NowProj.NumberOfAnions
    Call unitsys_set_number_in_base_units(frmInputKineticParameters.txtLiquidDiffCorrelation, _
      NowProj.Anion(i).Kinetic.LiquidDiffusivityCorrelation)
    Call unitsys_set_number_in_base_units(frmInputKineticParameters.txtLiquidDiffUserInput, _
      NowProj.Anion(i).Kinetic.LiquidDiffusivityUserInput)
    Call unitsys_set_number_in_base_units(frmInputKineticParameters.txtIonicTransportCoeffCorr, _
      NowProj.Anion(i).Kinetic.IonicTransportCoeffCorrelation)
    Call unitsys_set_number_in_base_units(frmInputKineticParameters.txtIonicTransCoeffUser, _
      NowProj.Anion(i).Kinetic.IonicTransportCoeffUserInput)
    Call unitsys_set_number_in_base_units(frmInputKineticParameters.txtPoreDiffusivityCorr, _
      NowProj.Anion(i).Kinetic.PoreDiffusivityCorrelation)
    Call unitsys_set_number_in_base_units(frmInputKineticParameters.txtPoreDiffusivityUser, _
      NowProj.Anion(i).Kinetic.PoreDiffusivityUserInput)
  Next i
  
  

End Sub
