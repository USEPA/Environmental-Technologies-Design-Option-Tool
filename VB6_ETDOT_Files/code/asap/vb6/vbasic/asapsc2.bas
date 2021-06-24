Attribute VB_Name = "AsapSc2_Mod"
Option Explicit

Global UsersFlowsLoadingsOption As Integer
Global InitialPressureDrop As Double
Global FinalPressureDrop As Double
Global PressureDropStep As Double

Global CurrentMode As Integer  '1=design, 2=rating

Sub CalculatePowerScreen2(CalculatedPower As Integer)
    Dim CalculatedBlowerPower As Integer
    Dim CalculatedPumpPower As Integer

    CalculatedBlowerPower = False
    If HaveValue(Scr2.AirFlowRate.value) And HaveValue(Scr2.TowerArea.value) And HaveValue(Scr2.OperatingPressure.value) And HaveValue(Scr2.AirPressureDrop.value) And HaveValue(Scr2.TowerHeight.value) And HaveValue(Scr2.AirDensity.value) Then
       Call PBLOWPT(Scr2.Power.BlowerBrakePower, Scr2.AirFlowRate.value, Scr2.TowerArea.value, Scr2.OperatingPressure.value, Scr2.AirPressureDrop.value, Scr2.SpecifiedTowerHeight.value, Scr2.AirDensity.value, Scr2.Power.InletAirTemperature, Scr2.Power.BlowerEfficiency)
       CalculatedBlowerPower = True
    End If

    CalculatedPumpPower = False
    If HaveValue(Scr2.WaterDensity.value) And HaveValue(Scr2.WaterFlowRate.value) And HaveValue(Scr2.TowerHeight.value) Then
       Call PPUMPPT(Scr2.Power.PumpBrakePower, Scr2.Power.PumpEfficiency, Scr2.WaterDensity.value, Scr2.WaterFlowRate.value, Scr2.SpecifiedTowerHeight.value)
       CalculatedPumpPower = True
    End If

    If CalculatedBlowerPower And CalculatedPumpPower Then
       Call PTOTALPT(Scr2.Power.TotalBrakePower, Scr2.Power.BlowerBrakePower, Scr2.Power.PumpBrakePower)
       CalculatedPower = True
    End If

End Sub

Sub GetContaminantConcentrationsScreen2()
Dim PercentRemoval As Double
Dim msg As String, Response As Integer
Dim Answer As String
Dim NewStep As Double
Dim Dummy As Double

    Call GetOndaMassTransferCoefficientScreen2
    Call GetDesignKLaOrKLaSafetyFactorScreen2

    'Update Contaminant of Interest | Influent Concentration.
    'frmPTADScreen2!lblDesignConcentrationValue(3).Caption = Format$(Scr2.DesignContaminant.Influent.Value, GetTheFormat(Scr2.DesignContaminant.Influent.Value))
    Call Unitted_NumberUpdate(frmPTADScreen2!UnitsInterest(3))
    
    'Update Contaminant of Interest | Treatment Objective.
    'frmPTADScreen2!lblDesignConcentrationValue(4).Caption = Format$(Scr2.DesignContaminant.TreatmentObjective.Value, GetTheFormat(Scr2.DesignContaminant.TreatmentObjective.Value))
    Call Unitted_NumberUpdate(frmPTADScreen2!UnitsInterest(4))
    
    Call EFFLPT2(Scr2.DesignContaminant.Effluent.value, Scr2.AirToWaterRatio.value, Scr2.DesignContaminant.HenrysConstant.value, Scr2.WaterFlowRate.value, Scr2.TowerArea.value, Scr2.SpecifiedTowerHeight.value, Scr2.DesignMassTransferCoefficient.value, Scr2.DesignContaminant.Influent.value)
    'Update Contaminant of Interest | Effluent Concentration.
    'frmPTADScreen2!lblDesignConcentrationValue(5).Caption = Format$(Scr2.DesignContaminant.Effluent.Value, GetTheFormat(Scr2.DesignContaminant.Effluent.Value))
    Call Unitted_NumberUpdate(frmPTADScreen2!UnitsInterest(5))

    Call REMOVPT(PercentRemoval, Scr2.DesignContaminant.Influent.value, Scr2.DesignContaminant.Effluent.value)
    frmPTADScreen2!lblDesignConcentrationValue(6).Caption = Format$(PercentRemoval, "0.0")

    'Determine Pressure Drop
    Scr2.AirPressureDrop.value = -1#
    InitialPressureDrop = 1#
    FinalPressureDrop = 1200#
    PressureDropStep = 1#

xxOLDxxPressureDrop:
    Call PDROP(Scr2.AirPressureDrop.value, Scr2.AirToWaterRatio.value, Scr2.AirLoadingRate.value, Scr2.Packing.PackingFactor, Scr2.WaterViscosity.value, Scr2.AirDensity.value, Scr2.WaterDensity.value, InitialPressureDrop, FinalPressureDrop, PressureDropStep)
    If Scr2.AirPressureDrop.value < 0 Then
'       Msg = "Failure to get within one percent of the "
'       Msg = Msg + "y-axis value on the Eckert curve "
'       Msg = Msg + "in the pressure drop range of "
'       Msg = Msg + Format$(InitialPressureDrop, "0.0") + " N/m2/m and " + Format$(FinalPressureDrop, "0.0")
'       Msg = Msg + " N/m2/m using a pressure drop step of " + Format$(PressureDropStep, "0.0000") + " N/m2/m."
'       Msg = Msg & Chr$(13) & Chr$(13)
'       Msg = Msg + "Would you like to specify a smaller value for pressure drop step "
'       Msg = Msg + "and attempt to achieve convergence again?"
'       Response = MsgBox(Msg, MB_ICONQUESTION + MB_YESNO, "Pressure Drop Convergence Error")
'       If Response = IDYES Then
'          If PressureDropStep <= .01 Then
'             Msg = "You can not specify a pressure drop "
'             Msg = Msg + "step smaller than 0.01. "
'             Msg = Msg + "Convergence not possible in this "
'             Msg = Msg + "case."
'             MsgBox Msg, MB_ICONEXCLAMATION, "Pressure Drop Convergence Error"
'             frmPTADScreen2!lblDesignConcentrationValue(7).Caption = "N/A"
'          Else
'NewPressureDrop:
'             If PressureDropStep / 10 < .01 Then
'                Answer$ = InputBox$("Enter new value for pressure drop step.", "Pressure Drop Step", Format$(.01, "0.000"))
'             Else
'                Answer$ = InputBox$("Enter new value for pressure drop step.", "Pressure Drop Step", Format$(PressureDropStep / 10, "0.000"))
'             End If
'             On Error GoTo NewPressureDrop:
'                NewStep = CDbl(Answer$)
'                If NewStep < .01 Then
'                   MsgBox "Pressure Drop step must exceed 0.01", MB_ICONEXCLAMATION, "Error"
'                   GoTo NewPressureDrop
'                Else
'                   PressureDropStep = NewStep
'                   GoTo PressureDrop:
'                End If
 '         End If

       NewStep = PressureDropStep / 10#
       If NewStep > 0.001 Then
          PressureDropStep = NewStep
          GoTo xxOLDxxPressureDrop:
       Else
          frmPTADScreen2!lblDesignConcentrationValue(7).Caption = "N/A"
       End If
    Else
       'frmPTADScreen2!lblDesignConcentrationValue(7).Caption = Format$(Scr2.AirPressureDrop.Value, "0.0")
       Call Unitted_NumberUpdate(frmPTADScreen2!UnitsInterest(7))
    End If

End Sub

Sub GetDesignKLaOrKLaSafetyFactorScreen2()

  If Scr2.KLaSafetyFactor.UserInput = True Then
    Call SpecifiedKLaSafetyFactorScreen2
  ElseIf Scr2.DesignMassTransferCoefficient.UserInput = True Then
    Call SpecifiedDesignKLaScreen2
  End If

End Sub

Sub GetFlowsAndLoadingsScreen2()


    If UsersFlowsLoadingsOption = 0 Then

       'Water Flow Rate, Air Flow Rate are specified -->
       '   Calculate:  Air to Water Ratio, Water Loading Rate,
       '               Air Loading Rate
       If HaveValue(Scr2.WaterFlowRate.value) And HaveValue(Scr2.AirFlowRate.value) And HaveValue(Scr2.WaterDensity.value) And HaveValue(Scr2.TowerArea.value) And HaveValue(Scr2.AirDensity.value) Then
          Call VQCALC(Scr2.AirToWaterRatio.value, Scr2.AirFlowRate.value, Scr2.WaterFlowRate.value)
          Call LDH2OPT2(Scr2.WaterLoadingRate.value, Scr2.WaterFlowRate.value, Scr2.WaterDensity.value, Scr2.TowerArea.value)
          Call LDAIRPT2(Scr2.AirLoadingRate.value, Scr2.AirFlowRate.value, Scr2.AirDensity.value, Scr2.TowerArea.value)
          frmPTADScreen2!txtFlowsLoadings(2).Text = Format$(Scr2.AirToWaterRatio.value, GetTheFormat(Scr2.AirToWaterRatio.value))
          'frmPTADScreen2!txtFlowsLoadings(3).Text = Format$(Scr2.WaterLoadingRate.Value, GetTheFormat(Scr2.WaterLoadingRate.Value))
          'frmPTADScreen2!txtFlowsLoadings(4).Text = Format$(Scr2.AirLoadingRate.Value, GetTheFormat(Scr2.AirLoadingRate.Value))
          'Call Unitted_NumberUpdate(frmPTADScreen2!UnitsFlows(2))
          Call Unitted_NumberUpdate(frmPTADScreen2!UnitsFlows(3))
          Call Unitted_NumberUpdate(frmPTADScreen2!UnitsFlows(4))
       End If

    ElseIf UsersFlowsLoadingsOption = 1 Then
       
       'Water Flow Rate, Air to Water Ratio are specified -->
       '   Calculate:  Air Flow Rate, Water Loading Rate,
       '               Air Loading Rate
       If HaveValue(Scr2.AirToWaterRatio.value) And HaveValue(Scr2.WaterFlowRate.value) And HaveValue(Scr2.WaterDensity.value) And HaveValue(Scr2.TowerArea.value) And HaveValue(Scr2.AirDensity.value) Then
          Call AIRFLO(Scr2.AirFlowRate.value, Scr2.AirToWaterRatio.value, Scr2.WaterFlowRate.value)
          Call LDH2OPT2(Scr2.WaterLoadingRate.value, Scr2.WaterFlowRate.value, Scr2.WaterDensity.value, Scr2.TowerArea.value)
          Call LDAIRPT2(Scr2.AirLoadingRate.value, Scr2.AirFlowRate.value, Scr2.AirDensity.value, Scr2.TowerArea.value)
          'frmPTADScreen2!txtFlowsLoadings(1).Text = Format$(Scr2.AirFlowRate.Value, GetTheFormat(Scr2.AirFlowRate.Value))
          'frmPTADScreen2!txtFlowsLoadings(3).Text = Format$(Scr2.WaterLoadingRate.Value, GetTheFormat(Scr2.WaterLoadingRate.Value))
          'frmPTADScreen2!txtFlowsLoadings(4).Text = Format$(Scr2.AirLoadingRate.Value, GetTheFormat(Scr2.AirLoadingRate.Value))
          Call Unitted_NumberUpdate(frmPTADScreen2!UnitsFlows(1))
          Call Unitted_NumberUpdate(frmPTADScreen2!UnitsFlows(3))
          Call Unitted_NumberUpdate(frmPTADScreen2!UnitsFlows(4))
       End If

    ElseIf UsersFlowsLoadingsOption = 2 Then

   'Water Loading Rate, Air Loading Rate are specified -->
       '   Calculate:  Water Flow Rate, Air Flow Rate,
       '               Air to Water Ratio

       If HaveValue(Scr2.WaterLoadingRate.value) And HaveValue(Scr2.WaterDensity.value) And HaveValue(Scr2.TowerArea.value) And HaveValue(Scr2.AirLoadingRate.value) And HaveValue(Scr2.AirDensity.value) Then
          Call QH2OPT2(Scr2.WaterFlowRate.value, Scr2.WaterLoadingRate.value, Scr2.WaterDensity.value, Scr2.TowerArea.value)
          Call QAIRPT2(Scr2.AirFlowRate.value, Scr2.AirLoadingRate.value, Scr2.AirDensity.value, Scr2.TowerArea.value)
          Call VQCALC(Scr2.AirToWaterRatio.value, Scr2.AirFlowRate.value, Scr2.WaterFlowRate.value)
          'frmPTADScreen2!txtFlowsLoadings(0).Text = Format$(Scr2.WaterFlowRate.Value, GetTheFormat(Scr2.WaterFlowRate.Value))
          'frmPTADScreen2!txtFlowsLoadings(1).Text = Format$(Scr2.AirFlowRate.Value, GetTheFormat(Scr2.AirFlowRate.Value))
          frmPTADScreen2!txtFlowsLoadings(2).Text = Format$(Scr2.AirToWaterRatio.value, GetTheFormat(Scr2.AirToWaterRatio.value))
          Call Unitted_NumberUpdate(frmPTADScreen2!UnitsFlows(0))
          Call Unitted_NumberUpdate(frmPTADScreen2!UnitsFlows(1))
          'Call Unitted_NumberUpdate(frmPTADScreen2!UnitsFlows(2))
       End If

    End If
End Sub

Sub GetOndaMassTransferCoefficientScreen2()
Dim Dummy As Double

    If HaveValue(Scr2.Packing.CriticalSurfaceTension) And HaveValue(Scr2.WaterSurfaceTension.value) And HaveValue(Scr2.WaterLoadingRate.value) And HaveValue(Scr2.Packing.SpecificSurfaceArea) And HaveValue(Scr2.WaterViscosity.value) And HaveValue(Scr2.WaterDensity.value) And HaveValue(Scr2.DesignContaminant.LiquidDiffusivity.value) And HaveValue(Scr2.Packing.NominalSize) And HaveValue(Scr2.AirLoadingRate.value) And HaveValue(Scr2.AirViscosity.value) And HaveValue(Scr2.AirDensity.value) And HaveValue(Scr2.DesignContaminant.GasDiffusivity.value) And HaveValue(Scr2.DesignContaminant.HenrysConstant.value) Then
       Call AWCALC(Scr2.Packing.OndaWettedSurfaceArea, Scr2.Packing.CriticalSurfaceTension, Scr2.WaterSurfaceTension.value, Scr2.WaterLoadingRate.value, Scr2.Packing.SpecificSurfaceArea, Scr2.WaterViscosity.value, Scr2.WaterDensity.value, Scr2.Onda.ReynoldsNumber, Scr2.Onda.FroudeNumber, Scr2.Onda.WeberNumber)
       Call ONDAKLPT(Scr2.Onda.LiquidPhaseMassTransferCoefficient, Scr2.WaterLoadingRate.value, Scr2.Packing.OndaWettedSurfaceArea, Scr2.WaterViscosity.value, Scr2.WaterDensity.value, Scr2.DesignContaminant.LiquidDiffusivity.value, Scr2.Packing.SpecificSurfaceArea, Scr2.Packing.NominalSize)
       Call ONDAKGPT(Scr2.Onda.GasPhaseMassTransferCoefficient, Scr2.AirLoadingRate.value, Scr2.Packing.SpecificSurfaceArea, Scr2.AirViscosity.value, Scr2.AirDensity.value, Scr2.DesignContaminant.GasDiffusivity.value, Scr2.Packing.NominalSize)
       Call ONDKLAPT(Scr2.Onda.OverallMassTransferCoefficient, Scr2.Onda.LiquidPhaseMassTransferResistance, Scr2.Onda.GasPhaseMassTransferResistance, Scr2.Onda.TotalMassTransferResistance, Scr2.Onda.LiquidPhaseMassTransferCoefficient, Scr2.Packing.OndaWettedSurfaceArea, Scr2.Onda.GasPhaseMassTransferCoefficient, Scr2.DesignContaminant.HenrysConstant.value)

       'Update Contaminant of Interest | Onda KLa.
       'frmPTADScreen2.lblDesignConcentrationValue(0).Caption = Format$(Scr2.Onda.OverallMassTransferCoefficient, GetTheFormat(Scr2.Onda.OverallMassTransferCoefficient))
       Call Unitted_NumberUpdate(frmPTADScreen2!UnitsInterest(0))

       Scr2.Onda.ValChanged = True

       Call ShowOndaKLaPropertiesScreen2

    End If

End Sub

Sub GetTowerAreaAndVolume()
    
  If HaveValue(Scr2.SpecifiedTowerDiameter.value) And HaveValue(Scr2.SpecifiedTowerHeight.value) Then
    Call AREAPT2(Scr2.TowerArea.value, Scr2.SpecifiedTowerDiameter.value)
    Call TVOLPT2(Scr2.TowerVolume.value, Scr2.TowerArea.value, Scr2.SpecifiedTowerHeight.value)
    'frmPTADScreen2!lblTowerParameters(2).Caption = Format$(Scr2.TowerArea.Value, GetTheFormat(Scr2.TowerArea.Value))
    'frmPTADScreen2!lblTowerParameters(3).Caption = Format$(Scr2.TowerVolume.Value, GetTheFormat(Scr2.TowerVolume.Value))
    Call Unitted_NumberUpdate(frmPTADScreen2!UnitsTowerParam(2))
    Call Unitted_NumberUpdate(frmPTADScreen2!UnitsTowerParam(3))
  End If

End Sub

Sub LoadContaminantListScreen2()
    Dim FileID As String, msg As String
    Dim Pressure As Double, Temperature As Double
    Dim i As Integer
    Dim NotSpecifiedAtOperatingTemperature As Integer
    Dim NotSpecifiedAtOperatingPressure As Integer

    Call LoadFile(Filename)
    
    If Filename$ <> "" Then
       FileID = ""
       Open Filename$ For Input As #1
       On Error Resume Next
       Input #1, FileID
       If FileID <> CONTAMINANTS_PTAD_FILEID Then
          msg = "Invalid Contaminant File"
          MsgBox msg, 48, "Error"
          Close #1
          Exit Sub
       End If

       'frmListContaminantScreen2.ListContaminants.Clear
       frmPTADScreen2!cboSelectCompo.Clear

       i = 0
       NotSpecifiedAtOperatingTemperature = False
       NotSpecifiedAtOperatingPressure = False
       Do Until EOF(1)
          i = i + 1
          Input #1, Scr2.Contaminant(i).Pressure, Scr2.Contaminant(i).Temperature, Scr2.Contaminant(i).Name, Scr2.Contaminant(i).MolecularWeight.value, Scr2.Contaminant(i).HenrysConstant.value, Scr2.Contaminant(i).MolarVolume.value, Scr2.Contaminant(i).NormalBoilingPoint.value, Scr2.Contaminant(i).LiquidDiffusivity.value, Scr2.Contaminant(i).GasDiffusivity.value, Scr2.Contaminant(i).Influent.value, Scr2.Contaminant(i).TreatmentObjective.value
          'frmListContaminantScreen2.ListContaminants.AddItem Scr2.Contaminant(i).Name
          frmPTADScreen2!cboSelectCompo.AddItem Scr2.Contaminant(i).Name

          If Not NotSpecifiedAtOperatingTemperature Then
             If Abs(Scr2.Contaminant(i).Temperature - Scr2.operatingtemperature.value) > TOLERANCE Then
                NotSpecifiedAtOperatingTemperature = True
             End If
          End If
          If Not NotSpecifiedAtOperatingPressure Then
             If Abs(Scr2.Contaminant(i).Pressure - Scr2.OperatingPressure.value) > TOLERANCE Then
                NotSpecifiedAtOperatingPressure = True
             End If
          End If

       Loop
       Scr2.NumChemical = i
          
       Close #1

       'If frmListContaminantScreen2.mnuOptionsManipulateContaminant(1).Enabled = False Then
       '   frmListContaminantScreen2.mnuOptionsManipulateContaminant(1).Enabled = True
       '   frmListContaminantScreen2.mnuOptionsManipulateContaminant(3).Enabled = True
       '   frmListContaminantScreen2.mnuOptionsManipulateContaminant(4).Enabled = True
       '   frmListContaminantScreen2.mnuOptionsSave.Enabled = True
       '   frmListContaminantScreen2.mnuOptionsView.Enabled = True
       'End If

       'frmListContaminantScreen2.ListContaminants.Selected(0) = True

       If NotSpecifiedAtOperatingPressure And NotSpecifiedAtOperatingTemperature Then
          MsgBox "For one or more contaminants, the temperature and pressure at which the contaminant properties are specified differs from the operating temperature and pressure.", MB_ICONINFORMATION, "Warning"
       ElseIf NotSpecifiedAtOperatingTemperature Then
          MsgBox "For one or more contaminants, the temperature at which the contaminant properties are specified differs from the operating temperature.", MB_ICONINFORMATION, "Warning"
       ElseIf NotSpecifiedAtOperatingPressure Then
          MsgBox "For one or more contaminants, the pressure at which the contaminant properties are specified differs from the operating pressure.", MB_ICONINFORMATION, "Warning"
       End If

    End If

End Sub

Sub LoadFileScreen2(Filename As String)
Dim Ctl As Control
Set Ctl = frmPTADScreen2.CommonDialog1

    On Error Resume Next
    
    'frmPTADScreen2!CMDialog1.DefaultExt = "rat"
    'frmPTADScreen2!CMDialog1.Filter = "Rating Mode Files (*.rat)|*.rat"
    'frmPTADScreen2!CMDialog1.DialogTitle = "Load Packed Tower Aeration Rating Mode File"
    'frmPTADScreen2!CMDialog1.flags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST
    'frmPTADScreen2!CMDialog1.Action = 1
    'Filename$ = frmPTADScreen2!CMDialog1.Filename
    Ctl.DefaultExt = "rat"
    Ctl.Filter = "Rating Mode Files (*.rat)|*.rat"
    Ctl.DialogTitle = "Load Packed Tower Aeration Rating Mode File"
    Ctl.flags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST
    Ctl.Action = 1
    Filename$ = Ctl.Filename
    If Err = 32755 Then   'Cancel selected by user
       Filename$ = ""
    End If

End Sub

Function loadscreen2(OverrideFilename As String) As Boolean
Dim FileID As String, msg As String
Dim i As Integer
Dim FoundCurrentPacking As Integer  'Whether packing user specified is currently in the user-modified database or if we have to add it when the database is the user-modified one.
Dim CurrPackingIndex As Integer
Dim FlowsLoadingsString As String
Dim UsersFlowAndLoadingOption As Integer
ReDim u(10) As String
Dim xu As rec_Units_frmContaminantPropertyEdit

    If (OverrideFilename <> "") Then
      Filename = OverrideFilename
    Else
      If Filename = "TheDefaultCaseScreen2" Then
        Filename = App.Path & "\dbase\default.rat"
      Else
        Call LoadFileScreen2(Filename)
      End If
    End If
    
    If Filename$ <> "" Then
       FileID = ""
       If (fileexists(Filename) = False) Then
         Call Error_Unavailable_File( _
            Filename, _
            "Packed Tower Aeration Rating Mode")
         loadscreen2 = False
         Exit Function
       End If
       Open Filename$ For Input As #1
       On Error Resume Next
       Input #1, FileID
       If FileID <> SCREEN2_PTAD2_FILEID Then
          msg = "Invalid Optimization File"
          MsgBox msg, 48, "Error"
          Close #1
          Exit Function
       End If

       'frmListContaminantScreen2.ListContaminants.Clear
       frmPTADScreen2!cboSelectCompo.Clear
      
       '**********************
       '*
       '* Tower Parameters
       '*
       '**********************

       Input #1, Scr2.TowerDiameter.value
       frmPTADScreen2!lblDesignParameters(0).Caption = Format$(Scr2.TowerDiameter.value, GetTheFormat(Scr2.TowerDiameter.value))

       Input #1, Scr2.TowerHeight.value
       frmPTADScreen2!lblDesignParameters(1).Caption = Format$(Scr2.TowerHeight.value, GetTheFormat(Scr2.TowerHeight.value))

       Input #1, Scr2.SpecifiedTowerDiameter.value
       frmPTADScreen2!txtTowerParameters(0).Text = Format$(Scr2.SpecifiedTowerDiameter.value, GetTheFormat(Scr2.SpecifiedTowerDiameter.value))

       Input #1, Scr2.SpecifiedTowerHeight.value
       frmPTADScreen2!txtTowerParameters(1).Text = Format$(Scr2.SpecifiedTowerHeight.value, GetTheFormat(Scr2.SpecifiedTowerHeight.value))

       Call GetTowerAreaAndVolume


       '********************************************************************
       '*
       '*  Pressure, Temperature, and Physical Properties of Air and Water
       '*
       '********************************************************************

       Input #1, Scr2.OperatingPressure.value
       frmPTADScreen2!txtOperatingPressure.Text = Format$(Scr2.OperatingPressure.value * 101325# / 1#, "0.00")
       Scr2.OperatingPressure.ValChanged = True

       Input #1, Scr2.operatingtemperature.value
       frmPTADScreen2!txtOperatingTemperature.Text = Format$(Scr2.operatingtemperature.value - 273.15, "0.0")
       Scr2.operatingtemperature.ValChanged = True

       If HaveValue(Scr2.OperatingPressure.value) And HaveValue(Scr2.operatingtemperature.value) Then
          Call CalculateAirWaterPropertiesScreen2
       End If


       '************
       '*
       '* Packing
       '*
       '************

       Input #1, Scr2.Packing.Name, Scr2.Packing.NominalSize, Scr2.Packing.PackingFactor, Scr2.Packing.SpecificSurfaceArea, Scr2.Packing.CriticalSurfaceTension, Scr2.Packing.Material, Scr2.Packing.source, Scr2.Packing.UserInput, Scr2.Packing.SourceDatabase
       frmPTADScreen2!lblPackingType.Caption = Scr2.Packing.Name

       If PackingDatabaseSource <> Scr2.Packing.SourceDatabase Then
          frmSelectPacking!cboSelectPacking.Clear
          If Scr2.Packing.SourceDatabase = ORIGINALPACKINGDATABASE Then
             frmSelectPacking!mnuPackDatabase(0).Checked = True
             frmSelectPacking!mnuPackDatabase(1).Checked = False
             frmSelectPacking!mnuPackDatabaseOptions(0).Enabled = False
          
             For i = 1 To NumPackingsInDatabase
                 frmSelectPacking!cboSelectPacking.AddItem DatabasePacking(i).Name
             Next i
             frmSelectPacking!mnuPackDatabase(3).Enabled = False
          ElseIf Scr2.Packing.SourceDatabase = USERMODIFIEDPACKINGDATABASE Then
             frmSelectPacking!mnuPackDatabase(0).Checked = False
             frmSelectPacking!mnuPackDatabase(1).Checked = True
       
             For i = 1 To NumUserPackings
                 frmSelectPacking!cboSelectPacking.AddItem UserPacking(i).Name
             Next i
             frmSelectPacking!mnuPackDatabase(3).Enabled = True
          End If
       End If

       If Scr2.Packing.SourceDatabase = USERMODIFIEDPACKINGDATABASE Then
             FoundCurrentPacking = False
             For i = 1 To NumUserPackings
                 If UserPacking(i).Name = Scr2.Packing.Name Then
                    FoundCurrentPacking = True
                    CurrPackingIndex = i
                 End If
             Next i

             If FoundCurrentPacking Then
                If Scr2.Packing.NominalSize <> UserPacking(CurrPackingIndex).NominalSize Or Scr2.Packing.PackingFactor <> UserPacking(CurrPackingIndex).PackingFactor Or Scr2.Packing.SpecificSurfaceArea <> UserPacking(CurrPackingIndex).SpecificSurfaceArea Or Scr2.Packing.CriticalSurfaceTension <> UserPacking(CurrPackingIndex).CriticalSurfaceTension Or Scr2.Packing.Material <> UserPacking(CurrPackingIndex).Material Or Scr2.Packing.source <> UserPacking(CurrPackingIndex).source Then
                   msg = "Name of packing to be loaded matches the name "
                   msg = msg + "of a packing in the user-modified packing "
                   msg = msg + "database, but the properties of the two "
                   msg = msg + "packings differ." & Chr$(13) & Chr$(13)
                   msg = msg + "The properties of the packing to be loaded "
                   msg = msg + "will overwrite the properties currently "
                   msg = msg + "in the user-modified packing database."
                   MsgBox msg, MB_ICONEXCLAMATION, "Name of Packing Conflict"
                   UserPacking(CurrPackingIndex) = Scr2.Packing

                End If
             End If

             If Not FoundCurrentPacking Then
                NumUserPackings = NumUserPackings + 1
                UserPacking(NumUserPackings) = Scr2.Packing
                frmSelectPacking!cboSelectPacking.AddItem Scr2.Packing.Name
                frmSelectPacking!cboSelectPacking.ListIndex = NumUserPackings - 1
             End If
       End If


       '*******************************
       '*
       '* Flow and Loading Parameters
       '*
       '*******************************

       Input #1, FlowsLoadingsString$
       If FlowsLoadingsString = "Specified Water Flow Rate and Air Flow Rate" Then
          Input #1, Scr2.WaterFlowRate.value
          Input #1, Scr2.AirFlowRate.value
          frmPTADScreen2!txtFlowsLoadings(0).Text = Trim$(Str$(Scr2.WaterFlowRate.value))
          frmPTADScreen2!txtFlowsLoadings(1).Text = Format$(Scr2.AirFlowRate.value, GetTheFormat(Scr2.AirFlowRate.value))
          frmFlowsLoadingsScreen2!optFlowsLoadings(0).value = True
          UsersFlowAndLoadingOption = 0
       ElseIf FlowsLoadingsString = "Specified Water Flow Rate and Air to Water Ratio" Then
          Input #1, Scr2.WaterFlowRate.value
          Input #1, Scr2.AirToWaterRatio.value
          frmPTADScreen2!txtFlowsLoadings(0).Text = Trim$(Str$(Scr2.WaterFlowRate.value))
          frmPTADScreen2!txtFlowsLoadings(2).Text = Format$(Scr2.AirToWaterRatio.value, GetTheFormat(Scr2.AirToWaterRatio.value))
          frmFlowsLoadingsScreen2!optFlowsLoadings(1).value = True
          UsersFlowAndLoadingOption = 1
       ElseIf FlowsLoadingsString = "Specified Water Loading Rate and Air Loading Rate" Then
          Input #1, Scr2.WaterLoadingRate.value
          Input #1, Scr2.AirLoadingRate.value
          frmPTADScreen2!txtFlowsLoadings(3).Text = Format$(Scr2.WaterLoadingRate.value, GetTheFormat(Scr2.WaterLoadingRate.value))
          frmPTADScreen2!txtFlowsLoadings(4).Text = Format$(Scr2.AirLoadingRate.value, GetTheFormat(Scr2.AirLoadingRate.value))
          frmFlowsLoadingsScreen2!optFlowsLoadings(2).value = True
          UsersFlowAndLoadingOption = 2
       End If
       Call SetUpFlowsLoadingsTextBoxes(UsersFlowAndLoadingOption)
       UsersFlowsLoadingsOption = UsersFlowAndLoadingOption
       Call GetFlowsAndLoadingsScreen2


       '***********************************
       '*
       '* Contaminant Properties
       '*
       '***********************************

       Input #1, Scr2.NumChemical
       For i = 1 To Scr2.NumChemical
           Input #1, Scr2.Contaminant(i).Pressure, Scr2.Contaminant(i).Temperature, Scr2.Contaminant(i).Name, Scr2.Contaminant(i).MolecularWeight.value, Scr2.Contaminant(i).HenrysConstant.value, Scr2.Contaminant(i).MolarVolume.value, Scr2.Contaminant(i).NormalBoilingPoint.value, Scr2.Contaminant(i).LiquidDiffusivity.value, Scr2.Contaminant(i).GasDiffusivity.value, Scr2.Contaminant(i).Influent.value, Scr2.Contaminant(i).TreatmentObjective.value
           'frmListContaminantScreen2.ListContaminants.AddItem Scr2.Contaminant(i).Name
           frmPTADScreen2!cboSelectCompo.AddItem Scr2.Contaminant(i).Name
       Next i
       Input #1, Scr2.DesignContaminant.Name

       For i = 1 To Scr2.NumChemical
           If Scr2.DesignContaminant.Name = Scr2.Contaminant(i).Name Then
              Scr2.DesignContaminant = Scr2.Contaminant(i)
              'frmListContaminantScreen2!ListContaminants.Selected(i - 1) = True
              frmPTADScreen2!cboSelectCompo.ListIndex = i - 1
              Exit For
           End If
       Next i

       'If frmListContaminantScreen2!mnuOptionsManipulateContaminant(1).Enabled = False Then
       '   frmListContaminantScreen2!mnuOptionsManipulateContaminant(1).Enabled = True
       '   frmListContaminantScreen2!mnuOptionsManipulateContaminant(3).Enabled = True
       '   frmListContaminantScreen2!mnuOptionsManipulateContaminant(4).Enabled = True
       '   frmListContaminantScreen2!mnuOptionsSave.Enabled = True
       '   frmListContaminantScreen2!mnuOptionsView.Enabled = True
       '
       '   frmPTADScreen2!mnuFile(4).Enabled = True
       '   frmPTADScreen2!mnuFile(5).Enabled = True
       '   frmPTADScreen2!mnuOptions(0).Enabled = True
       'End If
       
       'Call SetDesignContaminantEnabledScreen2(CInt(frmListContaminantScreen2!ListContaminants.ListCount))


       '*************************************
       '*
       '* Mass Transfer Properties
       '*
       '*************************************

       Input #1, Scr2.KLaSafetyFactor.value, Scr2.KLaSafetyFactor.UserInput
       Input #1, Scr2.DesignMassTransferCoefficient.value, Scr2.DesignMassTransferCoefficient.UserInput
       If Scr2.KLaSafetyFactor.UserInput = True Then
          frmPTADScreen2!txtDesignConcentrationValue(1).Text = Format$(Scr2.KLaSafetyFactor.value, GetTheFormat(Scr2.KLaSafetyFactor.value))
       ElseIf Scr2.DesignMassTransferCoefficient.UserInput = True Then
          frmPTADScreen2!txtDesignConcentrationValue(2).Text = Format$(Scr2.DesignMassTransferCoefficient.value, GetTheFormat(Scr2.DesignMassTransferCoefficient.value))
       End If

       'Input the units of this screen.
       Input #1, u(1), u(2)
       Call SetUnits(frmPTADScreen2!UnitsDesignBasis(0), u(1))
       Call SetUnits(frmPTADScreen2!UnitsDesignBasis(1), u(2))
  
       Input #1, u(1), u(2), u(3), u(4)
       Call SetUnits(frmPTADScreen2!UnitsTowerParam(0), u(1))
       Call SetUnits(frmPTADScreen2!UnitsTowerParam(1), u(2))
       Call SetUnits(frmPTADScreen2!UnitsTowerParam(2), u(3))
       Call SetUnits(frmPTADScreen2!UnitsTowerParam(3), u(4))
       
       Input #1, u(1), u(2)
       Call SetUnits(frmPTADScreen2!UnitsOpCond(0), u(1))
       Call SetUnits(frmPTADScreen2!UnitsOpCond(1), u(2))
  
       Input #1, u(1), u(2), u(3), u(4)
       Call SetUnits(frmPTADScreen2!UnitsFlows(0), u(1))
       Call SetUnits(frmPTADScreen2!UnitsFlows(1), u(2))
       Call SetUnits(frmPTADScreen2!UnitsFlows(2), u(3))
       Call SetUnits(frmPTADScreen2!UnitsFlows(3), u(4))
  
       Input #1, u(1), u(2), u(3), u(4), u(5), u(6)
       Call SetUnits(frmPTADScreen2!UnitsInterest(0), u(1))
       Call SetUnits(frmPTADScreen2!UnitsInterest(2), u(2))
       Call SetUnits(frmPTADScreen2!UnitsInterest(3), u(3))
       Call SetUnits(frmPTADScreen2!UnitsInterest(4), u(4))
       Call SetUnits(frmPTADScreen2!UnitsInterest(5), u(5))
       Call SetUnits(frmPTADScreen2!UnitsInterest(7), u(6))

       'Input the units of frmContaminantPropertyEdit.
       xu = Units_frmContaminantPropertyEdit
       Input #1, xu.UnitsProp(0), xu.UnitsProp(2), xu.UnitsProp(3), xu.UnitsProp(4), xu.UnitsProp(5)
       Input #1, xu.UnitsConc(0), xu.UnitsConc(1)
       Units_frmContaminantPropertyEdit = xu
       
       Close #1

       Call GetContaminantConcentrationsScreen2

       ShownScreen1Previously = False

       frmPTADScreen2.Caption = "Packed Tower Aeration - Rating Mode"
       If Right$(Filename, 11) = "default.des" Or Right$(Filename, 11) = "default.rat" Then
          frmPTADScreen2.Caption = frmPTADScreen2.Caption & " (" & "untitled.rat" & ")"
       Else
          frmPTADScreen2.Caption = frmPTADScreen2.Caption & " (" & Filename & ")"
       End If

       'Add this file to the last-few-files list.
       Call LastFewFiles_MoveFilenameToTop(Filename)

    End If
    
    loadscreen2 = True
    
End Function

Sub NewPagePTADScreen2()

          Printer.NewPage
          Printer.FontSize = 12
          Printer.FontBold = True
          Printer.Print "Packed Tower Aeration - Rating Mode (continued)"
          Printer.Print
          Printer.Print
          Printer.FontSize = 10
          Printer.FontBold = False

End Sub

Sub PrintPTADScreen2()
    Dim i As Integer, j As Integer
    Dim CalculatedPower As Integer
    ReDim OndaKLa(1 To MAXCHEMICAL) As Double
    Dim KLaSafetyFactor As Double
    ReDim DesignKLa(1 To MAXCHEMICAL) As Double
    ReDim PackingWettedSurfaceArea(1 To MAXCHEMICAL) As Double
    Dim ReynoldsNumber As Double
    Dim FroudeNumber As Double
    Dim WeberNumber As Double
    Dim LiquidPhaseMassTransferCoefficient As Double
    Dim GasPhaseMassTransferCoefficient As Double
    Dim LiquidPhaseMassTransferResistance As Double
    Dim GasPhaseMassTransferResistance As Double
    Dim TotalMassTransferResistance As Double
    ReDim DesiredPercentRemoval(1 To MAXCHEMICAL) As Double
    ReDim Effluent(1 To MAXCHEMICAL) As Double
    ReDim AchievedPercentRemoval(1 To MAXCHEMICAL) As Double

    On Error GoTo PrinterError

          Printer.ScaleLeft = -1440
          Printer.ScaleTop = -1440
          Printer.CurrentX = 0
          Printer.CurrentY = 0
          Printer.FontSize = 12
          Printer.FontBold = True
          Printer.Print "Packed Tower Aeration - Rating Mode"
          Printer.Print
          Printer.Print
          Printer.FontUnderline = True
          Printer.FontSize = 10
          Printer.Print "Property:"; Tab(VALUE_TAB); "Value:"
          Printer.Print
          Printer.FontUnderline = False
          Printer.FontBold = False
          If ShownScreen1Previously Then
             Printer.Print "Design based on:  "; frmPTADScreen2!lblDesignParametersLabel(0).Caption & " (" & frmPTADScreen2!UnitsDesignBasis(0) & ")"; Tab(VALUE_TAB); frmPTADScreen2!lblDesignParameters(0).Caption
             Printer.Print "Design based on:  "; frmPTADScreen2!lblDesignParametersLabel(1).Caption & " (" & frmPTADScreen2!UnitsDesignBasis(1) & ")"; Tab(VALUE_TAB); frmPTADScreen2!lblDesignParameters(1).Caption
          End If
          Printer.Print frmPTADScreen2!lblTowerParametersLabel(0).Caption & " (" & frmPTADScreen2!UnitsTowerParam(0) & ")"; Tab(VALUE_TAB); frmPTADScreen2!txtTowerParameters(0).Text
          Printer.Print frmPTADScreen2!lblTowerParametersLabel(1).Caption & " (" & frmPTADScreen2!UnitsTowerParam(1) & ")"; Tab(VALUE_TAB); frmPTADScreen2!txtTowerParameters(1).Text
          Printer.Print frmPTADScreen2!lblTowerParametersLabel(2).Caption & " (" & frmPTADScreen2!UnitsTowerParam(2) & ")"; Tab(VALUE_TAB); frmPTADScreen2!lblTowerParameters(2).Caption
          Printer.Print frmPTADScreen2!lblTowerParametersLabel(3).Caption & " (" & frmPTADScreen2!UnitsTowerParam(3) & ")"; Tab(VALUE_TAB); frmPTADScreen2!lblTowerParameters(3).Caption
          Printer.Print
          Printer.Print "Operating Pressure" & " (" & frmPTADScreen2!UnitsOpCond(0) & ")"; Tab(VALUE_TAB); frmPTADScreen2!txtOperatingPressure.Text
          Printer.Print "Operating Temperature" & " (" & frmPTADScreen2!UnitsOpCond(1) & ")"; Tab(VALUE_TAB); frmPTADScreen2!txtOperatingTemperature.Text
          Printer.Print frmAirWaterProperties!lblAirWaterProperties(0).Caption; Tab(VALUE_TAB); frmAirWaterProperties!txtAirWaterProperties(0).Text
          Printer.Print frmAirWaterProperties!lblAirWaterProperties(1).Caption; Tab(VALUE_TAB); frmAirWaterProperties!txtAirWaterProperties(1).Text
          Printer.Print frmAirWaterProperties!lblAirWaterProperties(2).Caption; Tab(VALUE_TAB); frmAirWaterProperties!txtAirWaterProperties(2).Text
          Printer.Print frmAirWaterProperties!lblAirWaterProperties(3).Caption; Tab(VALUE_TAB); frmAirWaterProperties!txtAirWaterProperties(3).Text
          Printer.Print frmAirWaterProperties!lblAirWaterProperties(4).Caption; Tab(VALUE_TAB); frmAirWaterProperties!txtAirWaterProperties(4).Text
          Printer.Print
          CurrentScreen = Scr2
          Printer.Print "Packing Name:  "; Trim$(CurrentScreen.Packing.Name)
          Printer.Print "Packing "; frmSelectPacking!lblPackingProperties(1).Caption; Tab(VALUE_TAB); Format$(CurrentScreen.Packing.NominalSize, GetTheFormat(CurrentScreen.Packing.NominalSize))
          Printer.Print frmSelectPacking!lblPackingProperties(2).Caption; Tab(VALUE_TAB); Format$(CurrentScreen.Packing.PackingFactor, GetTheFormat(CurrentScreen.Packing.PackingFactor))
          Printer.Print "Packing "; frmSelectPacking!lblPackingProperties(3).Caption; Tab(VALUE_TAB); Format$(CurrentScreen.Packing.SpecificSurfaceArea, GetTheFormat(CurrentScreen.Packing.SpecificSurfaceArea))
          Printer.Print "Packing "; frmSelectPacking!lblPackingProperties(4).Caption; Tab(VALUE_TAB); Format$(CurrentScreen.Packing.CriticalSurfaceTension, GetTheFormat(CurrentScreen.Packing.CriticalSurfaceTension))
          Printer.Print "Packing "; frmSelectPacking!lblPackingProperties(5).Caption; Tab(VALUE_TAB); Trim$(CurrentScreen.Packing.Material)
          Printer.Print "Packing "; frmSelectPacking!lblPackingProperties(6).Caption; Tab(VALUE_TAB); Trim$(CurrentScreen.Packing.source)
          Printer.Print "Source of This Packing Data in Program"; Tab(VALUE_TAB);
          If PackingDatabaseSource = ORIGINALPACKINGDATABASE Then
             Printer.Print "Original Packing Database"
          Else
             Printer.Print "User Input"
          End If
          
          Printer.Print
          Printer.Print frmPTADScreen2!lblFlowsLoadingsLabel(0).Caption & " (" & frmPTADScreen2!UnitsFlows(0) & ")"; Tab(VALUE_TAB); frmPTADScreen2!txtFlowsLoadings(0).Text
          Printer.Print frmPTADScreen2!lblFlowsLoadingsLabel(1).Caption & " (" & frmPTADScreen2!UnitsFlows(1) & ")"; Tab(VALUE_TAB); frmPTADScreen2!txtFlowsLoadings(1).Text
          Printer.Print frmPTADScreen2!lblFlowsLoadingsLabel(2).Caption & " (-)"; Tab(VALUE_TAB); frmPTADScreen2!txtFlowsLoadings(2).Text
          Printer.Print frmPTADScreen2!lblFlowsLoadingsLabel(3).Caption & " (" & frmPTADScreen2!UnitsFlows(3) & ")"; Tab(VALUE_TAB); frmPTADScreen2!txtFlowsLoadings(3).Text
          Printer.Print frmPTADScreen2!lblFlowsLoadingsLabel(4).Caption & " (" & frmPTADScreen2!UnitsFlows(4) & ")"; Tab(VALUE_TAB); frmPTADScreen2!txtFlowsLoadings(4).Text
          Printer.Print
          Printer.Print frmPTADScreen2!lblDesignConcentration(7).Caption; Tab(VALUE_TAB); frmPTADScreen2!lblDesignConcentrationValue(7).Caption
          Printer.Print
          Printer.Print
          Printer.FontBold = True
          Call SetPowerPTADScreen2(CalculatedPower)
          Printer.Print "Power Calculation:"
          Printer.FontUnderline = True
          Printer.Print
          Printer.Print "Property:"; Tab(VALUE_TAB); "Value:"
          Printer.FontBold = False
          Printer.FontUnderline = False
          Printer.Print
          Printer.Print frmPowerScreen2!lblPowerLabel(0).Caption; Tab(VALUE_TAB); frmPowerScreen2!txtPower(0).Text
          Printer.Print frmPowerScreen2!lblPowerLabel(1).Caption; Tab(VALUE_TAB); frmPowerScreen2!txtPower(1).Text
          Printer.Print frmPowerScreen2!lblPowerLabel(2).Caption; Tab(VALUE_TAB); frmPowerScreen2!lblPower(2).Caption
          Printer.Print frmPowerScreen2!lblPowerLabel(3).Caption; Tab(VALUE_TAB); frmPowerScreen2!txtPower(3).Text
          Printer.Print frmPowerScreen2!lblPowerLabel(4).Caption; Tab(VALUE_TAB); frmPowerScreen2!lblPower(4).Caption
          Printer.Print frmPowerScreen2!lblPowerLabel(5).Caption; Tab(VALUE_TAB); frmPowerScreen2!lblPower(5).Caption
          Printer.Print
          Printer.Print
          Printer.FontUnderline = True
          Printer.Print "Contaminant Glossary:"
          Printer.FontUnderline = False
          For i = 1 To Scr2.NumChemical
              Printer.Print Format$(i, "0"); " = "; Trim$(Scr2.Contaminant(i).Name)
          Next i
          Call NewPagePTADScreen2
          Printer.FontBold = True
          Printer.Print "Contaminant Properties:"
          Printer.Print
          Printer.FontUnderline = True
          Printer.Print "Con.:"; Tab(MWT_TAB); "MWT"; Tab(HC_TAB); "HC"; Tab(VB_TAB); "Vb"; Tab(DIFL_TAB); "NBP"; Tab(MTCOEFF_TAB); "DIFL"; Tab(STANTON_TAB); "DIFG"
          Printer.Print
          Printer.FontUnderline = False
          Printer.FontBold = False
          For i = 1 To Scr2.NumChemical
              Printer.Print Format$(i, "0"); Tab(MWT_TAB); Format$(Scr2.Contaminant(i).MolecularWeight.value, "0.00"); Tab(HC_TAB); Format$(Scr2.Contaminant(i).HenrysConstant.value, GetTheFormat(Scr2.Contaminant(i).HenrysConstant.value)); Tab(VB_TAB); Format$(Scr2.Contaminant(i).MolarVolume.value, GetTheFormat(Scr2.Contaminant(i).MolarVolume.value)); Tab(DIFL_TAB); Format$(Scr2.Contaminant(i).NormalBoilingPoint.value - 273.15, GetTheFormat(Scr2.Contaminant(i).NormalBoilingPoint.value - 273.15)); Tab(MTCOEFF_TAB); Format$(Scr2.Contaminant(i).LiquidDiffusivity.value, GetTheFormat(Scr2.Contaminant(i).LiquidDiffusivity.value)); Tab(STANTON_TAB); Format$(Scr2.Contaminant(i).GasDiffusivity.value, GetTheFormat(Scr2.Contaminant(i).GasDiffusivity.value))
          Next i
          Printer.Print
          Printer.FontUnderline = True
          Printer.Print "Glossary:"
          Printer.FontUnderline = False
          Printer.Print "Con. = Contaminant Number (see Contaminant Glossary on page 1)"
          Printer.Print "MWT = Molecular Weight (g/gmol)"
          Printer.Print "HC = Henry's Constant (dimensionless)"
          Printer.Print "Vb = Molar Volume (m³/kmol)"
          Printer.Print "NBP = Normal Boiling Point (Celcius)"
          Printer.Print "DIFL = Liquid Diffusivity (m²/s)"
          Printer.Print "DIFG = Gas Diffusivity (m²/s)"
          Printer.Print
          Printer.Print
          Printer.Print
          Printer.FontBold = True
          Printer.Print "Contaminant Mass Transfer Parameters:"
          Printer.Print
          Printer.FontUnderline = True
          Printer.Print "Con.:"; Tab(MWT_TAB); "Onda KLa"; Tab(HC_TAB); "KLa SF"; Tab(VB_TAB); "Des. KLa"
          Printer.Print
          Printer.FontUnderline = False
          Printer.FontBold = False
          KLaSafetyFactor = Scr2.KLaSafetyFactor.value
          For i = 1 To Scr2.NumChemical
              If Scr2.DesignContaminant.Name = Scr2.Contaminant(i).Name Then
                 PackingWettedSurfaceArea(i) = Scr2.Packing.OndaWettedSurfaceArea
                 OndaKLa(i) = Scr2.Onda.OverallMassTransferCoefficient
                 DesignKLa(i) = Scr2.DesignMassTransferCoefficient.value
              Else
                 Call AWCALC(PackingWettedSurfaceArea(i), Scr2.Packing.CriticalSurfaceTension, Scr2.WaterSurfaceTension.value, Scr2.WaterLoadingRate.value, Scr2.Packing.SpecificSurfaceArea, Scr2.WaterViscosity.value, Scr2.WaterDensity.value, ReynoldsNumber, FroudeNumber, WeberNumber)
                 Call ONDAKLPT(LiquidPhaseMassTransferCoefficient, Scr2.WaterLoadingRate.value, PackingWettedSurfaceArea(i), Scr2.WaterViscosity.value, Scr2.WaterDensity.value, Scr2.Contaminant(i).LiquidDiffusivity.value, Scr2.Packing.SpecificSurfaceArea, Scr2.Packing.NominalSize)
                 Call ONDAKGPT(GasPhaseMassTransferCoefficient, Scr2.AirLoadingRate.value, Scr2.Packing.SpecificSurfaceArea, Scr2.AirViscosity.value, Scr2.AirDensity.value, Scr2.Contaminant(i).GasDiffusivity.value, Scr2.Packing.NominalSize)
                 Call ONDKLAPT(OndaKLa(i), LiquidPhaseMassTransferResistance, GasPhaseMassTransferResistance, TotalMassTransferResistance, LiquidPhaseMassTransferCoefficient, PackingWettedSurfaceArea(i), GasPhaseMassTransferCoefficient, Scr2.Contaminant(i).HenrysConstant.value)
                 Call KLACOR(DesignKLa(i), OndaKLa(i), KLaSafetyFactor)
              End If
              Printer.Print Format$(i, "0"); Tab(MWT_TAB); Format$(OndaKLa(i), GetTheFormat(OndaKLa(i))); Tab(HC_TAB); Format$(KLaSafetyFactor, GetTheFormat(KLaSafetyFactor)); Tab(VB_TAB); Format$(DesignKLa(i), GetTheFormat(DesignKLa(i)))
          Next i
          Printer.Print
          Printer.FontUnderline = True
          Printer.Print "Glossary:"
          Printer.FontUnderline = False
          Printer.Print "Con. = Contaminant Number (see Contaminant Glossary on page 1)"
          Printer.Print "Onda KLa = "; frmPTADScreen2!lblDesignConcentration(0).Caption & " (1/s)"
          Printer.Print "KLa SF = "; frmPTADScreen2!lblDesignConcentration(1).Caption & " (dimensionless)"
          Printer.Print "Des. KLa = "; frmPTADScreen2!lblDesignConcentration(2).Caption & " (1/s)"
          If Scr2.NumChemical > 6 Then
             Call NewPagePTADScreen2
          Else
             Printer.Print
             Printer.Print
             Printer.Print
          End If
          Printer.FontBold = True
          Printer.Print "Concentration Results:"
          Printer.Print
          Printer.FontUnderline = True
          Printer.Print "Con.:"; Tab(MWT_TAB); "Cinf"; Tab(HC_TAB); "Cto"; Tab(VB_TAB); "De. % Rem."; Tab(DIFL_TAB); "Ceff"; Tab(MTCOEFF_TAB); "Ach. % Rem."
          Printer.Print
          Printer.FontUnderline = False
          Printer.FontBold = False
          For i = 1 To Scr2.NumChemical
              If Scr2.DesignContaminant.Name = Scr2.Contaminant(i).Name Then
                 Call REMOVPT(DesiredPercentRemoval(i), Scr2.DesignContaminant.Influent.value, Scr2.DesignContaminant.TreatmentObjective.value)
                 Effluent(i) = Scr2.DesignContaminant.Effluent.value
                 Call REMOVPT(AchievedPercentRemoval(i), Scr2.DesignContaminant.Influent.value, Effluent(i))
              Else
                 Call REMOVPT(DesiredPercentRemoval(i), Scr2.Contaminant(i).Influent.value, Scr2.Contaminant(i).TreatmentObjective.value)
                 Call EFFLPT2(Effluent(i), Scr2.AirToWaterRatio.value, Scr2.Contaminant(i).HenrysConstant.value, Scr2.WaterFlowRate.value, Scr2.TowerArea.value, Scr2.SpecifiedTowerHeight.value, DesignKLa(i), Scr2.Contaminant(i).Influent.value)
                 Call REMOVPT(AchievedPercentRemoval(i), Scr2.Contaminant(i).Influent.value, Effluent(i))
              End If
              Printer.Print Format$(i, "0"); Tab(MWT_TAB); Format$(Scr2.Contaminant(i).Influent.value, GetTheFormat(Scr2.Contaminant(i).Influent.value)); Tab(HC_TAB); Format$(Scr2.Contaminant(i).TreatmentObjective.value, GetTheFormat(Scr2.Contaminant(i).TreatmentObjective.value)); Tab(VB_TAB); Format$(DesiredPercentRemoval(i), GetTheFormat(DesiredPercentRemoval(i))); Tab(DIFL_TAB); Format$(Effluent(i), GetTheFormat(Effluent(i))); Tab(MTCOEFF_TAB); Format$(AchievedPercentRemoval(i), GetTheFormat(AchievedPercentRemoval(i)))
          Next i
          Printer.Print
          Printer.FontUnderline = True
          Printer.Print "Glossary:"
          Printer.FontUnderline = False
          Printer.Print "Con. = Contaminant Number (see Contaminant Glossary on page 1)"
          Printer.Print "Cinf = "; frmPTADScreen2!lblDesignConcentration(3).Caption & " (µg/L)"
          Printer.Print "Cto = "; frmPTADScreen2!lblDesignConcentration(4).Caption & " (µg/L)"
          Printer.Print "De. % Rem. = "; "Desired Percent Removal"
          Printer.Print "Ceff = "; frmPTADScreen2!lblDesignConcentration(5).Caption & " (µg/L)"
          Printer.Print "Ach. % Rem. = "; "Achieved Percent Removal"

          Printer.EndDoc

    Exit Sub

PrinterError:
    MsgBox error$(Err)
    Resume ExitPrint:

ExitPrint:

End Sub

Sub PrintPTADScreen2ToFile()
    Dim i As Integer, j As Integer
    Dim CalculatedPower As Integer
    ReDim OndaKLa(1 To MAXCHEMICAL) As Double
    Dim KLaSafetyFactor As Double
    ReDim DesignKLa(1 To MAXCHEMICAL) As Double
    ReDim PackingWettedSurfaceArea(1 To MAXCHEMICAL) As Double
    Dim ReynoldsNumber As Double
    Dim FroudeNumber As Double
    Dim WeberNumber As Double
    Dim LiquidPhaseMassTransferCoefficient As Double
    Dim GasPhaseMassTransferCoefficient As Double
    Dim LiquidPhaseMassTransferResistance As Double
    Dim GasPhaseMassTransferResistance As Double
    Dim TotalMassTransferResistance As Double
    ReDim DesiredPercentRemoval(1 To MAXCHEMICAL) As Double
    ReDim Effluent(1 To MAXCHEMICAL) As Double
    ReDim AchievedPercentRemoval(1 To MAXCHEMICAL) As Double

        Call GetPrintFileName(PrintFileName)
        If PrintFileName$ = "" Then Exit Sub

        Open PrintFileName$ For Output As #1

          Print #1, "Packed Tower Aeration - Rating Mode"
          Print #1,
          Print #1,
          Print #1, "Property:"; Tab(VALUE_TAB); "Value:"
          Print #1,
          If ShownScreen1Previously Then
             Print #1, "Design based on:  "; frmPTADScreen2!lblDesignParametersLabel(0).Caption & " (" & frmPTADScreen2!UnitsDesignBasis(0) & ")"; Tab(VALUE_TAB); frmPTADScreen2!lblDesignParameters(0).Caption
             Print #1, "Design based on:  "; frmPTADScreen2!lblDesignParametersLabel(1).Caption & " (" & frmPTADScreen2!UnitsDesignBasis(1) & ")"; Tab(VALUE_TAB); frmPTADScreen2!lblDesignParameters(1).Caption
          End If
          Print #1, frmPTADScreen2!lblTowerParametersLabel(0).Caption & " (" & frmPTADScreen2!UnitsTowerParam(0) & ")"; Tab(VALUE_TAB); frmPTADScreen2!txtTowerParameters(0).Text
          Print #1, frmPTADScreen2!lblTowerParametersLabel(1).Caption & " (" & frmPTADScreen2!UnitsTowerParam(1) & ")"; Tab(VALUE_TAB); frmPTADScreen2!txtTowerParameters(1).Text
          Print #1, frmPTADScreen2!lblTowerParametersLabel(2).Caption & " (" & frmPTADScreen2!UnitsTowerParam(2) & ")"; Tab(VALUE_TAB); frmPTADScreen2!lblTowerParameters(2).Caption
          Print #1, frmPTADScreen2!lblTowerParametersLabel(3).Caption & " (" & frmPTADScreen2!UnitsTowerParam(3) & ")"; Tab(VALUE_TAB); frmPTADScreen2!lblTowerParameters(3).Caption
          Printer.Print
          Print #1, "Operating Pressure" & " (" & frmPTADScreen2!UnitsOpCond(0) & ")"; Tab(VALUE_TAB); frmPTADScreen2!txtOperatingPressure.Text
          Print #1, "Operating Temperature" & " (" & frmPTADScreen2!UnitsOpCond(1) & ")"; Tab(VALUE_TAB); frmPTADScreen2!txtOperatingTemperature.Text
          Print #1, frmAirWaterProperties!lblAirWaterProperties(0).Caption; Tab(VALUE_TAB); frmAirWaterProperties!txtAirWaterProperties(0).Text
          Print #1, frmAirWaterProperties!lblAirWaterProperties(1).Caption; Tab(VALUE_TAB); frmAirWaterProperties!txtAirWaterProperties(1).Text
          Print #1, frmAirWaterProperties!lblAirWaterProperties(2).Caption; Tab(VALUE_TAB); frmAirWaterProperties!txtAirWaterProperties(2).Text
          Print #1, frmAirWaterProperties!lblAirWaterProperties(3).Caption; Tab(VALUE_TAB); frmAirWaterProperties!txtAirWaterProperties(3).Text
          Print #1, frmAirWaterProperties!lblAirWaterProperties(4).Caption; Tab(VALUE_TAB); frmAirWaterProperties!txtAirWaterProperties(4).Text
          Print #1,
          CurrentScreen = Scr2
          Print #1, "Packing Name:  "; Trim$(CurrentScreen.Packing.Name)
          Print #1, "Packing "; frmSelectPacking!lblPackingProperties(1).Caption; Tab(VALUE_TAB); Format$(CurrentScreen.Packing.NominalSize, GetTheFormat(CurrentScreen.Packing.NominalSize))
          Print #1, frmSelectPacking!lblPackingProperties(2).Caption; Tab(VALUE_TAB); Format$(CurrentScreen.Packing.PackingFactor, GetTheFormat(CurrentScreen.Packing.PackingFactor))
          Print #1, "Packing "; frmSelectPacking!lblPackingProperties(3).Caption; Tab(VALUE_TAB); Format$(CurrentScreen.Packing.SpecificSurfaceArea, GetTheFormat(CurrentScreen.Packing.SpecificSurfaceArea))
          Print #1, "Packing "; frmSelectPacking!lblPackingProperties(4).Caption; Tab(VALUE_TAB); Format$(CurrentScreen.Packing.CriticalSurfaceTension, GetTheFormat(CurrentScreen.Packing.CriticalSurfaceTension))
          Print #1, "Packing "; frmSelectPacking!lblPackingProperties(5).Caption; Tab(VALUE_TAB); Trim$(CurrentScreen.Packing.Material)
          Print #1, "Packing "; frmSelectPacking!lblPackingProperties(6).Caption; Tab(VALUE_TAB); Trim$(CurrentScreen.Packing.source)
          Print #1, "Source of This Packing Data in Program"; Tab(VALUE_TAB);
          If PackingDatabaseSource = ORIGINALPACKINGDATABASE Then
             Print #1, "Original Packing Database"
          Else
             Print #1, "User Input"
          End If
          
          Print #1,
          Print #1, frmPTADScreen2!lblFlowsLoadingsLabel(0).Caption & " (" & frmPTADScreen2!UnitsFlows(0) & ")"; Tab(VALUE_TAB); frmPTADScreen2!txtFlowsLoadings(0).Text
          Print #1, frmPTADScreen2!lblFlowsLoadingsLabel(1).Caption & " (" & frmPTADScreen2!UnitsFlows(1) & ")"; Tab(VALUE_TAB); frmPTADScreen2!txtFlowsLoadings(1).Text
          Print #1, frmPTADScreen2!lblFlowsLoadingsLabel(2).Caption & " (-)"; Tab(VALUE_TAB); frmPTADScreen2!txtFlowsLoadings(2).Text
          Print #1, frmPTADScreen2!lblFlowsLoadingsLabel(3).Caption & " (" & frmPTADScreen2!UnitsFlows(3) & ")"; Tab(VALUE_TAB); frmPTADScreen2!txtFlowsLoadings(3).Text
          Print #1, frmPTADScreen2!lblFlowsLoadingsLabel(4).Caption & " (" & frmPTADScreen2!UnitsFlows(4) & ")"; Tab(VALUE_TAB); frmPTADScreen2!txtFlowsLoadings(4).Text
          Print #1,
          Print #1, frmPTADScreen2!lblDesignConcentration(7).Caption; Tab(VALUE_TAB); frmPTADScreen2!lblDesignConcentrationValue(7).Caption
          Print #1,
          Print #1,
          Call SetPowerPTADScreen2(CalculatedPower)
          Print #1, "Power Calculation:"
          Print #1,
          Print #1, "Property:"; Tab(VALUE_TAB); "Value:"
          Print #1,
          Print #1, frmPowerScreen2!lblPowerLabel(0).Caption; Tab(VALUE_TAB); frmPowerScreen2!txtPower(0).Text
          Print #1, frmPowerScreen2!lblPowerLabel(1).Caption; Tab(VALUE_TAB); frmPowerScreen2!txtPower(1).Text
          Print #1, frmPowerScreen2!lblPowerLabel(2).Caption; Tab(VALUE_TAB); frmPowerScreen2!lblPower(2).Caption
          Print #1, frmPowerScreen2!lblPowerLabel(3).Caption; Tab(VALUE_TAB); frmPowerScreen2!txtPower(3).Text
          Print #1, frmPowerScreen2!lblPowerLabel(4).Caption; Tab(VALUE_TAB); frmPowerScreen2!lblPower(4).Caption
          Print #1, frmPowerScreen2!lblPowerLabel(5).Caption; Tab(VALUE_TAB); frmPowerScreen2!lblPower(5).Caption
          Print #1,
          Print #1,
          Print #1, "Contaminant Glossary:"
          For i = 1 To Scr2.NumChemical
              Print #1, Format$(i, "0"); " = "; Trim$(Scr2.Contaminant(i).Name)
          Next i
          Print #1,
          Print #1,
          Print #1,
          Print #1, "Contaminant Properties:"
          Print #1,
          Print #1, "Con.:"; Tab(MWT_TAB); "MWT"; Tab(HC_TAB); "HC"; Tab(VB_TAB); "Vb"; Tab(DIFL_TAB); "NBP"; Tab(MTCOEFF_TAB); "DIFL"; Tab(STANTON_TAB); "DIFG"
          Print #1,
          For i = 1 To Scr2.NumChemical
              Print #1, Format$(i, "0"); Tab(MWT_TAB); Format$(Scr2.Contaminant(i).MolecularWeight.value, "0.00"); Tab(HC_TAB); Format$(Scr2.Contaminant(i).HenrysConstant.value, GetTheFormat(Scr2.Contaminant(i).HenrysConstant.value)); Tab(VB_TAB); Format$(Scr2.Contaminant(i).MolarVolume.value, GetTheFormat(Scr2.Contaminant(i).MolarVolume.value)); Tab(DIFL_TAB); Format$(Scr2.Contaminant(i).NormalBoilingPoint.value - 273.15, GetTheFormat(Scr2.Contaminant(i).NormalBoilingPoint.value - 273.15)); Tab(MTCOEFF_TAB); Format$(Scr2.Contaminant(i).LiquidDiffusivity.value, GetTheFormat(Scr2.Contaminant(i).LiquidDiffusivity.value)); Tab(STANTON_TAB); Format$(Scr2.Contaminant(i).GasDiffusivity.value, GetTheFormat(Scr2.Contaminant(i).GasDiffusivity.value))
          Next i
          Print #1,
         
          Print #1, "Glossary:"
          Print #1, "Con. = Contaminant Number (see Contaminant Glossary on page 1)"
          Print #1, "MWT = Molecular Weight (g/gmol)"
          Print #1, "HC = Henry's Constant (dimensionless)"
          Print #1, "Vb = Molar Volume (m³/kmol)"
          Print #1, "NBP = Normal Boiling Point (Celcius)"
          Print #1, "DIFL = Liquid Diffusivity (m²/s)"
          Print #1, "DIFG = Gas Diffusivity (m²/s)"
          Print #1,
          Print #1,
          Print #1,
          Print #1, "Contaminant Mass Transfer Parameters:"
          Print #1,
          Print #1, "Con.:"; Tab(MWT_TAB); "Onda KLa"; Tab(HC_TAB); "KLa SF"; Tab(VB_TAB); "Des. KLa"
          Print #1,
          KLaSafetyFactor = Scr2.KLaSafetyFactor.value
          For i = 1 To Scr2.NumChemical
              If Scr2.DesignContaminant.Name = Scr2.Contaminant(i).Name Then
                 PackingWettedSurfaceArea(i) = Scr2.Packing.OndaWettedSurfaceArea
                 OndaKLa(i) = Scr2.Onda.OverallMassTransferCoefficient
                 DesignKLa(i) = Scr2.DesignMassTransferCoefficient.value
              Else
                 Call AWCALC(PackingWettedSurfaceArea(i), Scr2.Packing.CriticalSurfaceTension, Scr2.WaterSurfaceTension.value, Scr2.WaterLoadingRate.value, Scr2.Packing.SpecificSurfaceArea, Scr2.WaterViscosity.value, Scr2.WaterDensity.value, ReynoldsNumber, FroudeNumber, WeberNumber)
                 Call ONDAKLPT(LiquidPhaseMassTransferCoefficient, Scr2.WaterLoadingRate.value, PackingWettedSurfaceArea(i), Scr2.WaterViscosity.value, Scr2.WaterDensity.value, Scr2.Contaminant(i).LiquidDiffusivity.value, Scr2.Packing.SpecificSurfaceArea, Scr2.Packing.NominalSize)
                 Call ONDAKGPT(GasPhaseMassTransferCoefficient, Scr2.AirLoadingRate.value, Scr2.Packing.SpecificSurfaceArea, Scr2.AirViscosity.value, Scr2.AirDensity.value, Scr2.Contaminant(i).GasDiffusivity.value, Scr2.Packing.NominalSize)
                 Call ONDKLAPT(OndaKLa(i), LiquidPhaseMassTransferResistance, GasPhaseMassTransferResistance, TotalMassTransferResistance, LiquidPhaseMassTransferCoefficient, PackingWettedSurfaceArea(i), GasPhaseMassTransferCoefficient, Scr2.Contaminant(i).HenrysConstant.value)
                 Call KLACOR(DesignKLa(i), OndaKLa(i), KLaSafetyFactor)
              End If
              Print #1, Format$(i, "0"); Tab(MWT_TAB); Format$(OndaKLa(i), GetTheFormat(OndaKLa(i))); Tab(HC_TAB); Format$(KLaSafetyFactor, GetTheFormat(KLaSafetyFactor)); Tab(VB_TAB); Format$(DesignKLa(i), GetTheFormat(DesignKLa(i)))
          Next i
          Print #1,
          Print #1, "Glossary:"
          Print #1, "Con. = Contaminant Number (see Contaminant Glossary on page 1)"
          Print #1, "Onda KLa = "; frmPTADScreen2!lblDesignConcentration(0).Caption & " (1/s)"
          Print #1, "KLa SF = "; frmPTADScreen2!lblDesignConcentration(1).Caption & " (dimensionless)"
          Print #1, "Des. KLa = "; frmPTADScreen2!lblDesignConcentration(2).Caption & " (1/s)"
             Print #1,
             Print #1,
             Print #1,
          Print #1, "Concentration Results:"
          Print #1,
          Print #1, "Con.:"; Tab(MWT_TAB); "Cinf"; Tab(HC_TAB); "Cto"; Tab(VB_TAB); "De. % Rem."; Tab(DIFL_TAB); "Ceff"; Tab(MTCOEFF_TAB); "Ach. % Rem."
          Print #1,
          For i = 1 To Scr2.NumChemical
              If Scr2.DesignContaminant.Name = Scr2.Contaminant(i).Name Then
                 Call REMOVPT(DesiredPercentRemoval(i), Scr2.DesignContaminant.Influent.value, Scr2.DesignContaminant.TreatmentObjective.value)
                 Effluent(i) = Scr2.DesignContaminant.Effluent.value
                 Call REMOVPT(AchievedPercentRemoval(i), Scr2.DesignContaminant.Influent.value, Effluent(i))
              Else
                 Call REMOVPT(DesiredPercentRemoval(i), Scr2.Contaminant(i).Influent.value, Scr2.Contaminant(i).TreatmentObjective.value)
                 Call EFFLPT2(Effluent(i), Scr2.AirToWaterRatio.value, Scr2.Contaminant(i).HenrysConstant.value, Scr2.WaterFlowRate.value, Scr2.TowerArea.value, Scr2.SpecifiedTowerHeight.value, DesignKLa(i), Scr2.Contaminant(i).Influent.value)
                 Call REMOVPT(AchievedPercentRemoval(i), Scr2.Contaminant(i).Influent.value, Effluent(i))
              End If
              Print #1, Format$(i, "0"); Tab(MWT_TAB); Format$(Scr2.Contaminant(i).Influent.value, GetTheFormat(Scr2.Contaminant(i).Influent.value)); Tab(HC_TAB); Format$(Scr2.Contaminant(i).TreatmentObjective.value, GetTheFormat(Scr2.Contaminant(i).TreatmentObjective.value)); Tab(VB_TAB); Format$(DesiredPercentRemoval(i), GetTheFormat(DesiredPercentRemoval(i))); Tab(DIFL_TAB); Format$(Effluent(i), GetTheFormat(Effluent(i))); Tab(MTCOEFF_TAB); Format$(AchievedPercentRemoval(i), GetTheFormat(AchievedPercentRemoval(i)))
          Next i
          Print #1,
          Print #1, "Glossary:"
          Print #1, "Con. = Contaminant Number (see Contaminant Glossary on page 1)"
          Print #1, "Cinf = "; frmPTADScreen2!lblDesignConcentration(3).Caption & " (µg/L)"
          Print #1, "Cto = "; frmPTADScreen2!lblDesignConcentration(4).Caption & " (µg/L)"
          Print #1, "De. % Rem. = "; "Desired Percent Removal"
          Print #1, "Ceff = "; frmPTADScreen2!lblDesignConcentration(5).Caption & " (µg/L)"
          Print #1, "Ach. % Rem. = "; "Achieved Percent Removal"
          Close #1

End Sub

Sub SaveContaminantListScreen2()
    Dim FileID As String
    Dim i As Integer

    Call SaveFile(Filename)

    If Filename$ <> "" Then
       FileID = CONTAMINANTS_PTAD_FILEID
       Open Filename$ For Output As #1
       
       Write #1, FileID
      
       For i = 1 To Scr2.NumChemical
           Write #1, Scr2.Contaminant(i).Pressure, Scr2.Contaminant(i).Temperature, Scr2.Contaminant(i).Name, Scr2.Contaminant(i).MolecularWeight.value, Scr2.Contaminant(i).HenrysConstant.value, Scr2.Contaminant(i).MolarVolume.value, Scr2.Contaminant(i).NormalBoilingPoint.value, Scr2.Contaminant(i).LiquidDiffusivity.value, Scr2.Contaminant(i).GasDiffusivity.value, Scr2.Contaminant(i).Influent.value, Scr2.Contaminant(i).TreatmentObjective.value
       Next i

       Close #1

    End If

End Sub

Sub savefilescreen2(Filename As String)
Dim Ctl As Control
Set Ctl = frmPTADScreen2.CommonDialog1

    On Error Resume Next
    'frmPTADScreen2!CMDialog1.DefaultExt = "rat"
    'frmPTADScreen2!CMDialog1.Filter = "Rating Mode Files (*.rat)|*.rat"
    'frmPTADScreen2!CMDialog1.DialogTitle = "Save Packed Tower Aeration Rating Mode File"
    'frmPTADScreen2!CMDialog1.flags = OFN_OVERWRITEPROMPT Or OFN_PATHMUSTEXIST
    'frmPTADScreen2!CMDialog1.Action = 2
    'Filename$ = frmPTADScreen2!CMDialog1.Filename
    Ctl.DefaultExt = "rat"
    Ctl.Filter = "Rating Mode Files (*.rat)|*.rat"
    Ctl.DialogTitle = "Save Packed Tower Aeration Rating Mode File"
    Ctl.flags = OFN_OVERWRITEPROMPT Or OFN_PATHMUSTEXIST
    Ctl.Action = 2
    Filename$ = Ctl.Filename
    If Err = 32755 Then   'Cancel selected by user
       Filename$ = ""
    End If

End Sub

Sub SaveScreen2()
Dim FileID As String
Dim i As Integer
Dim xu As rec_Units_frmContaminantPropertyEdit

  If (IsThisADemo() = True) Then
    Call Demo_ShowError("Saving is not allowed in the demonstration version.")
    Exit Sub
  End If
    
    If Right$(frmPTADScreen2.Caption, 14) = "(untitled.rat)" Then
       Call savefilescreen2(Filename)
    End If

    If Filename$ <> "" Then
       FileID = SCREEN2_PTAD2_FILEID
       Open Filename$ For Output As #1
       
       Write #1, FileID
      
       Write #1, Scr2.TowerDiameter.value
       Write #1, Scr2.TowerHeight.value
       Write #1, Scr2.SpecifiedTowerDiameter.value
       Write #1, Scr2.SpecifiedTowerHeight.value
       Write #1, Scr2.OperatingPressure.value
       Write #1, Scr2.operatingtemperature.value
       Write #1, Scr2.Packing.Name, Scr2.Packing.NominalSize, Scr2.Packing.PackingFactor, Scr2.Packing.SpecificSurfaceArea, Scr2.Packing.CriticalSurfaceTension, Scr2.Packing.Material, Scr2.Packing.source, Scr2.Packing.UserInput, Scr2.Packing.SourceDatabase
       If frmFlowsLoadingsScreen2!optFlowsLoadings(0).value = True Then
          Write #1, "Specified Water Flow Rate and Air Flow Rate"
          Write #1, Scr2.WaterFlowRate.value
          Write #1, Scr2.AirFlowRate.value
       ElseIf frmFlowsLoadingsScreen2!optFlowsLoadings(1).value = True Then
          Write #1, "Specified Water Flow Rate and Air to Water Ratio"
          Write #1, Scr2.WaterFlowRate.value
          Write #1, Scr2.AirToWaterRatio.value
       ElseIf frmFlowsLoadingsScreen2!optFlowsLoadings(2).value = True Then
          Write #1, "Specified Water Loading Rate and Air Loading Rate"
          Write #1, Scr2.WaterLoadingRate.value
          Write #1, Scr2.AirLoadingRate.value
       End If

       Write #1, Scr2.NumChemical
       For i = 1 To Scr2.NumChemical
           Write #1, Scr2.Contaminant(i).Pressure, Scr2.Contaminant(i).Temperature, Scr2.Contaminant(i).Name, Scr2.Contaminant(i).MolecularWeight.value, Scr2.Contaminant(i).HenrysConstant.value, Scr2.Contaminant(i).MolarVolume.value, Scr2.Contaminant(i).NormalBoilingPoint.value, Scr2.Contaminant(i).LiquidDiffusivity.value, Scr2.Contaminant(i).GasDiffusivity.value, Scr2.Contaminant(i).Influent.value, Scr2.Contaminant(i).TreatmentObjective.value
       Next i
       Write #1, Scr2.DesignContaminant.Name

       Write #1, Scr2.KLaSafetyFactor.value, Scr2.KLaSafetyFactor.UserInput
       Write #1, Scr2.DesignMassTransferCoefficient.value, Scr2.DesignMassTransferCoefficient.UserInput

       'Output the units of this screen.
       Write #1, GetUnits(frmPTADScreen2!UnitsDesignBasis(0)), GetUnits(frmPTADScreen2!UnitsDesignBasis(1))
       Write #1, GetUnits(frmPTADScreen2!UnitsTowerParam(0)), GetUnits(frmPTADScreen2!UnitsTowerParam(1)), GetUnits(frmPTADScreen2!UnitsTowerParam(2)), GetUnits(frmPTADScreen2!UnitsTowerParam(3))
       Write #1, GetUnits(frmPTADScreen2!UnitsOpCond(0)), GetUnits(frmPTADScreen2!UnitsOpCond(1))
       Write #1, GetUnits(frmPTADScreen2!UnitsFlows(0)), GetUnits(frmPTADScreen2!UnitsFlows(1)), GetUnits(frmPTADScreen2!UnitsFlows(3)), GetUnits(frmPTADScreen2!UnitsFlows(4))
       Write #1, GetUnits(frmPTADScreen2!UnitsInterest(0)), GetUnits(frmPTADScreen2!UnitsInterest(2)), GetUnits(frmPTADScreen2!UnitsInterest(3)), GetUnits(frmPTADScreen2!UnitsInterest(4)), GetUnits(frmPTADScreen2!UnitsInterest(5)), GetUnits(frmPTADScreen2!UnitsInterest(7))

       'Output the units of frmContaminantPropertyEdit.
       xu = Units_frmContaminantPropertyEdit
       Write #1, xu.UnitsProp(0), xu.UnitsProp(2), xu.UnitsProp(3), xu.UnitsProp(4), xu.UnitsProp(5)
       Write #1, xu.UnitsConc(0), xu.UnitsConc(1)
       
       Close #1

       frmPTADScreen2.Caption = "Packed Tower Aeration - Rating Mode"
       frmPTADScreen2.Caption = frmPTADScreen2.Caption & " (" & Filename & ")"

    End If

End Sub

Sub screen2_results()
    Dim i As Integer, j As Integer
    Dim ContaminantGlossaryBottom As Integer, GlossaryBottom As Integer
    Dim KLaSafetyFactor As Double
    Dim ReynoldsNumber As Double
    Dim FroudeNumber As Double
    Dim WeberNumber As Double
    Dim LiquidPhaseMassTransferCoefficient As Double
    Dim GasPhaseMassTransferCoefficient As Double
    Dim LiquidPhaseMassTransferResistance As Double
    Dim GasPhaseMassTransferResistance As Double
    Dim TotalMassTransferResistance As Double
    
    ReDim DesiredPercentRemoval(1 To MAXCHEMICAL) As Double
    ReDim Effluent(1 To MAXCHEMICAL) As Double
    ReDim AchievedPercentRemoval(1 To MAXCHEMICAL) As Double
    ReDim OndaKLa(1 To MAXCHEMICAL) As Double
    ReDim DesignKLa(1 To MAXCHEMICAL) As Double
    ReDim PackingWettedSurfaceArea(1 To MAXCHEMICAL) As Double
          
          
          KLaSafetyFactor = Scr2.KLaSafetyFactor.value
          For i = 1 To Scr2.NumChemical
              If Scr2.DesignContaminant.Name = Scr2.Contaminant(i).Name Then
                 PackingWettedSurfaceArea(i) = Scr2.Packing.OndaWettedSurfaceArea
                 OndaKLa(i) = Scr2.Onda.OverallMassTransferCoefficient
                 DesignKLa(i) = Scr2.DesignMassTransferCoefficient.value
                 Call REMOVPT(DesiredPercentRemoval(i), Scr2.DesignContaminant.Influent.value, Scr2.DesignContaminant.TreatmentObjective.value)
                 Effluent(i) = Scr2.DesignContaminant.Effluent.value
                 Call REMOVPT(AchievedPercentRemoval(i), Scr2.DesignContaminant.Influent.value, Effluent(i))
              Else
                 Call AWCALC(PackingWettedSurfaceArea(i), Scr2.Packing.CriticalSurfaceTension, Scr2.WaterSurfaceTension.value, Scr2.WaterLoadingRate.value, Scr2.Packing.SpecificSurfaceArea, Scr2.WaterViscosity.value, Scr2.WaterDensity.value, ReynoldsNumber, FroudeNumber, WeberNumber)
                 Call ONDAKLPT(LiquidPhaseMassTransferCoefficient, Scr2.WaterLoadingRate.value, PackingWettedSurfaceArea(i), Scr2.WaterViscosity.value, Scr2.WaterDensity.value, Scr2.Contaminant(i).LiquidDiffusivity.value, Scr2.Packing.SpecificSurfaceArea, Scr2.Packing.NominalSize)
                 Call ONDAKGPT(GasPhaseMassTransferCoefficient, Scr2.AirLoadingRate.value, Scr2.Packing.SpecificSurfaceArea, Scr2.AirViscosity.value, Scr2.AirDensity.value, Scr2.Contaminant(i).GasDiffusivity.value, Scr2.Packing.NominalSize)
                 Call ONDKLAPT(OndaKLa(i), LiquidPhaseMassTransferResistance, GasPhaseMassTransferResistance, TotalMassTransferResistance, LiquidPhaseMassTransferCoefficient, PackingWettedSurfaceArea(i), GasPhaseMassTransferCoefficient, Scr2.Contaminant(i).HenrysConstant.value)
                 Call KLACOR(DesignKLa(i), OndaKLa(i), KLaSafetyFactor)
                 Call REMOVPT(DesiredPercentRemoval(i), Scr2.Contaminant(i).Influent.value, Scr2.Contaminant(i).TreatmentObjective.value)
                 Call EFFLPT2(Effluent(i), Scr2.AirToWaterRatio.value, Scr2.Contaminant(i).HenrysConstant.value, Scr2.WaterFlowRate.value, Scr2.TowerArea.value, Scr2.SpecifiedTowerHeight.value, DesignKLa(i), Scr2.Contaminant(i).Influent.value)
                 Call REMOVPT(AchievedPercentRemoval(i), Scr2.Contaminant(i).Influent.value, Effluent(i))
              End If
          Next i

    For i = 0 To MAXCHEMICAL - 1
        frmViewEffluentConcentrationsASAP!lblContaminantNumber(i).Visible = False
        frmViewEffluentConcentrationsASAP!lblContaminantNumber(i + 10).Visible = False
        frmViewEffluentConcentrationsASAP!lblInfluentConcentration(i).Visible = False
        frmViewEffluentConcentrationsASAP!lblTreatmentObjective(i).Visible = False
        frmViewEffluentConcentrationsASAP!lblDesiredPercentRemoval(i).Visible = False
        frmViewEffluentConcentrationsASAP!lblEffluentConcentration(i).Visible = False
        frmViewEffluentConcentrationsASAP!lblAchievedPercentRemoval(i).Visible = False
        frmViewEffluentConcentrationsASAP!lblContaminantName(i).Visible = False

    Next i

    For i = 1 To Scr2.NumChemical
        frmViewEffluentConcentrationsASAP!lblContaminantNumber(i - 1).Visible = True
        frmViewEffluentConcentrationsASAP!lblContaminantNumber(i + 10 - 1).Visible = True
        frmViewEffluentConcentrationsASAP!lblInfluentConcentration(i - 1).Visible = True
        frmViewEffluentConcentrationsASAP!lblTreatmentObjective(i - 1).Visible = True
        frmViewEffluentConcentrationsASAP!lblDesiredPercentRemoval(i - 1).Visible = True
        frmViewEffluentConcentrationsASAP!lblEffluentConcentration(i - 1).Visible = True
        frmViewEffluentConcentrationsASAP!lblAchievedPercentRemoval(i - 1).Visible = True
        frmViewEffluentConcentrationsASAP!lblContaminantName(i - 1).Visible = True

        frmViewEffluentConcentrationsASAP!lblInfluentConcentration(i - 1).Caption = Format$(Scr2.Contaminant(i).Influent.value, GetTheFormat(Scr2.Contaminant(i).Influent.value))
        frmViewEffluentConcentrationsASAP!lblTreatmentObjective(i - 1).Caption = Format$(Scr2.Contaminant(i).TreatmentObjective.value, GetTheFormat(Scr2.Contaminant(i).TreatmentObjective.value))
        frmViewEffluentConcentrationsASAP!lblDesiredPercentRemoval(i - 1).Caption = Format$(DesiredPercentRemoval(i), "0.0")
        frmViewEffluentConcentrationsASAP!lblEffluentConcentration(i - 1).Caption = Format$(Effluent(i), GetTheFormat(Effluent(i)))
        frmViewEffluentConcentrationsASAP!lblAchievedPercentRemoval(i - 1).Caption = Format$(AchievedPercentRemoval(i), "0.0")
        frmViewEffluentConcentrationsASAP!lblContaminantName(i - 1).Caption = Trim$(LCase$(Scr2.Contaminant(i).Name))

    Next i

    frmViewEffluentConcentrationsASAP!fraConcentrationResults.Height = frmViewEffluentConcentrationsASAP!lblContaminantNumber(Scr2.NumChemical - 1).Top + frmViewEffluentConcentrationsASAP!lblContaminantNumber(Scr2.NumChemical - 1).Height + 120
    frmViewEffluentConcentrationsASAP!fraContaminantGlossary.Height = frmViewEffluentConcentrationsASAP!lblContaminantNumber(Scr2.NumChemical + 10 - 1).Top + frmViewEffluentConcentrationsASAP!lblContaminantNumber(Scr2.NumChemical + 10 - 1).Height + 120
    frmViewEffluentConcentrationsASAP!fraContaminantGlossary.Top = frmViewEffluentConcentrationsASAP!fraConcentrationResults.Top + frmViewEffluentConcentrationsASAP!fraConcentrationResults.Height + 120
    frmViewEffluentConcentrationsASAP!fraGlossary.Top = frmViewEffluentConcentrationsASAP!fraContaminantGlossary.Top
    ContaminantGlossaryBottom = frmViewEffluentConcentrationsASAP!fraContaminantGlossary.Top + frmViewEffluentConcentrationsASAP!fraContaminantGlossary.Height
    GlossaryBottom = frmViewEffluentConcentrationsASAP!fraGlossary.Top + frmViewEffluentConcentrationsASAP!fraGlossary.Height
    If GlossaryBottom > ContaminantGlossaryBottom Then
       frmViewEffluentConcentrationsASAP!cmdOK.Top = GlossaryBottom + 360
    Else
       frmViewEffluentConcentrationsASAP!cmdOK.Top = ContaminantGlossaryBottom + 360
    End If
    frmViewEffluentConcentrationsASAP.Height = frmViewEffluentConcentrationsASAP!cmdOK.Top + frmViewEffluentConcentrationsASAP!cmdOK.Height + 500   '420

    frmViewEffluentConcentrationsASAP.Show 1


End Sub

Function screen2_savechanges() As Integer
Dim i As Integer, Response As Integer
Dim msg As String

msg = "Would you like to save the parameters "
msg = msg + "for this rating case to a file "
msg = msg + "?" & Chr$(13) & Chr$(13)
msg = msg + "Note:  Any information not saved will be permanently lost."
Response = MsgBox(msg, MB_ICONquestion + MB_YESNOCANCEL, "Save Current Design")
                
If Response = IDCANCEL Then
 Screen.MousePointer = 0
 screen2_savechanges = 1
 Exit Function
End If
              
If Response = IDYES Then
   Call SaveScreen2
                    
   If StrComp(Filename, "") = 0 Then Response = 5
      
      Do While Response = 5
         msg = "Would you like to save the parameters "
         msg = msg + "for this rating case to a file "
         msg = msg + "?" & Chr$(13) & Chr$(13)
         msg = msg + "Note:  Any information not saved will be permanently lost."
         Response = MsgBox(msg, MB_ICONquestion + MB_YESNOCANCEL, "Save Current Design")
                         
         If Response = IDCANCEL Then
            Screen.MousePointer = 0
            screen2_savechanges = 1
            Exit Function
        End If
                        
        If Response = IDYES Then Call SaveScreen2
        If StrComp(Filename, "") = 0 And Response <> IDNO Then Response = 5
      Loop
End If

End Function

Sub SetDesignContaminantEnabledScreen2(NumInList As Long)
    Dim i As Integer

    If NumInList = 0 Then
       frmPTADScreen2!mnuFile(4).Enabled = False
       frmPTADScreen2!mnuFile(5).Enabled = False
       frmPTADScreen2!mnuOptions(0).Enabled = False
       'frmPTADScreen2!fraDesignContaminant.Enabled = False
       frmPTADScreen2!cboSelectCompo.Enabled = False
       Scr2.AirPressureDrop.value = -1#

       For i = 0 To 7
           Select Case i
              Case 0, 3 To 7
                 frmPTADScreen2!lblDesignConcentrationValue(i).Caption = "0.0"
                 frmPTADScreen2!lblDesignConcentrationValue(i).Enabled = False
              Case 1 To 2
                 If i = 2 Then frmPTADScreen2!txtDesignConcentrationValue(i).Text = "0.0"
                 frmPTADScreen2!txtDesignConcentrationValue(i).Enabled = False
           End Select
       Next i

       If Scr2.KLaSafetyFactor.UserInput = False Then
          Scr2.KLaSafetyFactor.UserInput = True
          Scr2.KLaSafetyFactor.value = 1#
          frmPTADScreen2!txtDesignConcentrationValue(1).Text = "1.0"
          frmPTADScreen2!txtDesignConcentrationValue(2).Text = "0.0"
       End If
    Else
      
       frmPTADScreen2!mnuFile(4).Enabled = True
       frmPTADScreen2!mnuFile(5).Enabled = True
      
       frmPTADScreen2!mnuOptions(0).Enabled = True

       'frmPTADScreen2!fraDesignContaminant.Enabled = True
       frmPTADScreen2!cboSelectCompo.Enabled = True
       For i = 0 To 7
           Select Case i
              Case 0, 3 To 7
                 frmPTADScreen2!lblDesignConcentrationValue(i).Enabled = True
              Case 1 To 2
                 frmPTADScreen2!txtDesignConcentrationValue(i).Enabled = True
           End Select
       Next i
    End If
    Call frmPTADScreen2.LOCAL___Reset_DemoVersionDisablings
End Sub

Sub SetPowerPTADScreen2(CalculatedPower As Integer)

          Scr2.Power.InletAirTemperature = Scr2.operatingtemperature.value - 273.15
          Call CalculatePowerScreen2(CalculatedPower)
          If CalculatedPower Then
             frmPowerScreen2!txtPower(0).Text = Format$(Scr2.Power.InletAirTemperature, GetTheFormat(Scr2.Power.InletAirTemperature))
             frmPowerScreen2!txtPower(1).Text = Format$(Scr2.Power.BlowerEfficiency, GetTheFormat(Scr2.Power.BlowerEfficiency))
             frmPowerScreen2!lblPower(2).Caption = Format$(Scr2.Power.BlowerBrakePower, GetTheFormat(Scr2.Power.BlowerBrakePower))
             frmPowerScreen2!txtPower(3).Text = Format$(Scr2.Power.PumpEfficiency, GetTheFormat(Scr2.Power.PumpEfficiency))
             frmPowerScreen2!lblPower(4).Caption = Format$(Scr2.Power.PumpBrakePower, GetTheFormat(Scr2.Power.PumpBrakePower))
             frmPowerScreen2!lblPower(5).Caption = Format$(Scr2.Power.TotalBrakePower, GetTheFormat(Scr2.Power.TotalBrakePower))
          End If

End Sub

Sub SetUpFlowsLoadingsTextBoxes(UsersFlowAndLoadingOption As Integer)
Dim i As Integer

    Select Case UsersFlowAndLoadingOption

       Case 0  'Specify water flow rate and air flow rate
            Scr2.WaterFlowRate.UserInput = True
            Scr2.AirFlowRate.UserInput = True
            Scr2.AirToWaterRatio.UserInput = False
            Scr2.WaterLoadingRate.UserInput = False
            Scr2.AirLoadingRate.UserInput = False
            For i = 0 To 1
                frmPTADScreen2!txtFlowsLoadings(i).Enabled = True
                frmPTADScreen2!txtFlowsLoadings(i).TabStop = True
            Next i
            For i = 2 To 4
                frmPTADScreen2!txtFlowsLoadings(i).Enabled = False
                frmPTADScreen2!txtFlowsLoadings(i).TabStop = False
            Next i

       Case 1  'Specify water flow rate and air to water ratio
            Scr2.WaterFlowRate.UserInput = True
            Scr2.AirFlowRate.UserInput = False
            Scr2.AirToWaterRatio.UserInput = True
            Scr2.WaterLoadingRate.UserInput = False
            Scr2.AirLoadingRate.UserInput = False
            For i = 0 To 4
                Select Case i
                   Case 0, 2
                        frmPTADScreen2!txtFlowsLoadings(i).Enabled = True
                        frmPTADScreen2!txtFlowsLoadings(i).TabStop = True
                   Case 1, 3, 4
                        frmPTADScreen2!txtFlowsLoadings(i).Enabled = False
                        frmPTADScreen2!txtFlowsLoadings(i).TabStop = False
                 End Select
            Next i
        
       Case 2  'Specify water loading rate and air loading rate
            Scr2.WaterFlowRate.UserInput = False
            Scr2.AirFlowRate.UserInput = False
            Scr2.AirToWaterRatio.UserInput = False
            Scr2.WaterLoadingRate.UserInput = True
            Scr2.AirLoadingRate.UserInput = True
            For i = 0 To 4
                Select Case i
                   Case 3, 4
                        frmPTADScreen2!txtFlowsLoadings(i).Enabled = True
                        frmPTADScreen2!txtFlowsLoadings(i).TabStop = True
                   Case 0 To 2
                        frmPTADScreen2!txtFlowsLoadings(i).Enabled = False
                        frmPTADScreen2!txtFlowsLoadings(i).TabStop = False
                 End Select
            Next i

    End Select

End Sub

Sub ShowOndaKLaPropertiesScreen2()
    
       frmShowOndaKLaProperties!lblOndaProperties(0).Caption = Format$(Scr2.Onda.ReynoldsNumber, GetTheFormat(Scr2.Onda.ReynoldsNumber))
       frmShowOndaKLaProperties!lblOndaProperties(1).Caption = Format$(Scr2.Onda.FroudeNumber, GetTheFormat(Scr2.Onda.FroudeNumber))
       frmShowOndaKLaProperties!lblOndaProperties(2).Caption = Format$(Scr2.Onda.WeberNumber, GetTheFormat(Scr2.Onda.WeberNumber))
       frmShowOndaKLaProperties!lblOndaProperties(3).Caption = Format$(Scr2.Packing.OndaWettedSurfaceArea, GetTheFormat(Scr2.Packing.OndaWettedSurfaceArea))
       frmShowOndaKLaProperties!lblOndaProperties(4).Caption = Format$(Scr2.Onda.LiquidPhaseMassTransferResistance, GetTheFormat(Scr2.Onda.LiquidPhaseMassTransferResistance))
       frmShowOndaKLaProperties!lblOndaProperties(5).Caption = Format$(Scr2.Onda.GasPhaseMassTransferResistance, GetTheFormat(Scr2.Onda.GasPhaseMassTransferResistance))
       frmShowOndaKLaProperties!lblOndaProperties(6).Caption = Format$(Scr2.Onda.TotalMassTransferResistance, GetTheFormat(Scr2.Onda.TotalMassTransferResistance))
       frmShowOndaKLaProperties!lblOndaProperties(7).Caption = Format$(Scr2.Onda.LiquidPhaseMassTransferCoefficient, GetTheFormat(Scr2.Onda.LiquidPhaseMassTransferCoefficient))
       frmShowOndaKLaProperties!lblOndaProperties(8).Caption = Format$(Scr2.Onda.GasPhaseMassTransferCoefficient, GetTheFormat(Scr2.Onda.GasPhaseMassTransferCoefficient))
       frmShowOndaKLaProperties!lblOndaProperties(9).Caption = Format$(Scr2.Onda.OverallMassTransferCoefficient, GetTheFormat(Scr2.Onda.OverallMassTransferCoefficient))

End Sub

Sub SpecifiedDesignKLaScreen2()

    If HaveValue(Scr2.Onda.OverallMassTransferCoefficient) Then
       Call GETSAF(Scr2.KLaSafetyFactor.value, Scr2.Onda.OverallMassTransferCoefficient, Scr2.DesignMassTransferCoefficient.value)
       Scr2.KLaSafetyFactor.ValChanged = True
       frmPTADScreen2!txtDesignConcentrationValue(1).Text = Format$(Scr2.KLaSafetyFactor.value, GetTheFormat(Scr2.KLaSafetyFactor.value))
    ElseIf Scr2.KLaSafetyFactor.value > 0# Then
       Call KLaOverSpecificationMessage
       Scr2.KLaSafetyFactor.value = 0#
       frmPTADScreen2!txtDesignConcentrationValue(1).Text = "0.0"
    End If
    Scr2.KLaSafetyFactor.UserInput = False

End Sub

Sub SpecifiedKLaSafetyFactorScreen2()
Dim Dummy As Double

  If HaveValue(Scr2.Onda.OverallMassTransferCoefficient) Then
    Call KLACOR(Scr2.DesignMassTransferCoefficient.value, Scr2.Onda.OverallMassTransferCoefficient, Scr2.KLaSafetyFactor.value)
    Scr2.DesignMassTransferCoefficient.ValChanged = True
    'frmPTADScreen2!txtDesignConcentrationValue(2).Text = Format$(Scr2.DesignMassTransferCoefficient.Value, GetTheFormat(Scr2.DesignMassTransferCoefficient.Value))
    
    'Update Contaminant of Interest | Design KLa.
    Dummy = Scr2.DesignMassTransferCoefficient.value
    Call Unitted_UnitChange(UNITS_INVERSETIME, Dummy, frmPTADScreen2!UnitsInterest(2), frmPTADScreen2!txtDesignConcentrationValue(2))
  ElseIf Scr2.DesignMassTransferCoefficient.value > 0# Then
    Call KLaOverSpecificationMessage
    Scr2.DesignMassTransferCoefficient.value = 0#
    frmPTADScreen2!txtDesignConcentrationValue(2).Text = "0.0"
  End If
  Scr2.DesignMassTransferCoefficient.UserInput = False

End Sub

Function StartScreen2DefaultCase() As Boolean

    Filename = "TheDefaultCaseScreen2"
    StartScreen2DefaultCase = loadscreen2("")

End Function

Sub CalculateAirWaterPropertiesScreen2()
    Dim Pressure As Double
    Dim Temperature As Double
    Dim WaterDensity As Double
    Dim WaterViscosity As Double
    Dim WaterSurfaceTension As Double
    Dim AirDensity As Double
    Dim AirViscosity As Double
    Dim i As Integer
    
    If Scr2.OperatingPressure.ValChanged Or Scr2.operatingtemperature.ValChanged Then
       Pressure = Scr2.OperatingPressure.value
       Temperature = Scr2.operatingtemperature.value

       For i = 0 To 4
           If frmAirWaterProperties.chkUpdateValues(i).value = True Then
              Select Case i
                 Case 0
                    If HaveValue(Temperature) Then
                       Call H2ODENS(WaterDensity, Temperature)
                       Scr2.WaterDensity.value = WaterDensity
                       Scr2.WaterDensity.UserInput = False
                       Scr2.WaterDensity.ValChanged = True
                       frmAirWaterProperties.txtAirWaterProperties(0).Text = Format$(WaterDensity, "0.00")
                       frmAirWaterProperties.lblValueSource(0).Caption = "Correlation"
                    End If
                 Case 1
                    If HaveValue(Temperature) Then
                       Call H2OVISC(WaterViscosity, Temperature)
                       Scr2.WaterViscosity.value = WaterViscosity
                       Scr2.WaterViscosity.UserInput = False
                       Scr2.WaterViscosity.ValChanged = True
                       frmAirWaterProperties.txtAirWaterProperties(1).Text = Format$(WaterViscosity, GetTheFormat(WaterViscosity))
                       frmAirWaterProperties.lblValueSource(1).Caption = "Correlation"
                    End If
                 Case 2
                    If HaveValue(Temperature) Then
                       Call H2OST(WaterSurfaceTension, Temperature)
                       Scr2.WaterSurfaceTension.value = WaterSurfaceTension
                       Scr2.WaterSurfaceTension.UserInput = False
                       Scr2.WaterSurfaceTension.ValChanged = True
                       frmAirWaterProperties.txtAirWaterProperties(2).Text = Format$(WaterSurfaceTension, GetTheFormat(WaterSurfaceTension))
                       frmAirWaterProperties.lblValueSource(2).Caption = "Correlation"
                    End If
                 Case 3
                    If HaveValue(Temperature) And HaveValue(Pressure) Then
                       Call AIRDENS(AirDensity, Temperature, Pressure)
                       Scr2.AirDensity.value = AirDensity
                       Scr2.AirDensity.UserInput = False
                       Scr2.AirDensity.ValChanged = True
                       frmAirWaterProperties.txtAirWaterProperties(3).Text = Format$(AirDensity, GetTheFormat(AirDensity))
                       frmAirWaterProperties.lblValueSource(3).Caption = "Correlation"
                    End If
                 Case 4
                    If HaveValue(Temperature) Then
                       Call AIRVISC(AirViscosity, Temperature)
                       Scr2.AirViscosity.value = AirViscosity
                       Scr2.AirViscosity.UserInput = False
                       Scr2.AirViscosity.ValChanged = True
                       frmAirWaterProperties.txtAirWaterProperties(4).Text = Format$(AirViscosity, GetTheFormat(AirViscosity))
                       frmAirWaterProperties.lblValueSource(4).Caption = "Correlation"
                    End If
              End Select
          End If
       Next i
    End If

End Sub

Sub xOLDxGetContaminantConcentrationsScreen2()
    Dim PercentRemoval As Double
    Dim msg As String, Response As Integer
    Dim Answer As String
    Dim NewStep As Double

    Call GetOndaMassTransferCoefficientScreen2
    Call GetDesignKLaOrKLaSafetyFactorScreen2

    frmPTADScreen2!lblDesignConcentrationValue(3).Caption = Format$(Scr2.DesignContaminant.Influent.value, GetTheFormat(Scr2.DesignContaminant.Influent.value))
    frmPTADScreen2!lblDesignConcentrationValue(4).Caption = Format$(Scr2.DesignContaminant.TreatmentObjective.value, GetTheFormat(Scr2.DesignContaminant.TreatmentObjective.value))

    Call EFFLPT2(Scr2.DesignContaminant.Effluent.value, Scr2.AirToWaterRatio.value, Scr2.DesignContaminant.HenrysConstant.value, Scr2.WaterFlowRate.value, Scr2.TowerArea.value, Scr2.SpecifiedTowerHeight.value, Scr2.DesignMassTransferCoefficient.value, Scr2.DesignContaminant.Influent.value)
    frmPTADScreen2!lblDesignConcentrationValue(5).Caption = Format$(Scr2.DesignContaminant.Effluent.value, GetTheFormat(Scr2.DesignContaminant.Effluent.value))

    Call REMOVPT(PercentRemoval, Scr2.DesignContaminant.Influent.value, Scr2.DesignContaminant.Effluent.value)
    frmPTADScreen2!lblDesignConcentrationValue(6).Caption = Format$(PercentRemoval, "0.0")

    'Determine Pressure Drop
    Scr2.AirPressureDrop.value = -1#
    InitialPressureDrop = 1#
    FinalPressureDrop = 1200#
    PressureDropStep = 1#

PressureDrop:
    Call PDROP(Scr2.AirPressureDrop.value, Scr2.AirToWaterRatio.value, Scr2.AirLoadingRate.value, Scr2.Packing.PackingFactor, Scr2.WaterViscosity.value, Scr2.AirDensity.value, Scr2.WaterDensity.value, InitialPressureDrop, FinalPressureDrop, PressureDropStep)
    If Scr2.AirPressureDrop.value < 0 Then
       msg = "Failure to get within one percent of the "
       msg = msg + "y-axis value on the Eckert curve "
       msg = msg + "in the pressure drop range of "
       msg = msg + Format$(InitialPressureDrop, "0.0") + " N/m2/m and " + Format$(FinalPressureDrop, "0.0")
       msg = msg + " N/m2/m using a pressure drop step of " + Format$(PressureDropStep, "0.0000") + " N/m2/m."
       msg = msg & Chr$(13) & Chr$(13)
       msg = msg + "Would you like to specify a smaller value for pressure drop step "
       msg = msg + "and attempt to achieve convergence again?"
       Response = MsgBox(msg, MB_ICONquestion + MB_YESNO, "Pressure Drop Convergence Error")
       If Response = IDYES Then
          If PressureDropStep <= 0.01 Then
             msg = "You can not specify a pressure drop "
             msg = msg + "step smaller than 0.01. "
             msg = msg + "Convergence not possible in this "
             msg = msg + "case."
             MsgBox msg, MB_ICONEXCLAMATION, "Pressure Drop Convergence Error"
             frmPTADScreen2!lblDesignConcentrationValue(7).Caption = "N/A"
          Else
NewPressureDrop:
             If PressureDropStep / 10 < 0.01 Then
                Answer$ = InputBox$("Enter new value for pressure drop step.", "Pressure Drop Step", Format$(0.01, "0.000"))
             Else
                Answer$ = InputBox$("Enter new value for pressure drop step.", "Pressure Drop Step", Format$(PressureDropStep / 10, "0.000"))
             End If
             On Error GoTo NewPressureDrop:
                NewStep = CDbl(Answer$)
                If NewStep < 0.01 Then
                   MsgBox "Pressure Drop step must exceed 0.01", MB_ICONEXCLAMATION, "Error"
                   GoTo NewPressureDrop
                Else
                   PressureDropStep = NewStep
                   GoTo PressureDrop:
                End If
          End If
       Else
          frmPTADScreen2!lblDesignConcentrationValue(7).Caption = "N/A"
       End If
    Else
       frmPTADScreen2!lblDesignConcentrationValue(7).Caption = Format$(Scr2.AirPressureDrop.value, "0.0")
    End If

End Sub

