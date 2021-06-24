Attribute VB_Name = "BubbleMod"
Option Explicit

Global Const KLA_METHOD_CWO2_TRANSFER_TEST = 1
Global Const KLA_METHOD_USER_INPUT = 2
Global Const BUBBLE_FILEID = "Properties_Bubble_Aeration"
Global Const CONTAMINANTS_BUBBLE_FILEID = "Contaminants_Bubble_Aeration"
Global Const MAXIMUM_TANKS = 15   'Maximum number of tanks for design
Global Const DESIGN_MODE = 1
Global Const RATING_MODE = 2

Global BubbleAerationMode As Integer

Type BubbleInformationType
     value As Double
     UserInput As Integer
     ValChanged As Integer
End Type

Type BubbleInformationType2
     value As Long
     UserInput As Integer
     ValChanged As Integer
End Type

Type CleanWaterOxygenTransferTestDataType
     SOTR As BubbleInformationType
     SOTE As BubbleInformationType
     AirFlowRate_QAIR As BubbleInformationType
     BarometricPressure_PB As BubbleInformationType
     WaterDepth_DEPTHW As BubbleInformationType
     WaterVolumePerTank_VM3 As BubbleInformationType
     DOSaturationConc_CSTR20 As Double
     WeightDensityOfWater_GAMMAW As Double
     EffectiveSaturationDepth_DEFF As Double
     ApparentOxygenMTCoeff_KLA20 As Double
     WaterVolumePerTankLiters_V As Double
     TrueKLaAt20DegC_KLAT20 As Double
     Phi As Double
     TrueOxygenMTCoeffOperatingT_KLAO2 As Double
End Type

Type OxygenInformationType
     LiquidDiffusivity As BubbleInformationType
     MassTransferCoefficient As BubbleInformationType
     KLaMethod As Integer
     CWO2TestData As CleanWaterOxygenTransferTestDataType
End Type

Type BubbleContaminantPropertyType
     Name As String
     Pressure As Double
     Temperature As Double
     Effluent(0 To MAXIMUM_TANKS) As Double
     GasEffluent(1 To MAXIMUM_TANKS) As Double
     MolecularWeight As BubbleInformationType
     HenrysConstant As BubbleInformationType
     MolarVolume As BubbleInformationType
'     NormalBoilingPoint As BubbleInformationType
     LiquidDiffusivity As BubbleInformationType
'     GasDiffusivity As BubbleInformationType
     Influent As BubbleInformationType
     TreatmentObjective As BubbleInformationType
End Type

Type PowerTypeBubble
     BlowerBrakePower As Double
     TotalBrakePower As Double
     InletAirTemperature As Double
     BlowerEfficiency As Double
     TankWaterDepth As Double
     NumberOfBlowersinEachTank As Long
End Type

Type BubbleType
     OperatingPressure As BubbleInformationType
     operatingtemperature As BubbleInformationType
     WaterDensity As BubbleInformationType
     WaterViscosity As BubbleInformationType
     N_for_Finding_KLa As BubbleInformationType
     kgOVERkl_for_Finding_KLa As BubbleInformationType
     ContaminantMassTransferCoefficient As BubbleInformationType
     WaterFlowRate As BubbleInformationType
     MinimumAirToWaterRatio As BubbleInformationType
     AirToWaterRatio As BubbleInformationType
     AirFlowRate As BubbleInformationType
     TankHydraulicRetentionTime As BubbleInformationType
     TotalHydraulicRetentionTime As BubbleInformationType
     TankVolume As BubbleInformationType
     TotalTankVolume As BubbleInformationType
     StantonNumber As BubbleInformationType
     
     NumberOfTanks As BubbleInformationType2
     
     CodeForTausAndTankVolumes As Long
     DesiredPercentRemoval As Double
     AchievedPercentRemoval As Double
     ID_OptimalDesignContaminant As Integer
     
     Power As PowerTypeBubble
     
     Chemical As Integer
     NumChemical As Integer
     Contaminant(1 To MAXCHEMICAL) As BubbleContaminantPropertyType
     DesignContaminant As BubbleContaminantPropertyType
     
     Oxygen As OxygenInformationType
End Type

Global bub As BubbleType

Global ErrorFlagBub As Long   'Error Flag passed to Sub VOLBUB


'Constants for Printing

Global Const VALUE_TAB = 60
Global Const LIQUID_EFFLUENT_TAB = 45
Global Const GAS_EFFLUENT_TAB = 65
Global Const MWT_TAB = 10
Global Const HC_TAB = 21
Global Const VB_TAB = 32
Global Const DIFL_TAB = 43
Global Const MTCOEFF_TAB = 54
Global Const STANTON_TAB = 65

Sub CalculateAirToWaterRatio()

  Call VQBUB(bub.AirToWaterRatio.value, bub.AirFlowRate.value, bub.WaterFlowRate.value)
  frmBubble!txtFlowParameters(2).Text = Format$(bub.AirToWaterRatio.value, GetTheFormat(bub.AirToWaterRatio.value))
  bub.AirToWaterRatio.UserInput = False

End Sub

Sub CalculateApparentKLa()

  Call KLA20A(bub.Oxygen.CWO2TestData.ApparentOxygenMTCoeff_KLA20, bub.Oxygen.CWO2TestData.WaterVolumePerTankLiters_V, bub.Oxygen.CWO2TestData.WaterVolumePerTank_VM3.value, bub.Oxygen.CWO2TestData.DOSaturationConc_CSTR20, bub.Oxygen.CWO2TestData.SOTR.value)
  frmOxygenMassTransferCoeff!lblDataParameters(7).Caption = Format$(bub.Oxygen.CWO2TestData.ApparentOxygenMTCoeff_KLA20, GetTheFormat(bub.Oxygen.CWO2TestData.ApparentOxygenMTCoeff_KLA20))

End Sub

Sub CalculateContaminantMTCoeff()

  Call KLABUB(bub.ContaminantMassTransferCoefficient.value, bub.Oxygen.MassTransferCoefficient.value, bub.DesignContaminant.LiquidDiffusivity.value, bub.Oxygen.LiquidDiffusivity.value, bub.N_for_Finding_KLa.value, bub.kgOVERkl_for_Finding_KLa.value, bub.DesignContaminant.HenrysConstant.value)
   
  'UPDATED_UNITS.
  'frmBubble!txtDesignConcentrationValue(3).Text = Format$(bub.ContaminantMassTransferCoefficient.Value, GetTheFormat(bub.ContaminantMassTransferCoefficient.Value))
  Call Unitted_NumberUpdate(frmBubble!UnitsDesignContam(3))
   
  bub.ContaminantMassTransferCoefficient.UserInput = False

End Sub

Sub CalculateDOSaturationConc()

  Call GETCSTAR(bub.Oxygen.CWO2TestData.DOSaturationConc_CSTR20, bub.Oxygen.CWO2TestData.WeightDensityOfWater_GAMMAW, bub.Oxygen.CWO2TestData.EffectiveSaturationDepth_DEFF, bub.Oxygen.CWO2TestData.BarometricPressure_PB.value, bub.Oxygen.CWO2TestData.WaterDepth_DEPTHW.value)
  frmOxygenMassTransferCoeff!lblDataParameters(6).Caption = Format$(bub.Oxygen.CWO2TestData.DOSaturationConc_CSTR20, GetTheFormat(bub.Oxygen.CWO2TestData.DOSaturationConc_CSTR20))
   
End Sub

Sub CalculateEffluentConcentrationsBubble()
ReDim Effluent(0 To MAXIMUM_TANKS) As Double
ReDim GasEffluent(1 To MAXIMUM_TANKS) As Double
Dim i As Integer
Dim Dummy As Double

  Call EFFLBUB(Effluent(0), GasEffluent(1), bub.DesignContaminant.HenrysConstant.value, bub.DesignContaminant.Influent.value, bub.AirToWaterRatio.value, bub.NumberOfTanks.value, bub.StantonNumber.value)
  For i = 1 To bub.NumberOfTanks.value
    bub.DesignContaminant.Effluent(i) = Effluent(i)
    bub.DesignContaminant.GasEffluent(i) = GasEffluent(i)
  Next i

  'UPDATED_UNITS.
  'frmBubble!lblConcentrationResults(3).Caption = Format$(bub.DesignContaminant.Effluent(bub.NumberOfTanks.Value), GetTheFormat(bub.DesignContaminant.Effluent(bub.NumberOfTanks.Value)))
  Dummy = bub.DesignContaminant.Effluent(bub.NumberOfTanks.value)
  Call Unitted_NumberUpdate(frmBubble!UnitsConcResults(3))
  
  For i = 1 To bub.NumberOfTanks.value
    frmBubbleEffluentConcentrations!lblTankNumber(i).Visible = True
    frmBubbleEffluentConcentrations!lblLiquidEffluent(i).Visible = True
    frmBubbleEffluentConcentrations!lblGasEffluent(i).Visible = True
    frmBubbleEffluentConcentrations!lblTankNumber(i).Caption = Trim$(Str$(i))
    frmBubbleEffluentConcentrations.lblLiquidEffluent(i).Caption = Format$(bub.DesignContaminant.Effluent(i), GetTheFormat(bub.DesignContaminant.Effluent(i)))
    frmBubbleEffluentConcentrations.lblGasEffluent(i).Caption = Format$(bub.DesignContaminant.GasEffluent(i), GetTheFormat(bub.DesignContaminant.GasEffluent(i)))
  Next i

  For i = (bub.NumberOfTanks.value + 1) To MAXIMUM_TANKS
    frmBubbleEffluentConcentrations!lblTankNumber(i).Visible = False
    frmBubbleEffluentConcentrations!lblLiquidEffluent(i).Visible = False
    frmBubbleEffluentConcentrations!lblGasEffluent(i).Visible = False
  Next i

  i = bub.NumberOfTanks.value
  frmBubbleEffluentConcentrations!cmdOK.Top = frmBubbleEffluentConcentrations!lblTankNumber(i).Top + frmBubbleEffluentConcentrations!lblTankNumber(i).Height + 300
  frmBubbleEffluentConcentrations.Height = frmBubbleEffluentConcentrations!cmdOK.Top + frmBubbleEffluentConcentrations!cmdOK.Height + 500
  frmBubbleEffluentConcentrations!cmdOK.Left = frmBubbleEffluentConcentrations.Width / 2 - frmBubbleEffluentConcentrations!cmdOK.Width / 2
       
  Call CalculateAchievedPercentRemovalBubble

End Sub

Sub CalculateMinAirToWaterRatio()

  If BubbleAerationMode = DESIGN_MODE Then
    frmBubble!lblFlowParameters(1).Enabled = True
    Call VQMINBUB(bub.MinimumAirToWaterRatio.value, bub.DesignContaminant.Influent.value, bub.DesignContaminant.TreatmentObjective.value, bub.DesignContaminant.HenrysConstant.value, bub.NumberOfTanks.value)
    frmBubble!lblFlowParameters(1).Caption = Format$(bub.MinimumAirToWaterRatio.value, GetTheFormat(bub.MinimumAirToWaterRatio.value))
  Else
    frmBubble!lblFlowParameters(1).Caption = "N/A"
    frmBubble!lblFlowParameters(1).Enabled = False
  End If

End Sub

Sub CalculateOxygenLiquidDiffusivity()

  Call DIFO2(bub.Oxygen.LiquidDiffusivity.value, bub.operatingtemperature.value)
  'frmBubble!txtOxygen(1).Text = Format$(bub.Oxygen.LiquidDiffusivity.Value, GetTheFormat(bub.Oxygen.LiquidDiffusivity.Value))
  Call Unitted_NumberUpdate(frmBubble!UnitsOxygenRef(1))
  bub.Oxygen.LiquidDiffusivity.UserInput = False

End Sub

Sub CalculatePowerBubble()

  Call PCALCBUB(bub.Power.TotalBrakePower, bub.Power.BlowerBrakePower, bub.OperatingPressure.value, bub.Power.InletAirTemperature, bub.AirFlowRate.value, bub.Power.BlowerEfficiency, bub.WaterDensity.value, bub.Power.TankWaterDepth, bub.NumberOfTanks.value, bub.Power.NumberOfBlowersinEachTank)

End Sub

Sub CalculateRetentionTimesAndTankVolumes()
Dim Dummy As Double

  Call TAUSVOLS(bub.TotalHydraulicRetentionTime.value, bub.NumberOfTanks.value, bub.TankHydraulicRetentionTime.value, bub.TankVolume.value, bub.TotalTankVolume.value, bub.WaterFlowRate.value, bub.CodeForTausAndTankVolumes)

  Select Case bub.CodeForTausAndTankVolumes
    Case 1   'Input Fluid Residence Time of Each Tank
      frmBubble!txtTankParameters(2).Text = Format$(bub.TotalHydraulicRetentionTime.value, GetTheFormat(bub.TotalHydraulicRetentionTime.value))
      frmBubble!txtTankParameters(3).Text = Format$(bub.TankVolume.value, GetTheFormat(bub.TankVolume.value))
      frmBubble!txtTankParameters(4).Text = Format$(bub.TotalTankVolume.value, GetTheFormat(bub.TotalTankVolume.value))
      bub.TotalHydraulicRetentionTime.UserInput = False
      bub.TankVolume.UserInput = False
      bub.TotalTankVolume.UserInput = False
    Case 2   'Input Total Fluid Residence Time
      frmBubble!txtTankParameters(1).Text = Format$(bub.TankHydraulicRetentionTime.value, GetTheFormat(bub.TankHydraulicRetentionTime.value))
      frmBubble!txtTankParameters(3).Text = Format$(bub.TankVolume.value, GetTheFormat(bub.TankVolume.value))
      frmBubble!txtTankParameters(4).Text = Format$(bub.TotalTankVolume.value, GetTheFormat(bub.TotalTankVolume.value))
      bub.TankHydraulicRetentionTime.UserInput = False
      bub.TankVolume.UserInput = False
      bub.TotalTankVolume.UserInput = False
    Case 3   'Input Volume of Each Tank
      frmBubble!txtTankParameters(1).Text = Format$(bub.TankHydraulicRetentionTime.value, GetTheFormat(bub.TankHydraulicRetentionTime.value))
      frmBubble!txtTankParameters(2).Text = Format$(bub.TotalHydraulicRetentionTime.value, GetTheFormat(bub.TotalHydraulicRetentionTime.value))
      frmBubble!txtTankParameters(4).Text = Format$(bub.TotalTankVolume.value, GetTheFormat(bub.TotalTankVolume.value))
      bub.TankHydraulicRetentionTime.UserInput = False
      bub.TotalHydraulicRetentionTime.UserInput = False
      bub.TotalTankVolume.UserInput = False
    Case 4   'Input Total Volume of All Tanks
      frmBubble!txtTankParameters(1).Text = Format$(bub.TankHydraulicRetentionTime.value, GetTheFormat(bub.TankHydraulicRetentionTime.value))
      frmBubble!txtTankParameters(2).Text = Format$(bub.TotalHydraulicRetentionTime.value, GetTheFormat(bub.TotalHydraulicRetentionTime.value))
      frmBubble!txtTankParameters(3).Text = Format$(bub.TankVolume.value, GetTheFormat(bub.TankVolume.value))
      bub.TankHydraulicRetentionTime.UserInput = False
      bub.TotalHydraulicRetentionTime.UserInput = False
      bub.TankVolume.UserInput = False
  End Select

  Call Unitted_NumberUpdate(frmBubble!UnitsTankParam(1))
  Call Unitted_NumberUpdate(frmBubble!UnitsTankParam(2))
  Call Unitted_NumberUpdate(frmBubble!UnitsTankParam(3))
  Call Unitted_NumberUpdate(frmBubble!UnitsTankParam(4))

End Sub

Sub CalculateSOTE()

  Call GETSOTE(bub.Oxygen.CWO2TestData.SOTE.value, bub.Oxygen.CWO2TestData.SOTR.value, bub.Oxygen.CWO2TestData.AirFlowRate_QAIR.value)
  frmOxygenMassTransferCoeff!txtDataParameters(0).Text = Format$(bub.Oxygen.CWO2TestData.SOTE.value, GetTheFormat(bub.Oxygen.CWO2TestData.SOTE.value))
    
End Sub

Sub CalculateSOTR()

  Call GETSOTR(bub.Oxygen.CWO2TestData.SOTR.value, bub.Oxygen.CWO2TestData.SOTE.value, bub.Oxygen.CWO2TestData.AirFlowRate_QAIR.value)
  frmOxygenMassTransferCoeff!txtDataParameters(1).Text = Format$(bub.Oxygen.CWO2TestData.SOTR.value, GetTheFormat(bub.Oxygen.CWO2TestData.SOTR.value))

End Sub

Sub CalculateStantonNo()

  Call GETPHIB(bub.StantonNumber.value, bub.ContaminantMassTransferCoefficient.value, bub.TankVolume.value, bub.DesignContaminant.HenrysConstant.value, bub.AirFlowRate.value)
  frmBubble!lblStanton.Caption = Format$(bub.StantonNumber.value, GetTheFormat(bub.StantonNumber.value))
      
End Sub

Sub CalculateTankVolumeBubble()

CalculateVolume:
  Call VOLBUB(bub.TankVolume.value, bub.DesignContaminant.HenrysConstant.value, bub.AirFlowRate.value, bub.ContaminantMassTransferCoefficient.value, bub.DesignContaminant.Influent.value, bub.DesignContaminant.TreatmentObjective.value, bub.NumberOfTanks.value, bub.WaterFlowRate.value, ErrorFlagBub)
  If ErrorFlagBub <> -1 Then
    'frmBubble!txtTankParameters(3).Text = Format$(bub.TankVolume.Value, GetTheFormat(bub.TankVolume.Value))
    Call Unitted_NumberUpdate(frmBubble!UnitsTankParam(3))
    bub.TankVolume.UserInput = False
  Else
    If bub.AirToWaterRatio.value < bub.MinimumAirToWaterRatio.value Then
      frmBubbleAchievingRemovalEfficiency!lblAchieving(0).Caption = Format$(bub.MinimumAirToWaterRatio.value, GetTheFormat(bub.MinimumAirToWaterRatio.value))
      frmBubbleAchievingRemovalEfficiency!txtAchieving(1).Text = Format$(bub.AirToWaterRatio.value, GetTheFormat(bub.AirToWaterRatio.value))
      frmBubbleAchievingRemovalEfficiency!txtAchieving(2).Text = Format$(bub.NumberOfTanks.value, "0")
      frmBubbleAchievingRemovalEfficiency.Show 1
      GoTo CalculateVolume
    End If
  End If

End Sub

Sub CalculateTrueKLa()

  Call TrueKLa(bub.Oxygen.CWO2TestData.TrueOxygenMTCoeffOperatingT_KLAO2, bub.Oxygen.CWO2TestData.TrueKLaAt20DegC_KLAT20, bub.Oxygen.CWO2TestData.Phi, bub.Oxygen.CWO2TestData.AirFlowRate_QAIR.value, bub.Oxygen.CWO2TestData.WaterVolumePerTankLiters_V, bub.Oxygen.CWO2TestData.BarometricPressure_PB.value, bub.Oxygen.CWO2TestData.WeightDensityOfWater_GAMMAW, bub.Oxygen.CWO2TestData.ApparentOxygenMTCoeff_KLA20, bub.Oxygen.CWO2TestData.EffectiveSaturationDepth_DEFF, bub.operatingtemperature.value)
  frmOxygenMassTransferCoeff!lblDataParameters(8).Caption = Format$(bub.Oxygen.CWO2TestData.Phi, GetTheFormat(bub.Oxygen.CWO2TestData.Phi))
  frmOxygenMassTransferCoeff!lblDataParameters(9).Caption = Format$(bub.Oxygen.CWO2TestData.TrueKLaAt20DegC_KLAT20, GetTheFormat(bub.Oxygen.CWO2TestData.TrueKLaAt20DegC_KLAT20))
  frmOxygenMassTransferCoeff!lblDataParameters(10).Caption = "1.024"
  frmOxygenMassTransferCoeff!lblDataParameters(11).Caption = Format$(bub.Oxygen.CWO2TestData.TrueOxygenMTCoeffOperatingT_KLAO2, GetTheFormat(bub.Oxygen.CWO2TestData.TrueOxygenMTCoeffOperatingT_KLAO2))
   
End Sub

Sub CalculateWaterPropertiesBubble()
Dim Pressure As Double
Dim Temperature As Double
Dim WaterDensity As Double
Dim WaterViscosity As Double
Dim i As Integer
    
    
       Pressure = bub.OperatingPressure.value
       Temperature = bub.operatingtemperature.value

       For i = 0 To 1
           If frmWaterPropertiesBubble!chkUpdateValues(i).value = True Then
              Select Case i
                 Case 0
                    If HaveValue(Temperature) Then
                       Call H2ODENS(WaterDensity, Temperature)
                       bub.WaterDensity.value = WaterDensity
                       bub.WaterDensity.UserInput = False
                       bub.WaterDensity.ValChanged = True
                       frmWaterPropertiesBubble.txtAirWaterProperties(0).Text = Format$(WaterDensity, "###0.00")
                       frmWaterPropertiesBubble.lblValueSource(0).Caption = "Correlation"
                    End If
                 Case 1
                    If HaveValue(Temperature) Then
                       Call H2OVISC(WaterViscosity, Temperature)
                       bub.WaterViscosity.value = WaterViscosity
                       bub.WaterViscosity.UserInput = False
                       bub.WaterViscosity.ValChanged = True
                       frmWaterPropertiesBubble.txtAirWaterProperties(1).Text = Format$(WaterViscosity, "0.000E+##")
                       frmWaterPropertiesBubble.lblValueSource(1).Caption = "Correlation"
                    End If
              End Select
          End If
       Next i
    

End Sub

Sub InitializeCWO2TestData()

  bub.Oxygen.CWO2TestData.SOTR.value = 1469.6
  bub.Oxygen.CWO2TestData.SOTE.value = 12.974691
  bub.Oxygen.CWO2TestData.AirFlowRate_QAIR.value = 1699#
  bub.Oxygen.CWO2TestData.BarometricPressure_PB.value = 1#
  bub.Oxygen.CWO2TestData.WaterDepth_DEPTHW.value = 4#
  bub.Oxygen.CWO2TestData.WaterVolumePerTank_VM3.value = 485.63655

  frmOxygenMassTransferCoeff!txtDataParameters(1).Text = Trim$(Str$(bub.Oxygen.CWO2TestData.SOTR.value))
  frmOxygenMassTransferCoeff!txtDataParameters(0).Text = Trim$(Str$(bub.Oxygen.CWO2TestData.SOTE.value))
  frmOxygenMassTransferCoeff!txtDataParameters(2).Text = Trim$(Str$(bub.Oxygen.CWO2TestData.AirFlowRate_QAIR.value))
  frmOxygenMassTransferCoeff!txtDataParameters(3).Text = Trim$(Format$(bub.Oxygen.CWO2TestData.BarometricPressure_PB.value * 101325# / 1#))
  frmOxygenMassTransferCoeff!txtDataParameters(4).Text = Trim$(Str$(bub.Oxygen.CWO2TestData.WaterDepth_DEPTHW.value))
  frmOxygenMassTransferCoeff!txtDataParameters(5).Text = Trim$(Str$(bub.Oxygen.CWO2TestData.WaterVolumePerTank_VM3.value))

  frmOxygenMassTransferCoeff!optDataAvailable(0).value = True
  frmOxygenMassTransferCoeff!txtDataParameters(0).Enabled = False
  frmOxygenMassTransferCoeff!txtDataParameters(1).Enabled = True

  Call CalculateDOSaturationConc
  Call CalculateApparentKLa
  Call CalculateTrueKLa

End Sub

Sub InitializeOxygenMTCoeff()

  frmBubble!cboOxygen.ListIndex = 1   'User input
  bub.Oxygen.KLaMethod = KLA_METHOD_USER_INPUT
  bub.Oxygen.MassTransferCoefficient.value = 0.0046
  frmBubble!txtOxygen(2).Text = "0.0046"
  'frmBubble!UnitsOxygenRef(2).ListIndex = 0

End Sub

Sub InitializePressureTemperatureBubble()
    
  '*****************************************************
  '*                                                   *
  '* Initialize Pressure and Temperature to defaults:  *
  '*                                                   *
  '*  Operating Pressure = 1 atm                       *
  '*  Operating Temperature = 10.0 C                   *
  '*                                                   *
  '*****************************************************

  bub.OperatingPressure.value = 1#
  bub.OperatingPressure.ValChanged = True
  bub.operatingtemperature.value = 293.15
  bub.operatingtemperature.ValChanged = True

  frmBubble.txtOperatingPressure.Text = "101325.0"
  frmBubble.txtOperatingTemperature.Text = "20.00"

  Call CalculateWaterPropertiesBubble
  Call CalculateOxygenLiquidDiffusivity

End Sub

Function loadbubble(OverrideFilename As String) As Boolean
Dim FileID As String, msg As String
Dim i As Integer
Dim TransferTestDummy As Integer
Dim CommentDummy As String
Dim SelectedContaminant As Integer
ReDim u(10) As String
Dim xu As rec_Units_frmContaminantPropertyEdit

    If (OverrideFilename <> "") Then
      Filename = OverrideFilename
    Else
      If Filename = "TheDefaultCaseBubble" Then
        If BubbleAerationMode = DESIGN_MODE Then
          Filename = App.Path & "\dbase\defltdes.bub"
        Else
          Filename = App.Path & "\dbase\defltrat.bub"
        End If
      Else
        Call LoadFileBubble(Filename)
      End If
    End If
    
    If Filename$ <> "" Then
       FileID = ""
       If (fileexists(Filename) = False) Then
         Call Error_Unavailable_File( _
            Filename, _
            IIf(BubbleAerationMode = DESIGN_MODE, _
                "Bubble Aeration Design Mode", _
                "Bubble Aeration Rating Mode"))
         loadbubble = False
         Exit Function
       End If
       Open Filename$ For Input As #1
       On Error Resume Next
       Input #1, FileID
       If FileID <> BUBBLE_FILEID Then
          msg = "Invalid Design File"
          MsgBox msg, 48, "Error"
          Close #1
          Exit Function
       End If

       'frmListContaminantBubble.ListContaminants.Clear
       frmBubble!cboDesignContaminant.Clear

       Input #1, BubbleAerationMode, CommentDummy
       If BubbleAerationMode = DESIGN_MODE Then
          frmBubble.Caption = "Bubble Aeration - Design Mode"
          frmBubble!mnuFile(0).Caption = "Switch to &Rating Mode"
       ElseIf BubbleAerationMode = RATING_MODE Then
          frmBubble.Caption = "Bubble Aeration - Rating Mode"
          frmBubble!mnuFile(0).Caption = "Switch to &Design Mode"
       End If

       Input #1, bub.OperatingPressure.value, CommentDummy
       frmBubble!txtOperatingPressure.Text = Format$(bub.OperatingPressure.value * 101325# / 1#, "0.00")

       Input #1, bub.operatingtemperature.value, CommentDummy
       frmBubble!txtOperatingTemperature.Text = Format$(bub.operatingtemperature.value - 273.15, "0.0")

       Call CalculateWaterPropertiesBubble
       Call CalculateOxygenLiquidDiffusivity

       Input #1, bub.Oxygen.KLaMethod, CommentDummy
       If bub.Oxygen.KLaMethod = KLA_METHOD_USER_INPUT Then
          Input #1, bub.Oxygen.MassTransferCoefficient.value, CommentDummy
          frmBubble!cboOxygen.ListIndex = 1
          frmBubble!txtOxygen(2).Text = Trim$(Str$(bub.Oxygen.MassTransferCoefficient.value))
          Call CalculateTrueKLa  'Initialize Oxygen Mass Transfer Coefficient to Correct

       ElseIf bub.Oxygen.KLaMethod = KLA_METHOD_CWO2_TRANSFER_TEST Then
          Input #1, TransferTestDummy, CommentDummy
          If TransferTestDummy = 1 Then       'SOTR vs. QAIR data available
             Input #1, bub.Oxygen.CWO2TestData.SOTR.value, CommentDummy
             Input #1, bub.Oxygen.CWO2TestData.AirFlowRate_QAIR.value, CommentDummy
             Input #1, bub.Oxygen.CWO2TestData.BarometricPressure_PB.value, CommentDummy
             Input #1, bub.Oxygen.CWO2TestData.WaterDepth_DEPTHW.value, CommentDummy
             Input #1, bub.Oxygen.CWO2TestData.WaterVolumePerTank_VM3.value, CommentDummy

             frmOxygenMassTransferCoeff!txtDataParameters(1).Text = Trim$(Str$(bub.Oxygen.CWO2TestData.SOTR.value))
             frmOxygenMassTransferCoeff!txtDataParameters(2).Text = Trim$(Str$(bub.Oxygen.CWO2TestData.AirFlowRate_QAIR.value))
             frmOxygenMassTransferCoeff!txtDataParameters(3).Text = Trim$(Str$(bub.Oxygen.CWO2TestData.BarometricPressure_PB.value * 101325# / 1#))
             frmOxygenMassTransferCoeff!txtDataParameters(4).Text = Trim$(Str$(bub.Oxygen.CWO2TestData.WaterDepth_DEPTHW.value))
             frmOxygenMassTransferCoeff!txtDataParameters(5).Text = Trim$(Str$(bub.Oxygen.CWO2TestData.WaterVolumePerTank_VM3.value))

             frmOxygenMassTransferCoeff!optDataAvailable(0).value = True
             frmOxygenMassTransferCoeff!txtDataParameters(0).Enabled = False
             frmOxygenMassTransferCoeff!txtDataParameters(1).Enabled = True

             Call CalculateSOTE
             Call CalculateDOSaturationConc
             Call CalculateApparentKLa
             Call CalculateTrueKLa
             frmBubble!cboOxygen.ListIndex = 0
             frmOxygenMassTransferCoeff.Hide
             frmBubble!txtOxygen(2).Text = frmOxygenMassTransferCoeff!lblDataParameters(11).Caption
             bub.Oxygen.MassTransferCoefficient.value = bub.Oxygen.CWO2TestData.TrueOxygenMTCoeffOperatingT_KLAO2
             
          ElseIf TransferTestDummy = 2 Then   'SOTE vs. QAIR data available
   
             Input #1, bub.Oxygen.CWO2TestData.SOTE.value, CommentDummy
             Input #1, bub.Oxygen.CWO2TestData.AirFlowRate_QAIR.value, CommentDummy
             Input #1, bub.Oxygen.CWO2TestData.BarometricPressure_PB.value, CommentDummy
             Input #1, bub.Oxygen.CWO2TestData.WaterDepth_DEPTHW.value, CommentDummy
             Input #1, bub.Oxygen.CWO2TestData.WaterVolumePerTank_VM3.value, CommentDummy
                       
             frmOxygenMassTransferCoeff!txtDataParameters(0).Text = Trim$(Str$(bub.Oxygen.CWO2TestData.SOTE.value))
             frmOxygenMassTransferCoeff!txtDataParameters(2).Text = Trim$(Str$(bub.Oxygen.CWO2TestData.AirFlowRate_QAIR.value))
             frmOxygenMassTransferCoeff!txtDataParameters(3).Text = Trim$(Str$(bub.Oxygen.CWO2TestData.BarometricPressure_PB.value * 101325# / 1#))
             frmOxygenMassTransferCoeff!txtDataParameters(4).Text = Trim$(Str$(bub.Oxygen.CWO2TestData.WaterDepth_DEPTHW.value))
             frmOxygenMassTransferCoeff!txtDataParameters(5).Text = Trim$(Str$(bub.Oxygen.CWO2TestData.WaterVolumePerTank_VM3.value))

             frmOxygenMassTransferCoeff!optDataAvailable(1).value = True
             frmOxygenMassTransferCoeff!txtDataParameters(1).Enabled = False
             frmOxygenMassTransferCoeff!txtDataParameters(0).Enabled = True

             Call CalculateSOTR
             Call CalculateDOSaturationConc
             Call CalculateApparentKLa
             Call CalculateTrueKLa
             
             frmBubble!cboOxygen.ListIndex = 0
             frmOxygenMassTransferCoeff.Hide
             frmBubble!txtOxygen(2).Text = frmOxygenMassTransferCoeff!lblDataParameters(11).Caption
             bub.Oxygen.MassTransferCoefficient.value = bub.Oxygen.CWO2TestData.TrueOxygenMTCoeffOperatingT_KLAO2

          End If
       End If

       Input #1, bub.NumChemical, CommentDummy
       For i = 1 To bub.NumChemical
           Input #1, bub.Contaminant(i).Pressure, bub.Contaminant(i).Temperature, bub.Contaminant(i).Name, bub.Contaminant(i).MolecularWeight.value, bub.Contaminant(i).HenrysConstant.value, bub.Contaminant(i).MolarVolume.value, bub.Contaminant(i).LiquidDiffusivity.value, bub.Contaminant(i).Influent.value, bub.Contaminant(i).TreatmentObjective.value
           'frmListContaminantBubble.ListContaminants.AddItem bub.Contaminant(i).Name
           frmBubble!cboDesignContaminant.AddItem bub.Contaminant(i).Name
       Next i

       Input #1, bub.DesignContaminant.Name, CommentDummy

       Call SetDesignContaminantEnabledBubble(CInt(frmBubble!cboDesignContaminant.ListCount))

       For i = 1 To bub.NumChemical
           If bub.DesignContaminant.Name = bub.Contaminant(i).Name Then
              bub.DesignContaminant = bub.Contaminant(i)
              'frmListContaminantBubble!ListContaminants.Selected(i - 1) = True
              SelectedContaminant = i - 1
              Exit For
           End If
       Next i

       'If frmListContaminantBubble.mnuOptionsManipulateContaminant(1).Enabled = False Then
       '   frmListContaminantBubble.mnuOptionsManipulateContaminant(1).Enabled = True
       '   frmListContaminantBubble.mnuOptionsManipulateContaminant(3).Enabled = True
       '   frmListContaminantBubble.mnuOptionsManipulateContaminant(4).Enabled = True
       '   frmListContaminantBubble.mnuOptionsSave.Enabled = True
       '   frmListContaminantBubble.mnuOptionsView.Enabled = True
       '
       '   frmBubble!mnuFile(4).Enabled = True
       '   frmBubble!mnuFile(5).Enabled = True
       '   frmBubble!mnuOptions(0).Enabled = True
       'End If

       Call CalculateContaminantMTCoeff

       Input #1, bub.WaterFlowRate.value, CommentDummy
       frmBubble!txtFlowParameters(0).Text = Format$(bub.WaterFlowRate.value, GetTheFormat(bub.WaterFlowRate.value))

       Input #1, bub.AirToWaterRatio.value, bub.AirToWaterRatio.UserInput, CommentDummy
       frmBubble!txtFlowParameters(2).Text = Format$(bub.AirToWaterRatio.value, GetTheFormat(bub.AirToWaterRatio.value))

       Input #1, bub.AirFlowRate.value, bub.AirFlowRate.UserInput, CommentDummy
       frmBubble!txtFlowParameters(3).Text = Format$(bub.AirFlowRate.value, GetTheFormat(bub.AirFlowRate.value))

       If bub.AirToWaterRatio.UserInput = True Then
          Call CalculateAirFlowRate
       Else
          Call CalculateAirToWaterRatio
       End If

       Input #1, bub.NumberOfTanks.value, CommentDummy
       frmBubble!txtTankParameters(0).Text = Format$(bub.NumberOfTanks.value, "0")

       Input #1, bub.CodeForTausAndTankVolumes, CommentDummy

          Select Case bub.CodeForTausAndTankVolumes
             Case 1   'Input Hydraulic Retention Time for 1 Tank
                Input #1, bub.TankHydraulicRetentionTime.value, CommentDummy
                frmBubble!txtTankParameters(1).Text = Format$(bub.TankHydraulicRetentionTime.value, GetTheFormat(bub.TankHydraulicRetentionTime.value))
                bub.TankHydraulicRetentionTime.UserInput = True
             Case 2   'Input Hydraulic Retention Time for All Tanks
                Input #1, bub.TotalHydraulicRetentionTime.value, CommentDummy
                frmBubble!txtTankParameters(2).Text = Format$(bub.TotalHydraulicRetentionTime.value, GetTheFormat(bub.TotalHydraulicRetentionTime.value))
                bub.TotalHydraulicRetentionTime.UserInput = True
             Case 3   'Input Volume of Each Tank
                Input #1, bub.TankVolume.value, CommentDummy
                frmBubble!txtTankParameters(3).Text = Format$(bub.TankVolume.value, GetTheFormat(bub.TankVolume.value))
                bub.TankVolume.UserInput = True
             Case 4   'Input Volume of All Tanks
                Input #1, bub.TotalTankVolume.value, CommentDummy
                frmBubble!txtTankParameters(4).Text = Format$(bub.TotalTankVolume.value, GetTheFormat(bub.TotalTankVolume.value))
                bub.TotalTankVolume.UserInput = True
          End Select

       Input #1, bub.Power.BlowerEfficiency, CommentDummy
       Input #1, bub.Power.TankWaterDepth, CommentDummy
       Input #1, bub.Power.NumberOfBlowersinEachTank, CommentDummy
    

       If BubbleAerationMode = DESIGN_MODE Then
          bub.CodeForTausAndTankVolumes = 3
          Call CalculateTankVolumeBubble
          For i = 1 To 4
              frmBubble!txtTankParameters(i).Enabled = False
          Next i
       Else
          For i = 1 To 4
              frmBubble!txtTankParameters(i).Enabled = True
          Next i
       End If

       Call CalculateRetentionTimesAndTankVolumes
       frmBubble.cboDesignContaminant.ListIndex = SelectedContaminant

       'Input the units of this screen.
       Input #1, u(1), u(2)
       Call SetUnits(frmBubble!UnitsOpCond(0), u(1))
       Call SetUnits(frmBubble!UnitsOpCond(1), u(2))
      
       Input #1, u(1), u(2)
       Call SetUnits(frmBubble!UnitsOxygenRef(1), u(1))
       Call SetUnits(frmBubble!UnitsOxygenRef(2), u(2))
      
       Input #1, u(1), u(2), u(3)
       Call SetUnits(frmBubble!UnitsDesignContam(0), u(1))
       Call SetUnits(frmBubble!UnitsDesignContam(1), u(2))
       Call SetUnits(frmBubble!UnitsDesignContam(3), u(3))
      
       Input #1, u(1), u(2)
       Call SetUnits(frmBubble!UnitsFlowParam(0), u(1))
       Call SetUnits(frmBubble!UnitsFlowParam(3), u(2))
      
       Input #1, u(1), u(2), u(3), u(4)
       Call SetUnits(frmBubble!UnitsTankParam(1), u(1))
       Call SetUnits(frmBubble!UnitsTankParam(2), u(2))
       Call SetUnits(frmBubble!UnitsTankParam(3), u(3))
       Call SetUnits(frmBubble!UnitsTankParam(4), u(4))
      
       Input #1, u(1), u(2), u(3)
       Call SetUnits(frmBubble!UnitsConcResults(1), u(1))
       Call SetUnits(frmBubble!UnitsConcResults(2), u(2))
       Call SetUnits(frmBubble!UnitsConcResults(3), u(3))
      
       'Input the units of frmContaminantPropertyEdit.
       xu = Units_frmContaminantPropertyEdit
       Input #1, xu.UnitsProp(0), xu.UnitsProp(2), xu.UnitsProp(3), xu.UnitsProp(4), xu.UnitsProp(5)
       Input #1, xu.UnitsConc(0), xu.UnitsConc(1)
       Units_frmContaminantPropertyEdit = xu
      
       Close #1

       If Right$(Filename, 12) = "defltdes.bub" Or Right$(Filename, 12) = "defltrat.bub" Then
          frmBubble.Caption = frmBubble.Caption & " (" & "untitled.bub" & ")"
       Else
          frmBubble.Caption = frmBubble.Caption & " (" & Filename & ")"
       End If

       'Add this file to the last-few-files list.
       Call LastFewFiles_MoveFilenameToTop(Filename)
    
    End If

    loadbubble = True

End Function

Sub bubble_results()
    Dim i As Integer, j As Integer
    ReDim ContaminantMTCoeff(1 To MAXCHEMICAL) As Double
    ReDim StantonNumber(1 To MAXCHEMICAL) As Double
    ReDim DesiredPercentRemoval(1 To MAXCHEMICAL) As Double
    ReDim AchievedPercentRemoval(1 To MAXCHEMICAL) As Double
    ReDim Effluent(0 To MAXIMUM_TANKS) As Double
    ReDim GasEffluent(1 To MAXIMUM_TANKS) As Double
    Dim ContaminantGlossaryBottom As Integer, GlossaryBottom As Integer

          For i = 1 To bub.NumChemical
              If bub.DesignContaminant.Name = bub.Contaminant(i).Name Then
                 DesiredPercentRemoval(i) = bub.DesiredPercentRemoval
                 ContaminantMTCoeff(i) = bub.ContaminantMassTransferCoefficient.value
                 StantonNumber(i) = bub.StantonNumber.value
                 bub.Contaminant(i).Effluent(0) = bub.DesignContaminant.Effluent(0)
                 For j = 1 To bub.NumberOfTanks.value
                     bub.Contaminant(i).Effluent(j) = bub.DesignContaminant.Effluent(j)
                     bub.Contaminant(i).GasEffluent(j) = bub.DesignContaminant.GasEffluent(j)
                 Next j
                 AchievedPercentRemoval(i) = bub.AchievedPercentRemoval
              Else
                 Call REMOVBUB(DesiredPercentRemoval(i), bub.Contaminant(i).Influent.value, bub.Contaminant(i).TreatmentObjective.value)
                 Call KLABUB(ContaminantMTCoeff(i), bub.Oxygen.MassTransferCoefficient.value, bub.Contaminant(i).LiquidDiffusivity.value, bub.Oxygen.LiquidDiffusivity.value, bub.N_for_Finding_KLa.value, bub.kgOVERkl_for_Finding_KLa.value, bub.Contaminant(i).HenrysConstant.value)
                 Call GETPHIB(StantonNumber(i), ContaminantMTCoeff(i), bub.TankVolume.value, bub.Contaminant(i).HenrysConstant.value, bub.AirFlowRate.value)
                 Call EFFLBUB(Effluent(0), GasEffluent(1), bub.Contaminant(i).HenrysConstant.value, bub.Contaminant(i).Influent.value, bub.AirToWaterRatio.value, bub.NumberOfTanks.value, StantonNumber(i))
                 bub.Contaminant(i).Effluent(0) = Effluent(0)
                 For j = 1 To bub.NumberOfTanks.value
                     bub.Contaminant(i).Effluent(j) = Effluent(j)
                     bub.Contaminant(i).GasEffluent(j) = GasEffluent(j)
                 Next j
                 Call REMOVBUB(AchievedPercentRemoval(i), bub.Contaminant(i).Influent.value, bub.Contaminant(i).Effluent(bub.NumberOfTanks.value))
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

    For i = 1 To bub.NumChemical
        frmViewEffluentConcentrationsASAP!lblContaminantNumber(i - 1).Visible = True
        frmViewEffluentConcentrationsASAP!lblContaminantNumber(i + 10 - 1).Visible = True
        frmViewEffluentConcentrationsASAP!lblInfluentConcentration(i - 1).Visible = True
        frmViewEffluentConcentrationsASAP!lblTreatmentObjective(i - 1).Visible = True
        frmViewEffluentConcentrationsASAP!lblDesiredPercentRemoval(i - 1).Visible = True
        frmViewEffluentConcentrationsASAP!lblEffluentConcentration(i - 1).Visible = True
        frmViewEffluentConcentrationsASAP!lblAchievedPercentRemoval(i - 1).Visible = True
        frmViewEffluentConcentrationsASAP!lblContaminantName(i - 1).Visible = True

        frmViewEffluentConcentrationsASAP!lblInfluentConcentration(i - 1).Caption = Format$(bub.Contaminant(i).Influent.value, GetTheFormat(bub.Contaminant(i).Influent.value))
        frmViewEffluentConcentrationsASAP!lblTreatmentObjective(i - 1).Caption = Format$(bub.Contaminant(i).TreatmentObjective.value, GetTheFormat(bub.Contaminant(i).TreatmentObjective.value))
        frmViewEffluentConcentrationsASAP!lblDesiredPercentRemoval(i - 1).Caption = Format$(DesiredPercentRemoval(i), "0.0")
        frmViewEffluentConcentrationsASAP!lblEffluentConcentration(i - 1).Caption = Format$(bub.Contaminant(i).Effluent(bub.NumberOfTanks.value), GetTheFormat(bub.Contaminant(i).Effluent(bub.NumberOfTanks.value)))
        frmViewEffluentConcentrationsASAP!lblAchievedPercentRemoval(i - 1).Caption = Format$(AchievedPercentRemoval(i), "0.0")
        frmViewEffluentConcentrationsASAP!lblContaminantName(i - 1).Caption = Trim$(LCase$(bub.Contaminant(i).Name))

    Next i

    frmViewEffluentConcentrationsASAP!fraConcentrationResults.Height = frmViewEffluentConcentrationsASAP!lblContaminantNumber(bub.NumChemical - 1).Top + frmViewEffluentConcentrationsASAP!lblContaminantNumber(bub.NumChemical - 1).Height + 120
    frmViewEffluentConcentrationsASAP!fraContaminantGlossary.Height = frmViewEffluentConcentrationsASAP!lblContaminantNumber(bub.NumChemical + 10 - 1).Top + frmViewEffluentConcentrationsASAP!lblContaminantNumber(bub.NumChemical + 10 - 1).Height + 120
    frmViewEffluentConcentrationsASAP!fraContaminantGlossary.Top = frmViewEffluentConcentrationsASAP!fraConcentrationResults.Top + frmViewEffluentConcentrationsASAP!fraConcentrationResults.Height + 120
    frmViewEffluentConcentrationsASAP!fraGlossary.Top = frmViewEffluentConcentrationsASAP!fraContaminantGlossary.Top
    ContaminantGlossaryBottom = frmViewEffluentConcentrationsASAP!fraContaminantGlossary.Top + frmViewEffluentConcentrationsASAP!fraContaminantGlossary.Height
    GlossaryBottom = frmViewEffluentConcentrationsASAP!fraGlossary.Top + frmViewEffluentConcentrationsASAP!fraGlossary.Height
    If GlossaryBottom > ContaminantGlossaryBottom Then
       frmViewEffluentConcentrationsASAP!cmdOK.Top = GlossaryBottom + 360
    Else
       frmViewEffluentConcentrationsASAP!cmdOK.Top = ContaminantGlossaryBottom + 360
    End If
    frmViewEffluentConcentrationsASAP.Height = frmViewEffluentConcentrationsASAP!cmdOK.Top + frmViewEffluentConcentrationsASAP!cmdOK.Height + 500 '420

    frmViewEffluentConcentrationsASAP.Show 1

End Sub

   Sub bubble_save()

End Sub

Function bubble_savechanges() As Integer
Dim i As Integer
Dim msg As String, Response As Integer

  msg = "Would you like to save the parameters "
  msg = msg + "for this bubble aeration design case to a file?" & Chr$(13) & Chr$(13)
  msg = msg + "Note:  Any information not saved will be permanently lost."
  Response = MsgBox(msg, MB_ICONquestion + MB_YESNOCANCEL, "Save Current Design")
                
  If Response = IDCANCEL Then
   Screen.MousePointer = 0
   bubble_savechanges = 1
   Exit Function
  End If

  If Response = IDYES Then
   Call savebubble
    If StrComp(Filename, "") = 0 Then Response = 5
                      
     Do While Response = 5
     msg = "Would you like to save the parameters "
     msg = msg + "for this bubble aeration design case to a file?" & Chr$(13) & Chr$(13)
     msg = msg + "Note:  Any information not saved will be permanently lost."
     Response = MsgBox(msg, MB_ICONquestion + MB_YESNOCANCEL, "Save Current Design")
                         
     If Response = IDCANCEL Then
        Screen.MousePointer = 0
        bubble_savechanges = 1
        Exit Function
     End If
                           
     If Response = IDYES Then Call savebubble
     If StrComp(Filename, "") = 0 And Response <> IDNO Then Response = 5
                    
     Loop
                 
  End If

bubble_savechanges = 0
End Function

Sub CalculateAchievedPercentRemovalBubble()

  Call REMOVBUB(bub.AchievedPercentRemoval, bub.DesignContaminant.Influent.value, bub.DesignContaminant.Effluent(bub.NumberOfTanks.value))
  frmBubble!lblConcentrationResults(4).Caption = Format$(bub.AchievedPercentRemoval, GetTheFormat(bub.AchievedPercentRemoval))

End Sub

Sub CalculateAirFlowRate()

  Call AIRFLO(bub.AirFlowRate.value, bub.AirToWaterRatio.value, bub.WaterFlowRate.value)
  'frmBubble!txtFlowParameters(3).Text = Format$(bub.AirFlowRate.Value, GetTheFormat(bub.AirFlowRate.Value))
  Call Unitted_NumberUpdate(frmBubble!UnitsFlowParam(3))
  bub.AirFlowRate.UserInput = False

End Sub

Sub LoadContaminantListBubble()
    Dim FileID As String, msg As String
    Dim Pressure As Double, Temperature As Double
    Dim NormalBoilingPoint As Double, GasDiffusivity As Double
    Dim i As Integer
    Dim NotSpecifiedAtOperatingTemperature As Integer
    Dim NotSpecifiedAtOperatingPressure As Integer

    Call LoadFile(Filename)
    
    If Filename$ <> "" Then
       FileID = ""
       Open Filename$ For Input As #1
       On Error Resume Next
       Input #1, FileID
       If FileID <> CONTAMINANTS_BUBBLE_FILEID And FileID <> CONTAMINANTS_SURFACE_FILEID And FileID <> CONTAMINANTS_PTAD_FILEID Then
          msg = "Invalid Contaminant File"
          MsgBox msg, 48, "Error"
          Close #1
          Exit Sub
       End If

       'frmListContaminantBubble.ListContaminants.Clear
       frmBubble!cboDesignContaminant.Clear

       i = 0
       NotSpecifiedAtOperatingTemperature = False
       NotSpecifiedAtOperatingPressure = False
       Do Until EOF(1)
          i = i + 1
          If FileID = CONTAMINANTS_BUBBLE_FILEID Or FileID = CONTAMINANTS_SURFACE_FILEID Then
             Input #1, bub.Contaminant(i).Pressure, bub.Contaminant(i).Temperature, bub.Contaminant(i).Name, bub.Contaminant(i).MolecularWeight.value, bub.Contaminant(i).HenrysConstant.value, bub.Contaminant(i).MolarVolume.value, bub.Contaminant(i).LiquidDiffusivity.value, bub.Contaminant(i).Influent.value, bub.Contaminant(i).TreatmentObjective.value
          Else
             Input #1, bub.Contaminant(i).Pressure, bub.Contaminant(i).Temperature, bub.Contaminant(i).Name, bub.Contaminant(i).MolecularWeight.value, bub.Contaminant(i).HenrysConstant.value, bub.Contaminant(i).MolarVolume.value, NormalBoilingPoint, bub.Contaminant(i).LiquidDiffusivity.value, GasDiffusivity, bub.Contaminant(i).Influent.value, bub.Contaminant(i).TreatmentObjective.value
          End If
          'frmListContaminantBubble.ListContaminants.AddItem bub.Contaminant(i).Name
          frmBubble!cboDesignContaminant.AddItem bub.Contaminant(i).Name
          If Not NotSpecifiedAtOperatingTemperature Then
             If Abs(bub.Contaminant(i).Temperature - bub.operatingtemperature.value) > TOLERANCE Then
                NotSpecifiedAtOperatingTemperature = True
             End If
          End If
          If Not NotSpecifiedAtOperatingPressure Then
             If Abs(bub.Contaminant(i).Pressure - bub.OperatingPressure.value) > TOLERANCE Then
                NotSpecifiedAtOperatingPressure = True
             End If
          End If

       Loop
       bub.NumChemical = i
          
       Close #1

       'If frmListContaminantBubble.mnuOptionsManipulateContaminant(1).Enabled = False Then
       '   frmListContaminantBubble.mnuOptionsManipulateContaminant(1).Enabled = True
       '   frmListContaminantBubble.mnuOptionsManipulateContaminant(3).Enabled = True
       '   frmListContaminantBubble.mnuOptionsManipulateContaminant(4).Enabled = True
       '   frmListContaminantBubble.mnuOptionsSave.Enabled = True
       '   frmListContaminantBubble.mnuOptionsView.Enabled = True
       'End If

       'frmListContaminantBubble.ListContaminants.Selected(0) = True

       If NotSpecifiedAtOperatingPressure And NotSpecifiedAtOperatingTemperature Then
          MsgBox "For one or more contaminants, the temperature and pressure at which the contaminant properties are specified differs from the operating temperature and pressure.", MB_ICONINFORMATION, "Warning"
       ElseIf NotSpecifiedAtOperatingTemperature Then
          MsgBox "For one or more contaminants, the temperature at which the contaminant properties are specified differs from the operating temperature.", MB_ICONINFORMATION, "Warning"
       ElseIf NotSpecifiedAtOperatingPressure Then
          MsgBox "For one or more contaminants, the pressure at which the contaminant properties are specified differs from the operating pressure.", MB_ICONINFORMATION, "Warning"
       End If

    End If

End Sub

Sub LoadFileBubble(Filename As String)
Dim Ctl As Control
Set Ctl = frmBubble.CommonDialog1

    On Error Resume Next
    'frmBubble!CMDialog1.DefaultExt = "bub"
    'frmBubble!CMDialog1.Filter = "Bubble Aeration Files (*.bub)|*.bub"
    'frmBubble!CMDialog1.DialogTitle = "Load Bubble Aeration File"
    'frmBubble!CMDialog1.flags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST
    'frmBubble!CMDialog1.Action = 1
    'Filename$ = frmBubble!CMDialog1.Filename
    Ctl.DefaultExt = "bub"
    Ctl.Filter = "Bubble Aeration Files (*.bub)|*.bub"
    Ctl.DialogTitle = "Load Bubble Aeration File"
    Ctl.flags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST
    Ctl.Action = 1
    Filename$ = Ctl.Filename
    If Err = 32755 Then   'Cancel selected by user
       Filename$ = ""
    End If

End Sub

Sub NewPageBubble()

          Printer.NewPage
          Printer.FontSize = 12
          Printer.FontBold = True
          If BubbleAerationMode = DESIGN_MODE Then
             Printer.Print "Bubble Aeration - Design Mode (continued)"
          Else
             Printer.Print "Bubble Aeration - Rating Mode (continued)"
          End If
          Printer.Print
          Printer.Print
          Printer.FontSize = 10
          Printer.FontBold = False

End Sub

Sub PrintBubble()
    Dim i As Integer, j As Integer
    ReDim ContaminantMTCoeff(1 To MAXCHEMICAL) As Double
    ReDim StantonNumber(1 To MAXCHEMICAL) As Double
    ReDim DesiredPercentRemoval(1 To MAXCHEMICAL) As Double
    ReDim AchievedPercentRemoval(1 To MAXCHEMICAL) As Double
    ReDim Effluent(0 To MAXIMUM_TANKS) As Double
    ReDim GasEffluent(1 To MAXIMUM_TANKS) As Double

    On Error GoTo PrinterError

    Select Case BubbleAerationMode
       Case DESIGN_MODE

          Printer.ScaleLeft = -1440
          Printer.ScaleTop = -1440
          Printer.CurrentX = 0
          Printer.CurrentY = 0
          Printer.FontSize = 12
          Printer.FontBold = True
          Printer.Print "Bubble Aeration - Design Mode"
          Printer.Print
          Printer.Print
          Printer.FontUnderline = True
          Printer.FontSize = 10
          Printer.Print "Property:"; Tab(VALUE_TAB); "Value:"
          Printer.Print
          Printer.FontUnderline = False
          Printer.FontBold = False
          Printer.Print "Operating Pressure (" & frmBubble!UnitsOpCond(0) & ")"; Tab(VALUE_TAB); frmBubble!txtOperatingPressure.Text
          Printer.Print "Operating Temperature (" & frmBubble!UnitsOpCond(1) & ")"; Tab(VALUE_TAB); frmBubble!txtOperatingTemperature.Text
          Printer.Print frmWaterPropertiesBubble!lblAirWaterProperties(0).Caption; Tab(VALUE_TAB); frmWaterPropertiesBubble!txtAirWaterProperties(0).Text
          Printer.Print frmWaterPropertiesBubble!lblAirWaterProperties(1).Caption; Tab(VALUE_TAB); frmWaterPropertiesBubble!txtAirWaterProperties(1).Text
          Printer.Print
          Printer.Print "Oxygen " & frmBubble!lblOxygenLabel(1).Caption & " (" & frmBubble!UnitsOxygenRef(1) & ")"; Tab(VALUE_TAB); frmBubble!txtOxygen(1).Text
          Printer.Print "Method to Find Oxygen KLa"; Tab(VALUE_TAB); frmBubble!cboOxygen.Text
          If frmBubble!cboOxygen.ListIndex = 0 Then   'Method to Find Oxygen KLa is Clean Water Oxygen Transfer Test Data
             Printer.Print "Standardized Oxygen Transfer Efficiency, SOTE (%)"; Tab(VALUE_TAB); frmOxygenMassTransferCoeff!txtDataParameters(0).Text; ""
             Printer.Print "Standardized Oxygen Transfer Rate (kg O2/d)"; Tab(VALUE_TAB); frmOxygenMassTransferCoeff!txtDataParameters(1).Text
             Printer.Print "Air Flow Rate (standard m" & Chr$(179) & "/hr)"; Tab(VALUE_TAB); frmOxygenMassTransferCoeff!txtDataParameters(2).Text
             Printer.Print "Barometric Pressure (Pa)"; Tab(VALUE_TAB); frmOxygenMassTransferCoeff!txtDataParameters(3).Text
             Printer.Print "Tank Water Depth (m" & Chr$(179) & ")"; Tab(VALUE_TAB); frmOxygenMassTransferCoeff!txtDataParameters(4).Text
             Printer.Print "Tank Water Volume (m" & Chr$(179) & ")"; Tab(VALUE_TAB); frmOxygenMassTransferCoeff!txtDataParameters(5).Text
             Printer.Print "D.O. Saturation Concentration at Infinite Time (mg/L)"; Tab(VALUE_TAB); frmOxygenMassTransferCoeff!lblDataParameters(6).Caption; ""
             Printer.Print frmOxygenMassTransferCoeff!lblDataParametersLabel(7).Caption; Tab(VALUE_TAB); frmOxygenMassTransferCoeff!lblDataParameters(7).Caption
             Printer.Print frmOxygenMassTransferCoeff!lblDataParametersLabel(8).Caption; Tab(VALUE_TAB); frmOxygenMassTransferCoeff!lblDataParameters(8).Caption
             Printer.Print frmOxygenMassTransferCoeff!lblDataParametersLabel(9).Caption; Tab(VALUE_TAB); frmOxygenMassTransferCoeff!lblDataParameters(9).Caption
             Printer.Print frmOxygenMassTransferCoeff!lblDataParametersLabel(10).Caption; Tab(VALUE_TAB); frmOxygenMassTransferCoeff!lblDataParameters(10).Caption
             Printer.Print frmOxygenMassTransferCoeff!lblDataParametersLabel(11).Caption; Tab(VALUE_TAB); frmOxygenMassTransferCoeff!lblDataParameters(11).Caption
          End If
          Printer.Print "Oxygen " & frmBubble!lblOxygenLabel(2).Caption & " (" & frmBubble!UnitsOxygenRef(2) & ")"; Tab(VALUE_TAB); frmBubble!txtOxygen(2).Text
          Printer.Print
          Printer.Print "Design Contaminant:  "; frmBubble!cboDesignContaminant.Text
          Printer.Print "Molecular Weight (kg/kmol)"; Tab(VALUE_TAB); Format$(bub.DesignContaminant.MolecularWeight.value, "0.00")
          Printer.Print "Henry's Constant (-)"; Tab(VALUE_TAB); Format$(bub.DesignContaminant.HenrysConstant.value, GetTheFormat(bub.DesignContaminant.HenrysConstant.value))
          Printer.Print "Molar Volume (m" & Chr$(179) & "/kmol)"; Tab(VALUE_TAB); Format$(bub.DesignContaminant.MolarVolume.value, GetTheFormat(bub.DesignContaminant.MolarVolume.value))
          Printer.Print "Liquid Diffusivity (m" & Chr$(178) & "/sec)"; Tab(VALUE_TAB); Format$(bub.DesignContaminant.LiquidDiffusivity.value, GetTheFormat(bub.DesignContaminant.LiquidDiffusivity.value))
          Printer.Print frmBubble!lblDesignConcentration(0).Caption & " (" & frmBubble!UnitsDesignContam(0) & ")"; Tab(VALUE_TAB); frmBubble!lblDesignConcentrationValue(0).Caption
          Printer.Print frmBubble!lblDesignConcentration(1).Caption & " (" & frmBubble!UnitsDesignContam(1) & ")"; Tab(VALUE_TAB); frmBubble!lblDesignConcentrationValue(1).Caption
          Printer.Print frmBubble!lblDesignConcentration(2).Caption; Tab(VALUE_TAB); frmBubble!lblDesignConcentrationValue(2).Caption
          Printer.Print frmBubble!lblDesignConcentration(3).Caption & " (" & frmBubble!UnitsDesignContam(3) & ")"; Tab(VALUE_TAB); frmBubble!txtDesignConcentrationValue(3).Text
          Printer.Print
          Printer.Print frmBubble!lblFlowParametersLabel(0).Caption & " (" & frmBubble!UnitsFlowParam(0) & ")"; Tab(VALUE_TAB); frmBubble!txtFlowParameters(0).Text
          Printer.Print frmBubble!lblFlowParametersLabel(1).Caption; Tab(VALUE_TAB); frmBubble!lblFlowParameters(1).Caption
          Printer.Print frmBubble!lblFlowParametersLabel(2).Caption; Tab(VALUE_TAB); frmBubble!txtFlowParameters(2).Text
          Printer.Print frmBubble!lblFlowParametersLabel(3).Caption & " (" & frmBubble!UnitsFlowParam(3) & ")"; Tab(VALUE_TAB); frmBubble!txtFlowParameters(3).Text
          Printer.Print
          Printer.Print frmBubble!lblTankParametersLabel(0).Caption; Tab(VALUE_TAB); frmBubble!txtTankParameters(0).Text
          Printer.Print frmBubble!lblTankParametersLabel(1).Caption & " (" & frmBubble!UnitsTankParam(1) & ")"; Tab(VALUE_TAB); frmBubble!txtTankParameters(1).Text
          Printer.Print frmBubble!lblTankParametersLabel(2).Caption & " (" & frmBubble!UnitsTankParam(2) & ")"; Tab(VALUE_TAB); frmBubble!txtTankParameters(2).Text
          Printer.Print frmBubble!lblTankParametersLabel(3).Caption & " (" & frmBubble!UnitsTankParam(3) & ")"; Tab(VALUE_TAB); frmBubble!txtTankParameters(3).Text
          Printer.Print frmBubble!lblTankParametersLabel(4).Caption & " (" & frmBubble!UnitsTankParam(4) & ")"; Tab(VALUE_TAB); frmBubble!txtTankParameters(4).Text
          Printer.Print
          Printer.Print frmBubble!lblStantonLabel.Caption; Tab(VALUE_TAB); frmBubble!lblStanton.Caption
          Call NewPageBubble
          Printer.Print "Design Contaminant:"; Tab(VALUE_TAB); frmBubble!lblConcentrationResults(0).Caption
          Printer.Print "Liquid Phase Influent Concentration to Tank 1" & " (" & frmBubble!UnitsConcResults(1) & ")"; Tab(VALUE_TAB); frmBubble!lblConcentrationResults(1).Caption
          Printer.Print "Gas Phase Influent Concentration All Tanks" & " (" & frmBubble!UnitsConcResults(2) & ")"; Tab(VALUE_TAB); frmBubble!lblConcentrationResults(2).Caption
          Printer.Print "Liquid Phase Effluent from Last Tank" & " (" & frmBubble!UnitsConcResults(3) & ")"; Tab(VALUE_TAB); frmBubble!lblConcentrationResults(3).Caption
          Printer.Print "Achieved Percent Removal (%)"; Tab(VALUE_TAB); frmBubble!lblConcentrationResults(4).Caption
          Printer.Print
          Printer.Print
          Printer.Print
          Printer.FontBold = True
          Printer.Print "Effluent Concentrations from Each Tank in " & Chr$(181) & "g/L:"
          Printer.Print
          Printer.Print
          Printer.FontUnderline = True
          Printer.Print "Tank:"; Tab(LIQUID_EFFLUENT_TAB); "Liquid Phase"; Tab(GAS_EFFLUENT_TAB); "Gas Phase"
          Printer.Print
          Printer.FontBold = False
          Printer.FontUnderline = False
          For i = 1 To bub.NumberOfTanks.value
              Printer.Print Format$(i, "0"); Tab(LIQUID_EFFLUENT_TAB); Format$(bub.DesignContaminant.Effluent(i), GetTheFormat(bub.DesignContaminant.Effluent(i))); Tab(GAS_EFFLUENT_TAB); Format$(bub.DesignContaminant.GasEffluent(i), GetTheFormat(bub.DesignContaminant.GasEffluent(i)))
          Next i
          Printer.Print
          Printer.Print
          Printer.Print
          Printer.Print
          Printer.FontBold = True
          Call SetPowerBubble
          Printer.Print "Power Calculation:"
          Printer.FontUnderline = True
          Printer.Print
          Printer.Print
          Printer.Print "Property:"; Tab(VALUE_TAB); "Value:"
          Printer.FontBold = False
          Printer.FontUnderline = False
          Printer.Print
          Printer.Print frmBubblePower!lblPowerLabel(0).Caption; Tab(VALUE_TAB); frmBubblePower!txtPower(0).Text
          Printer.Print frmBubblePower!lblPowerLabel(1).Caption; Tab(VALUE_TAB); frmBubblePower!txtPower(1).Text
          Printer.Print frmBubblePower!lblPowerLabel(2).Caption; Tab(VALUE_TAB); frmBubblePower!txtPower(2).Text
          Printer.Print "Blower " & frmBubblePower!lblPowerLabel(3).Caption; Tab(VALUE_TAB); frmBubblePower!lblPower(3).Caption
          Printer.Print frmBubblePower!lblPowerLabel(4).Caption; Tab(VALUE_TAB); frmBubblePower!lblPower(4).Caption
          Printer.Print frmBubblePower!lblPowerLabel(5).Caption; Tab(VALUE_TAB); frmBubblePower!txtPower(5).Text
          Printer.Print frmBubblePower!lblPowerLabel(6).Caption; Tab(VALUE_TAB); frmBubblePower!lblPower(6).Caption

       Case RATING_MODE
          Printer.ScaleLeft = -1440
          Printer.ScaleTop = -1440
          Printer.CurrentX = 0
          Printer.CurrentY = 0
          Printer.FontSize = 12
          Printer.FontBold = True
          Printer.Print "Bubble Aeration - Rating Mode"
          Printer.Print
          Printer.Print
          Printer.FontUnderline = True
          Printer.FontSize = 10
          Printer.Print "Property:"; Tab(VALUE_TAB); "Value:"
          Printer.Print
          Printer.FontUnderline = False
          Printer.FontBold = False
          Printer.Print "Operating Pressure (" & frmBubble!UnitsOpCond(0) & ")"; Tab(VALUE_TAB); frmBubble!txtOperatingPressure.Text
          Printer.Print "Operating Temperature (" & frmBubble!UnitsOpCond(1) & ")"; Tab(VALUE_TAB); frmBubble!txtOperatingTemperature.Text
          Printer.Print frmWaterPropertiesBubble!lblAirWaterProperties(0).Caption; Tab(VALUE_TAB); frmWaterPropertiesBubble!txtAirWaterProperties(0).Text
          Printer.Print frmWaterPropertiesBubble!lblAirWaterProperties(1).Caption; Tab(VALUE_TAB); frmWaterPropertiesBubble!txtAirWaterProperties(1).Text
          Printer.Print
          Printer.Print "Oxygen " & frmBubble!lblOxygenLabel(1).Caption & " (" & frmBubble!UnitsOxygenRef(1) & ")"; Tab(VALUE_TAB); frmBubble!txtOxygen(1).Text
          Printer.Print "Method to Find Oxygen KLa"; Tab(VALUE_TAB); frmBubble!cboOxygen.Text
          If frmBubble!cboOxygen.ListIndex = 0 Then   'Method to Find Oxygen KLa is Clean Water Oxygen Transfer Test Data
             Printer.Print "Standardized Oxygen Transfer Efficiency, SOTE (%)"; Tab(VALUE_TAB); frmOxygenMassTransferCoeff!txtDataParameters(0).Text; ""
             Printer.Print "Standardized Oxygen Transfer Rate (kg O2/d)"; Tab(VALUE_TAB); frmOxygenMassTransferCoeff!txtDataParameters(1).Text
             Printer.Print "Air Flow Rate (standard m" & Chr$(179) & "/hr)"; Tab(VALUE_TAB); frmOxygenMassTransferCoeff!txtDataParameters(2).Text
             Printer.Print "Barometric Pressure (Pa)"; Tab(VALUE_TAB); frmOxygenMassTransferCoeff!txtDataParameters(3).Text
             Printer.Print "Tank Water Depth (m" & Chr$(179) & ")"; Tab(VALUE_TAB); frmOxygenMassTransferCoeff!txtDataParameters(4).Text
             Printer.Print "Tank Water Volume (m" & Chr$(179) & ")"; Tab(VALUE_TAB); frmOxygenMassTransferCoeff!txtDataParameters(5).Text
             Printer.Print "D.O. Saturation Concentration at Infinite Time (mg/L)"; Tab(VALUE_TAB); frmOxygenMassTransferCoeff!lblDataParameters(6).Caption; ""
             Printer.Print frmOxygenMassTransferCoeff!lblDataParametersLabel(7).Caption; Tab(VALUE_TAB); frmOxygenMassTransferCoeff!lblDataParameters(7).Caption
             Printer.Print frmOxygenMassTransferCoeff!lblDataParametersLabel(8).Caption; Tab(VALUE_TAB); frmOxygenMassTransferCoeff!lblDataParameters(8).Caption
             Printer.Print frmOxygenMassTransferCoeff!lblDataParametersLabel(9).Caption; Tab(VALUE_TAB); frmOxygenMassTransferCoeff!lblDataParameters(9).Caption
             Printer.Print frmOxygenMassTransferCoeff!lblDataParametersLabel(10).Caption; Tab(VALUE_TAB); frmOxygenMassTransferCoeff!lblDataParameters(10).Caption
             Printer.Print frmOxygenMassTransferCoeff!lblDataParametersLabel(11).Caption; Tab(VALUE_TAB); frmOxygenMassTransferCoeff!lblDataParameters(11).Caption
          End If
          Printer.Print "Oxygen " & frmBubble!lblOxygenLabel(2).Caption; Tab(VALUE_TAB); frmBubble!txtOxygen(2).Text
          Printer.Print
          Printer.Print frmBubble!lblFlowParametersLabel(0).Caption & " (" & frmBubble!UnitsFlowParam(0) & ")"; Tab(VALUE_TAB); frmBubble!txtFlowParameters(0).Text
          Printer.Print frmBubble!lblFlowParametersLabel(2).Caption; Tab(VALUE_TAB); frmBubble!txtFlowParameters(2).Text
          Printer.Print frmBubble!lblFlowParametersLabel(3).Caption & " (" & frmBubble!UnitsFlowParam(3) & ")"; Tab(VALUE_TAB); frmBubble!txtFlowParameters(3).Text
          Printer.Print
          Printer.Print frmBubble!lblTankParametersLabel(0).Caption; Tab(VALUE_TAB); frmBubble!txtTankParameters(0).Text
          Printer.Print frmBubble!lblTankParametersLabel(1).Caption & " (" & frmBubble!UnitsTankParam(1) & ")"; Tab(VALUE_TAB); frmBubble!txtTankParameters(1).Text
          Printer.Print frmBubble!lblTankParametersLabel(2).Caption & " (" & frmBubble!UnitsTankParam(2) & ")"; Tab(VALUE_TAB); frmBubble!txtTankParameters(2).Text
          Printer.Print frmBubble!lblTankParametersLabel(3).Caption & " (" & frmBubble!UnitsTankParam(3) & ")"; Tab(VALUE_TAB); frmBubble!txtTankParameters(3).Text
          Printer.Print frmBubble!lblTankParametersLabel(4).Caption & " (" & frmBubble!UnitsTankParam(4) & ")"; Tab(VALUE_TAB); frmBubble!txtTankParameters(4).Text
          Printer.Print
          Printer.Print
          Printer.FontBold = True
          Call SetPowerBubble
          Printer.Print "Power Calculation:"
          Printer.FontUnderline = True
          Printer.Print
          Printer.Print "Property:"; Tab(VALUE_TAB); "Value:"
          Printer.FontBold = False
          Printer.FontUnderline = False
          Printer.Print
          Printer.Print frmBubblePower!lblPowerLabel(0).Caption; Tab(VALUE_TAB); frmBubblePower!txtPower(0).Text
          Printer.Print frmBubblePower!lblPowerLabel(1).Caption; Tab(VALUE_TAB); frmBubblePower!txtPower(1).Text
          Printer.Print frmBubblePower!lblPowerLabel(2).Caption; Tab(VALUE_TAB); frmBubblePower!txtPower(2).Text
          Printer.Print "Blower " & frmBubblePower!lblPowerLabel(3).Caption; Tab(VALUE_TAB); frmBubblePower!lblPower(3).Caption
          Printer.Print frmBubblePower!lblPowerLabel(4).Caption; Tab(VALUE_TAB); frmBubblePower!lblPower(4).Caption
          Printer.Print frmBubblePower!lblPowerLabel(5).Caption; Tab(VALUE_TAB); frmBubblePower!txtPower(5).Text
          Printer.Print frmBubblePower!lblPowerLabel(6).Caption; Tab(VALUE_TAB); frmBubblePower!lblPower(6).Caption
          Printer.Print
          Printer.Print
          Printer.FontUnderline = True
          Printer.Print "Contaminant Glossary:"
          Printer.FontUnderline = False
          For i = 1 To bub.NumChemical
              Printer.Print Format$(i, "0"); " = "; Trim$(bub.Contaminant(i).Name)
          Next i

          Call NewPageBubble
          Printer.FontBold = True
          Printer.Print "Contaminant Properties:"
          Printer.Print
          Printer.FontUnderline = True
          Printer.Print "Con.:"; Tab(MWT_TAB); "MWT"; Tab(HC_TAB); "HC"; Tab(VB_TAB); "Vb"; Tab(DIFL_TAB); "DIFL"; Tab(MTCOEFF_TAB); "MT Coeff."; Tab(STANTON_TAB); "St."
          Printer.Print
          Printer.FontUnderline = False
          Printer.FontBold = False
          For i = 1 To bub.NumChemical
              If bub.DesignContaminant.Name = bub.Contaminant(i).Name Then
                 ContaminantMTCoeff(i) = bub.ContaminantMassTransferCoefficient.value
                 StantonNumber(i) = bub.StantonNumber.value
              Else
                 Call KLABUB(ContaminantMTCoeff(i), bub.Oxygen.MassTransferCoefficient.value, bub.Contaminant(i).LiquidDiffusivity.value, bub.Oxygen.LiquidDiffusivity.value, bub.N_for_Finding_KLa.value, bub.kgOVERkl_for_Finding_KLa.value, bub.Contaminant(i).HenrysConstant.value)
                 Call GETPHIB(StantonNumber(i), ContaminantMTCoeff(i), bub.TankVolume.value, bub.Contaminant(i).HenrysConstant.value, bub.AirFlowRate.value)
              End If
              Printer.Print Format$(i, "0"); Tab(MWT_TAB); Format$(bub.Contaminant(i).MolecularWeight.value, "0.00"); Tab(HC_TAB); Format$(bub.Contaminant(i).HenrysConstant.value, GetTheFormat(bub.Contaminant(i).HenrysConstant.value)); Tab(VB_TAB); Format$(bub.Contaminant(i).MolarVolume.value, GetTheFormat(bub.Contaminant(i).MolarVolume.value)); Tab(DIFL_TAB); Format$(bub.Contaminant(i).LiquidDiffusivity.value, GetTheFormat(bub.Contaminant(i).LiquidDiffusivity.value)); Tab(MTCOEFF_TAB); Format$(ContaminantMTCoeff(i), GetTheFormat(ContaminantMTCoeff(i))); Tab(STANTON_TAB); Format$(StantonNumber(i), GetTheFormat(StantonNumber(i)))
          Next i
          Printer.Print
          Printer.FontUnderline = True
          Printer.Print "Glossary:"
          Printer.FontUnderline = False
          Printer.Print "Con. = Contaminant Number (see Contaminant Glossary on page 1)"
          Printer.Print "MWT = Molecular Weight (kg/kmol)"
          Printer.Print "HC = Henry's Constant (-)"
          Printer.Print "Vb = Molar Volume (m" & Chr$(179) & "/kmol)"
          Printer.Print "DIFL = Liquid Diffusivity (m" & Chr$(178) & "/sec)"
          Printer.Print "MT Coeff. = Mass Transfer Coeff. (1/sec)"
          Printer.Print "St. = Stanton Number (-)"
          Printer.Print
          Printer.Print
          Printer.Print
          Printer.FontBold = True
          Printer.Print "Contaminant Concentration Results:"
          Printer.Print
          Printer.FontUnderline = True
          Printer.Print "Con.:"; Tab(MWT_TAB); "Cinf"; Tab(HC_TAB); "Cto"; Tab(VB_TAB); "De. % Rem."; Tab(DIFL_TAB); "Ceff"; Tab(MTCOEFF_TAB); "Ach. % Rem."
          Printer.Print
          Printer.FontUnderline = False
          Printer.FontBold = False
          For i = 1 To bub.NumChemical
              If bub.DesignContaminant.Name = bub.Contaminant(i).Name Then
                 DesiredPercentRemoval(i) = bub.DesiredPercentRemoval
                 bub.Contaminant(i).Effluent(0) = bub.DesignContaminant.Effluent(0)
                 For j = 1 To bub.NumberOfTanks.value
                     bub.Contaminant(i).Effluent(j) = bub.DesignContaminant.Effluent(j)
                     bub.Contaminant(i).GasEffluent(j) = bub.DesignContaminant.GasEffluent(j)
                 Next j
                 AchievedPercentRemoval(i) = bub.AchievedPercentRemoval
              Else
                 Call REMOVBUB(DesiredPercentRemoval(i), bub.Contaminant(i).Influent.value, bub.Contaminant(i).TreatmentObjective.value)
                 Call EFFLBUB(Effluent(0), GasEffluent(1), bub.Contaminant(i).HenrysConstant.value, bub.Contaminant(i).Influent.value, bub.AirToWaterRatio.value, bub.NumberOfTanks.value, StantonNumber(i))
                 bub.Contaminant(i).Effluent(0) = Effluent(0)
                 For j = 1 To bub.NumberOfTanks.value
                     bub.Contaminant(i).Effluent(j) = Effluent(j)
                     bub.Contaminant(i).GasEffluent(j) = GasEffluent(j)
                 Next j
                 Call REMOVBUB(AchievedPercentRemoval(i), bub.Contaminant(i).Influent.value, bub.Contaminant(i).Effluent(bub.NumberOfTanks.value))
              End If
              Printer.Print Format$(i, "0"); Tab(MWT_TAB); Format$(bub.Contaminant(i).Influent.value, GetTheFormat(bub.Contaminant(i).Influent.value)); Tab(HC_TAB); Format$(bub.Contaminant(i).TreatmentObjective.value, GetTheFormat(bub.Contaminant(i).TreatmentObjective.value)); Tab(VB_TAB); Format$(DesiredPercentRemoval(i), GetTheFormat(DesiredPercentRemoval(i))); Tab(DIFL_TAB); Format$(bub.Contaminant(i).Effluent(bub.NumberOfTanks.value), GetTheFormat(bub.Contaminant(i).Effluent(bub.NumberOfTanks.value))); Tab(MTCOEFF_TAB); Format$(AchievedPercentRemoval(i), GetTheFormat(AchievedPercentRemoval(i)))
          Next i
          Printer.Print
          Printer.FontUnderline = True
          Printer.Print "Glossary:"
          Printer.FontUnderline = False
          Printer.Print "Con. = Contaminant Number (see Contaminant Glossary on page 1)"
          Printer.Print "Cinf = "; "Liquid Phase " & frmBubble!lblDesignConcentration(0).Caption
          Printer.Print "Cto = "; frmBubble!lblDesignConcentration(1).Caption
          Printer.Print "De. % Rem. = "; frmBubble!lblDesignConcentration(2).Caption
          Printer.Print "Ceff = "; "Liquid Phase Effluent from Last Tank (" & Chr$(181) & "g/L)"
          Printer.Print "Ach. % Rem. = "; frmBubble!lblConcentrationResultsLabel(4).Caption
          Call NewPageBubble
          Printer.FontBold = True
          Printer.Print "Liquid Phase Effluent Concentrations from Each Tank in " & Chr$(181) & "g/L:"
          Printer.Print
          Printer.Print
          Printer.Print Tab(MWT_TAB); "Contaminant Number:"
          Printer.Print
          Printer.FontUnderline = True
          Select Case bub.NumChemical
             Case 1
                Printer.Print "Tank:"; Tab(MWT_TAB); "1:"
             Case 2
                Printer.Print "Tank:"; Tab(MWT_TAB); "1:"; Tab(HC_TAB); "2:"
             Case 3
                Printer.Print "Tank:"; Tab(MWT_TAB); "1:"; Tab(HC_TAB); "2:"; Tab(VB_TAB); "3:"
             Case 4
                Printer.Print "Tank:"; Tab(MWT_TAB); "1:"; Tab(HC_TAB); "2:"; Tab(VB_TAB); "3:"; Tab(DIFL_TAB); "4:"
             Case 5
                Printer.Print "Tank:"; Tab(MWT_TAB); "1:"; Tab(HC_TAB); "2:"; Tab(VB_TAB); "3:"; Tab(DIFL_TAB); "4:"; Tab(MTCOEFF_TAB); "5:"
             Case Else
                Printer.Print "Tank:"; Tab(MWT_TAB); "1:"; Tab(HC_TAB); "2:"; Tab(VB_TAB); "3:"; Tab(DIFL_TAB); "4:"; Tab(MTCOEFF_TAB); "5:"; Tab(STANTON_TAB); "6:"
          End Select
          Printer.Print
          Printer.FontBold = False
          Printer.FontUnderline = False

          'Print Liquid Phase Influent Concentrations of Each Contaminant
          Printer.Print "Cinf";
          For j = 1 To 6
              If bub.NumChemical < j Then
                 Exit For
              End If
              Select Case j
                 Case 1
                    Printer.Print Tab(MWT_TAB);
                 Case 2
                    Printer.Print Tab(HC_TAB);
                 Case 3
                    Printer.Print Tab(VB_TAB);
                 Case 4
                    Printer.Print Tab(DIFL_TAB);
                 Case 5
                    Printer.Print Tab(MTCOEFF_TAB);
                 Case 6
                    Printer.Print Tab(STANTON_TAB);
              End Select
              Printer.Print Format$(bub.Contaminant(j).Influent.value, GetTheFormat(bub.Contaminant(j).Influent.value));
          Next j
          Printer.Print

          'Print Liquid Phase Effluent Concentrations from each tank for each contaminant
          For i = 1 To bub.NumberOfTanks.value
              Printer.Print Format$(i, "0");
              For j = 1 To 6
                  If bub.NumChemical < j Then
                     Exit For
                  End If
                  Select Case j
                     Case 1
                        Printer.Print Tab(MWT_TAB);
                     Case 2
                        Printer.Print Tab(HC_TAB);
                     Case 3
                        Printer.Print Tab(VB_TAB);
                     Case 4
                        Printer.Print Tab(DIFL_TAB);
                     Case 5
                        Printer.Print Tab(MTCOEFF_TAB);
                     Case 6
                        Printer.Print Tab(STANTON_TAB);
                  End Select
                  Printer.Print Format$(bub.Contaminant(j).Effluent(i), GetTheFormat(bub.Contaminant(j).Effluent(i)));
             Next j
             Printer.Print
          Next i
          
          If bub.NumChemical < 7 Then
             Printer.Print
             Printer.FontUnderline = True
             Printer.Print "Glossary:"
             Printer.FontUnderline = False
             Printer.Print "Cinf = Liquid Phase Influent Concentration to Tank 1 (" & Chr$(181) & "g/L)"
             GoTo AfterLiquidEffluents
          End If
          Printer.Print
          Printer.Print
          Printer.FontBold = True
          Printer.Print Tab(MWT_TAB); "Contaminant Number:"
          Printer.Print
          Printer.FontUnderline = True
          Select Case bub.NumChemical
             Case 7
                Printer.Print "Tank:"; Tab(MWT_TAB); "7:"
             Case 8
                Printer.Print "Tank:"; Tab(MWT_TAB); "7:"; Tab(HC_TAB); "8:"
             Case 9
                Printer.Print "Tank:"; Tab(MWT_TAB); "7:"; Tab(HC_TAB); "8:"; Tab(VB_TAB); "9:"
             Case 10
                Printer.Print "Tank:"; Tab(MWT_TAB); "7:"; Tab(HC_TAB); "8:"; Tab(VB_TAB); "9:"; Tab(DIFL_TAB); "10:"
          End Select
          Printer.Print
          Printer.FontBold = False
          Printer.FontUnderline = False

          'Print Liquid Phase Influent Concentrations of Each Contaminant
          Printer.Print "Cinf";
          For j = 7 To 10
              If bub.NumChemical < j Then
                 Exit For
              End If
              Select Case j
                 Case 7
                    Printer.Print Tab(MWT_TAB);
                 Case 8
                    Printer.Print Tab(HC_TAB);
                 Case 9
                    Printer.Print Tab(VB_TAB);
                 Case 10
                    Printer.Print Tab(DIFL_TAB);
              End Select
              Printer.Print Format$(bub.Contaminant(j).Influent.value, GetTheFormat(bub.Contaminant(j).Influent.value));
          Next j
          Printer.Print

          'Print Liquid Phase Effluent Concentrations from each tank for each contaminant
          For i = 1 To bub.NumberOfTanks.value
              Printer.Print Format$(i, "0");
              For j = 7 To 10
                  If bub.NumChemical < j Then
                     Exit For
                  End If
                  Select Case j
                     Case 7
                        Printer.Print Tab(MWT_TAB);
                     Case 8
                        Printer.Print Tab(HC_TAB);
                     Case 9
                        Printer.Print Tab(VB_TAB);
                     Case 10
                        Printer.Print Tab(DIFL_TAB);
                  End Select
                  Printer.Print Format$(bub.Contaminant(j).Effluent(i), GetTheFormat(bub.Contaminant(j).Effluent(i)));
             Next j
             Printer.Print
          Next i
          
          Printer.Print
          Printer.FontUnderline = True
          Printer.Print "Glossary:"
          Printer.FontUnderline = False
          Printer.Print "Cinf = Liquid Phase Influent Concentration to Tank 1 (" & Chr$(181) & "g/L)"

AfterLiquidEffluents:

          If bub.NumChemical > 6 Then
             If bub.NumberOfTanks.value > 5 Then
                Call NewPageBubble
             Else
                Printer.Print
                Printer.Print
                Printer.Print
                Printer.Print
             End If
          Else
             Printer.Print
             Printer.Print
             Printer.Print
             Printer.Print
          End If
          
          Printer.FontBold = True
          Printer.Print "Gas Phase Effluent Concentrations from Each Tank in " & Chr$(181) & "g/L:"
          Printer.Print
          Printer.Print
          Printer.Print Tab(MWT_TAB); "Contaminant Number:"
          Printer.Print
          Printer.FontUnderline = True
          Select Case bub.NumChemical
             Case 1
                Printer.Print "Tank:"; Tab(MWT_TAB); "1:"
             Case 2
                Printer.Print "Tank:"; Tab(MWT_TAB); "1:"; Tab(HC_TAB); "2:"
             Case 3
                Printer.Print "Tank:"; Tab(MWT_TAB); "1:"; Tab(HC_TAB); "2:"; Tab(VB_TAB); "3:"
             Case 4
                Printer.Print "Tank:"; Tab(MWT_TAB); "1:"; Tab(HC_TAB); "2:"; Tab(VB_TAB); "3:"; Tab(DIFL_TAB); "4:"
             Case 5
                Printer.Print "Tank:"; Tab(MWT_TAB); "1:"; Tab(HC_TAB); "2:"; Tab(VB_TAB); "3:"; Tab(DIFL_TAB); "4:"; Tab(MTCOEFF_TAB); "5:"
             Case Else
                Printer.Print "Tank:"; Tab(MWT_TAB); "1:"; Tab(HC_TAB); "2:"; Tab(VB_TAB); "3:"; Tab(DIFL_TAB); "4:"; Tab(MTCOEFF_TAB); "5:"; Tab(STANTON_TAB); "6:"
          End Select
          Printer.Print
          Printer.FontBold = False
          Printer.FontUnderline = False

          'Print Gas Phase Influent Concentrations of Each Contaminant
          Printer.Print "Yinf";
          For j = 1 To 6
              If bub.NumChemical < j Then
                 Exit For
              End If
              Select Case j
                 Case 1
                    Printer.Print Tab(MWT_TAB);
                 Case 2
                    Printer.Print Tab(HC_TAB);
                 Case 3
                    Printer.Print Tab(VB_TAB);
                 Case 4
                    Printer.Print Tab(DIFL_TAB);
                 Case 5
                    Printer.Print Tab(MTCOEFF_TAB);
                 Case 6
                    Printer.Print Tab(STANTON_TAB);
              End Select
              Printer.Print "0";
          Next j
          Printer.Print

          'Print Gas Phase Effluent Concentrations from each tank for each contaminant
          For i = 1 To bub.NumberOfTanks.value
              Printer.Print Format$(i, "0");
              For j = 1 To 6
                  If bub.NumChemical < j Then
                     Exit For
                  End If
                  Select Case j
                     Case 1
                        Printer.Print Tab(MWT_TAB);
                     Case 2
                        Printer.Print Tab(HC_TAB);
                     Case 3
                        Printer.Print Tab(VB_TAB);
                     Case 4
                        Printer.Print Tab(DIFL_TAB);
                     Case 5
                        Printer.Print Tab(MTCOEFF_TAB);
                     Case 6
                        Printer.Print Tab(STANTON_TAB);
                  End Select
                  Printer.Print Format$(bub.Contaminant(j).GasEffluent(i), GetTheFormat(bub.Contaminant(j).GasEffluent(i)));
             Next j
             Printer.Print
          Next i
          
          If bub.NumChemical < 7 Then
             Printer.Print
             Printer.FontUnderline = True
             Printer.Print "Glossary:"
             Printer.FontUnderline = False
             Printer.Print "Yinf = Gas Phase Influent Concentration to All Tanks (" & Chr$(181) & "g/L)"
             GoTo AfterGasEffluents
          End If
          Printer.Print
          Printer.Print
          Printer.FontBold = True
          Printer.Print Tab(MWT_TAB); "Contaminant Number:"
          Printer.Print
          Printer.FontUnderline = True
          Select Case bub.NumChemical
             Case 7
                Printer.Print "Tank:"; Tab(MWT_TAB); "7:"
             Case 8
                Printer.Print "Tank:"; Tab(MWT_TAB); "7:"; Tab(HC_TAB); "8:"
             Case 9
                Printer.Print "Tank:"; Tab(MWT_TAB); "7:"; Tab(HC_TAB); "8:"; Tab(VB_TAB); "9:"
             Case 10
                Printer.Print "Tank:"; Tab(MWT_TAB); "7:"; Tab(HC_TAB); "8:"; Tab(VB_TAB); "9:"; Tab(DIFL_TAB); "10:"
          End Select
          Printer.Print
          Printer.FontBold = False
          Printer.FontUnderline = False

          'Print Gas Phase Influent Concentrations of Each Contaminant
          Printer.Print "Yinf";
          For j = 7 To 10
              If bub.NumChemical < j Then
                 Exit For
              End If
              Select Case j
                 Case 7
                    Printer.Print Tab(MWT_TAB);
                 Case 8
                    Printer.Print Tab(HC_TAB);
                 Case 9
                    Printer.Print Tab(VB_TAB);
                 Case 10
                    Printer.Print Tab(DIFL_TAB);
              End Select
              Printer.Print "0";
          Next j
          Printer.Print

          'Print Gas Phase Effluent Concentrations from each tank for each contaminant
          For i = 1 To bub.NumberOfTanks.value
              Printer.Print Format$(i, "0");
              For j = 7 To 10
                  If bub.NumChemical < j Then
                     Exit For
                  End If
                  Select Case j
                     Case 7
                        Printer.Print Tab(MWT_TAB);
                     Case 8
                        Printer.Print Tab(HC_TAB);
                     Case 9
                        Printer.Print Tab(VB_TAB);
                     Case 10
                        Printer.Print Tab(DIFL_TAB);
                  End Select
                  Printer.Print Format$(bub.Contaminant(j).GasEffluent(i), GetTheFormat(bub.Contaminant(j).GasEffluent(i)));
             Next j
             Printer.Print
          Next i

          Printer.Print
          Printer.FontUnderline = True
          Printer.Print "Glossary:"
          Printer.FontUnderline = False
          Printer.Print "Yinf = Gas Phase Influent Concentration to All Tanks (" & Chr$(181) & "g/L)"
          

AfterGasEffluents:

    End Select

    Printer.EndDoc

    Exit Sub

PrinterError:
    MsgBox error$(Err)
    Resume ExitPrint:

ExitPrint:

End Sub

Sub PrintBubbleToFile()
    Dim i As Integer, j As Integer
    ReDim ContaminantMTCoeff(1 To MAXCHEMICAL) As Double
    ReDim StantonNumber(1 To MAXCHEMICAL) As Double
    ReDim DesiredPercentRemoval(1 To MAXCHEMICAL) As Double
    ReDim AchievedPercentRemoval(1 To MAXCHEMICAL) As Double
    ReDim Effluent(0 To MAXIMUM_TANKS) As Double
    ReDim GasEffluent(1 To MAXIMUM_TANKS) As Double

        Call GetPrintFileName(PrintFileName)
        If PrintFileName$ = "" Then Exit Sub

        Open PrintFileName$ For Output As #1

    Select Case BubbleAerationMode
       Case DESIGN_MODE

          Print #1, "Bubble Aeration - Design Mode"
          Print #1,
          Print #1,
          Print #1, "Property:"; Tab(VALUE_TAB); "Value:"
          Print #1,
          Print #1, "Operating Pressure (" & frmBubble!UnitsOpCond(0) & ")"; Tab(VALUE_TAB); frmBubble!txtOperatingPressure.Text
          Print #1, "Operating Temperature (" & frmBubble!UnitsOpCond(1) & ")"; Tab(VALUE_TAB); frmBubble!txtOperatingTemperature.Text
          Print #1, frmWaterPropertiesBubble!lblAirWaterProperties(0).Caption; Tab(VALUE_TAB); frmWaterPropertiesBubble!txtAirWaterProperties(0).Text
          Print #1, frmWaterPropertiesBubble!lblAirWaterProperties(1).Caption; Tab(VALUE_TAB); frmWaterPropertiesBubble!txtAirWaterProperties(1).Text
          Print #1,
          Print #1, "Oxygen " & frmBubble!lblOxygenLabel(1).Caption & " (" & frmBubble!UnitsOxygenRef(1) & ")"; Tab(VALUE_TAB); frmBubble!txtOxygen(1).Text
          Print #1, "Method to Find Oxygen KLa"; Tab(VALUE_TAB); frmBubble!cboOxygen.Text
          If frmBubble!cboOxygen.ListIndex = 0 Then   'Method to Find Oxygen KLa is Clean Water Oxygen Transfer Test Data
             Print #1, "Standardized Oxygen Transfer Efficiency, SOTE (%)"; Tab(VALUE_TAB); frmOxygenMassTransferCoeff!txtDataParameters(0).Text; ""
             Print #1, "Standardized Oxygen Transfer Rate (kg O2/d)"; Tab(VALUE_TAB); frmOxygenMassTransferCoeff!txtDataParameters(1).Text
             Print #1, "Air Flow Rate (standard m" & Chr$(179) & "/hr)"; Tab(VALUE_TAB); frmOxygenMassTransferCoeff!txtDataParameters(2).Text
             Print #1, "Barometric Pressure (Pa)"; Tab(VALUE_TAB); frmOxygenMassTransferCoeff!txtDataParameters(3).Text
             Print #1, "Tank Water Depth (m" & Chr$(179) & ")"; Tab(VALUE_TAB); frmOxygenMassTransferCoeff!txtDataParameters(4).Text
             Print #1, "Tank Water Volume (m" & Chr$(179) & ")"; Tab(VALUE_TAB); frmOxygenMassTransferCoeff!txtDataParameters(5).Text
             Print #1, "D.O. Saturation Concentration at Infinite Time (mg/L)"; Tab(VALUE_TAB); frmOxygenMassTransferCoeff!lblDataParameters(6).Caption; ""
             Print #1, frmOxygenMassTransferCoeff!lblDataParametersLabel(7).Caption; Tab(VALUE_TAB); frmOxygenMassTransferCoeff!lblDataParameters(7).Caption
             Print #1, frmOxygenMassTransferCoeff!lblDataParametersLabel(8).Caption; Tab(VALUE_TAB); frmOxygenMassTransferCoeff!lblDataParameters(8).Caption
             Print #1, frmOxygenMassTransferCoeff!lblDataParametersLabel(9).Caption; Tab(VALUE_TAB); frmOxygenMassTransferCoeff!lblDataParameters(9).Caption
             Print #1, frmOxygenMassTransferCoeff!lblDataParametersLabel(10).Caption; Tab(VALUE_TAB); frmOxygenMassTransferCoeff!lblDataParameters(10).Caption
             Print #1, frmOxygenMassTransferCoeff!lblDataParametersLabel(11).Caption; Tab(VALUE_TAB); frmOxygenMassTransferCoeff!lblDataParameters(11).Caption
          End If
          Print #1, "Oxygen " & frmBubble!lblOxygenLabel(2).Caption & " (" & frmBubble!UnitsOxygenRef(2) & ")"; Tab(VALUE_TAB); frmBubble!txtOxygen(2).Text
          Print #1,
          Print #1, "Design Contaminant:  "; frmBubble!cboDesignContaminant.Text
          Print #1, "Molecular Weight (kg/kmol)"; Tab(VALUE_TAB); Format$(bub.DesignContaminant.MolecularWeight.value, "0.00")
          Print #1, "Henry's Constant (-)"; Tab(VALUE_TAB); Format$(bub.DesignContaminant.HenrysConstant.value, GetTheFormat(bub.DesignContaminant.HenrysConstant.value))
          Print #1, "Molar Volume (m" & Chr$(179) & "/kmol)"; Tab(VALUE_TAB); Format$(bub.DesignContaminant.MolarVolume.value, GetTheFormat(bub.DesignContaminant.MolarVolume.value))
          Print #1, "Liquid Diffusivity (m" & Chr$(178) & "/sec)"; Tab(VALUE_TAB); Format$(bub.DesignContaminant.LiquidDiffusivity.value, GetTheFormat(bub.DesignContaminant.LiquidDiffusivity.value))
          Print #1, frmBubble!lblDesignConcentration(0).Caption & " (" & frmBubble!UnitsDesignContam(0) & ")"; Tab(VALUE_TAB); frmBubble!lblDesignConcentrationValue(0).Caption
          Print #1, frmBubble!lblDesignConcentration(1).Caption & " (" & frmBubble!UnitsDesignContam(1) & ")"; Tab(VALUE_TAB); frmBubble!lblDesignConcentrationValue(1).Caption
          Print #1, frmBubble!lblDesignConcentration(2).Caption; Tab(VALUE_TAB); frmBubble!lblDesignConcentrationValue(2).Caption
          Print #1, frmBubble!lblDesignConcentration(3).Caption & " (" & frmBubble!UnitsDesignContam(3) & ")"; Tab(VALUE_TAB); frmBubble!txtDesignConcentrationValue(3).Text
          Print #1,
          Print #1, frmBubble!lblFlowParametersLabel(0).Caption & " (" & frmBubble!UnitsFlowParam(0) & ")"; Tab(VALUE_TAB); frmBubble!txtFlowParameters(0).Text
          Print #1, frmBubble!lblFlowParametersLabel(1).Caption; Tab(VALUE_TAB); frmBubble!lblFlowParameters(1).Caption
          Print #1, frmBubble!lblFlowParametersLabel(2).Caption; Tab(VALUE_TAB); frmBubble!txtFlowParameters(2).Text
          Print #1, frmBubble!lblFlowParametersLabel(3).Caption & " (" & frmBubble!UnitsFlowParam(3) & ")"; Tab(VALUE_TAB); frmBubble!txtFlowParameters(3).Text
          Print #1,
          Print #1, frmBubble!lblTankParametersLabel(0).Caption; Tab(VALUE_TAB); frmBubble!txtTankParameters(0).Text
          Print #1, frmBubble!lblTankParametersLabel(1).Caption & " (" & frmBubble!UnitsTankParam(1) & ")"; Tab(VALUE_TAB); frmBubble!txtTankParameters(1).Text
          Print #1, frmBubble!lblTankParametersLabel(2).Caption & " (" & frmBubble!UnitsTankParam(2) & ")"; Tab(VALUE_TAB); frmBubble!txtTankParameters(2).Text
          Print #1, frmBubble!lblTankParametersLabel(3).Caption & " (" & frmBubble!UnitsTankParam(3) & ")"; Tab(VALUE_TAB); frmBubble!txtTankParameters(3).Text
          Print #1, frmBubble!lblTankParametersLabel(4).Caption & " (" & frmBubble!UnitsTankParam(4) & ")"; Tab(VALUE_TAB); frmBubble!txtTankParameters(4).Text
          Print #1,
          Print #1, frmBubble!lblStantonLabel.Caption; Tab(VALUE_TAB); frmBubble!lblStanton.Caption
          Print #1,
          Print #1,
          Print #1, "Design Contaminant:"; Tab(VALUE_TAB); frmBubble!lblConcentrationResults(0).Caption
          Print #1, "Liquid Phase Influent Concentration to Tank 1 (" & Chr$(181) & "g/L)"; Tab(VALUE_TAB); frmBubble!lblConcentrationResults(1).Caption
          Print #1, "Gas Phase Influent Concentration All Tanks (" & Chr$(181) & "g/L)"; Tab(VALUE_TAB); frmBubble!lblConcentrationResults(2).Caption
          Print #1, "Liquid Phase Effluent from Last Tank (" & Chr$(181) & "g/L)"; Tab(VALUE_TAB); frmBubble!lblConcentrationResults(3).Caption
          Print #1, "Achieved Percent Removal (%)"; Tab(VALUE_TAB); frmBubble!lblConcentrationResults(4).Caption
          Print #1,
          Print #1,
          Print #1,
          Print #1, "Effluent Concentrations from Each Tank in " & Chr$(181) & "g/L:"
          Print #1,
          Print #1,
          Print #1, "Tank:"; Tab(LIQUID_EFFLUENT_TAB); "Liquid Phase"; Tab(GAS_EFFLUENT_TAB); "Gas Phase"
          Print #1,
          For i = 1 To bub.NumberOfTanks.value
              Print #1, Format$(i, "0"); Tab(LIQUID_EFFLUENT_TAB); Format$(bub.DesignContaminant.Effluent(i), GetTheFormat(bub.DesignContaminant.Effluent(i))); Tab(GAS_EFFLUENT_TAB); Format$(bub.DesignContaminant.GasEffluent(i), GetTheFormat(bub.DesignContaminant.GasEffluent(i)))
          Next i
          Print #1,
          Print #1,
          Print #1,
          Print #1,
          Print #1,
          Print #1,
          Print #1,
          Call SetPowerBubble
          Print #1, "Power Calculation:"
          Print #1,
          Print #1,
          Print #1, "Property:"; Tab(VALUE_TAB); "Value:"
          Print #1,
          Print #1, frmBubblePower!lblPowerLabel(0).Caption; Tab(VALUE_TAB); frmBubblePower!txtPower(0).Text
          Print #1, frmBubblePower!lblPowerLabel(1).Caption; Tab(VALUE_TAB); frmBubblePower!txtPower(1).Text
          Print #1, frmBubblePower!lblPowerLabel(2).Caption; Tab(VALUE_TAB); frmBubblePower!txtPower(2).Text
          Print #1, "Blower " & frmBubblePower!lblPowerLabel(3).Caption; Tab(VALUE_TAB); frmBubblePower!lblPower(3).Caption
          Print #1, frmBubblePower!lblPowerLabel(4).Caption; Tab(VALUE_TAB); frmBubblePower!lblPower(4).Caption
          Print #1, frmBubblePower!lblPowerLabel(5).Caption; Tab(VALUE_TAB); frmBubblePower!txtPower(5).Text
          Print #1, frmBubblePower!lblPowerLabel(6).Caption; Tab(VALUE_TAB); frmBubblePower!lblPower(6).Caption

       Case RATING_MODE
          Print #1, "Bubble Aeration - Rating Mode"
          Print #1,
          Print #1,
          Print #1, "Property:"; Tab(VALUE_TAB); "Value:"
          Print #1,
          Print #1, "Operating Pressure (" & frmBubble!UnitsOpCond(0) & ")"; Tab(VALUE_TAB); frmBubble!txtOperatingPressure.Text
          Print #1, "Operating Temperature (" & frmBubble!UnitsOpCond(1) & ")"; Tab(VALUE_TAB); frmBubble!txtOperatingTemperature.Text
          Print #1, frmWaterPropertiesBubble!lblAirWaterProperties(0).Caption; Tab(VALUE_TAB); frmWaterPropertiesBubble!txtAirWaterProperties(0).Text
          Print #1, frmWaterPropertiesBubble!lblAirWaterProperties(1).Caption; Tab(VALUE_TAB); frmWaterPropertiesBubble!txtAirWaterProperties(1).Text
          Print #1,
          Print #1, "Oxygen " & frmBubble!lblOxygenLabel(1).Caption & " (" & frmBubble!UnitsOxygenRef(1) & ")"; Tab(VALUE_TAB); frmBubble!txtOxygen(1).Text
          Print #1, "Method to Find Oxygen KLa"; Tab(VALUE_TAB); frmBubble!cboOxygen.Text
          If frmBubble!cboOxygen.ListIndex = 0 Then   'Method to Find Oxygen KLa is Clean Water Oxygen Transfer Test Data
             Print #1, "Standardized Oxygen Transfer Efficiency, SOTE (%)"; Tab(VALUE_TAB); frmOxygenMassTransferCoeff!txtDataParameters(0).Text; ""
             Print #1, "Standardized Oxygen Transfer Rate (kg O2/d)"; Tab(VALUE_TAB); frmOxygenMassTransferCoeff!txtDataParameters(1).Text
             Print #1, "Air Flow Rate (standard m" & Chr$(179) & "/hr)"; Tab(VALUE_TAB); frmOxygenMassTransferCoeff!txtDataParameters(2).Text
             Print #1, "Barometric Pressure (Pa)"; Tab(VALUE_TAB); frmOxygenMassTransferCoeff!txtDataParameters(3).Text
             Print #1, "Tank Water Depth (m" & Chr$(179) & ")"; Tab(VALUE_TAB); frmOxygenMassTransferCoeff!txtDataParameters(4).Text
             Print #1, "Tank Water Volume (m" & Chr$(179) & ")"; Tab(VALUE_TAB); frmOxygenMassTransferCoeff!txtDataParameters(5).Text
             Print #1, "D.O. Saturation Concentration at Infinite Time (mg/L)"; Tab(VALUE_TAB); frmOxygenMassTransferCoeff!lblDataParameters(6).Caption; ""
             Print #1, frmOxygenMassTransferCoeff!lblDataParametersLabel(7).Caption; Tab(VALUE_TAB); frmOxygenMassTransferCoeff!lblDataParameters(7).Caption
             Print #1, frmOxygenMassTransferCoeff!lblDataParametersLabel(8).Caption; Tab(VALUE_TAB); frmOxygenMassTransferCoeff!lblDataParameters(8).Caption
             Print #1, frmOxygenMassTransferCoeff!lblDataParametersLabel(9).Caption; Tab(VALUE_TAB); frmOxygenMassTransferCoeff!lblDataParameters(9).Caption
             Print #1, frmOxygenMassTransferCoeff!lblDataParametersLabel(10).Caption; Tab(VALUE_TAB); frmOxygenMassTransferCoeff!lblDataParameters(10).Caption
             Print #1, frmOxygenMassTransferCoeff!lblDataParametersLabel(11).Caption; Tab(VALUE_TAB); frmOxygenMassTransferCoeff!lblDataParameters(11).Caption
          End If
          Print #1, "Oxygen " & frmBubble!lblOxygenLabel(2).Caption & " (" & frmBubble!UnitsOxygenRef(2) & ")"; Tab(VALUE_TAB); frmBubble!txtOxygen(2).Text
          Print #1,
          Print #1, frmBubble!lblFlowParametersLabel(0).Caption & " (" & frmBubble!UnitsFlowParam(0) & ")"; Tab(VALUE_TAB); frmBubble!txtFlowParameters(0).Text
          Print #1, frmBubble!lblFlowParametersLabel(2).Caption; Tab(VALUE_TAB); frmBubble!txtFlowParameters(2).Text
          Print #1, frmBubble!lblFlowParametersLabel(3).Caption & " (" & frmBubble!UnitsFlowParam(3) & ")"; Tab(VALUE_TAB); frmBubble!txtFlowParameters(3).Text
          Print #1,
          Print #1, frmBubble!lblTankParametersLabel(0).Caption; Tab(VALUE_TAB); frmBubble!txtTankParameters(0).Text
          Print #1, frmBubble!lblTankParametersLabel(1).Caption & " (" & frmBubble!UnitsTankParam(1) & ")"; Tab(VALUE_TAB); frmBubble!txtTankParameters(1).Text
          Print #1, frmBubble!lblTankParametersLabel(2).Caption & " (" & frmBubble!UnitsTankParam(2) & ")"; Tab(VALUE_TAB); frmBubble!txtTankParameters(2).Text
          Print #1, frmBubble!lblTankParametersLabel(3).Caption & " (" & frmBubble!UnitsTankParam(3) & ")"; Tab(VALUE_TAB); frmBubble!txtTankParameters(3).Text
          Print #1, frmBubble!lblTankParametersLabel(4).Caption & " (" & frmBubble!UnitsTankParam(4) & ")"; Tab(VALUE_TAB); frmBubble!txtTankParameters(4).Text
          Print #1,
          Print #1,
          Call SetPowerBubble
          Print #1, "Power Calculation:"
          Print #1,
          Print #1, "Property:"; Tab(VALUE_TAB); "Value:"
          Print #1,
          Print #1, frmBubblePower!lblPowerLabel(0).Caption; Tab(VALUE_TAB); frmBubblePower!txtPower(0).Text
          Print #1, frmBubblePower!lblPowerLabel(1).Caption; Tab(VALUE_TAB); frmBubblePower!txtPower(1).Text
          Print #1, frmBubblePower!lblPowerLabel(2).Caption; Tab(VALUE_TAB); frmBubblePower!txtPower(2).Text
          Print #1, "Blower " & frmBubblePower!lblPowerLabel(3).Caption; Tab(VALUE_TAB); frmBubblePower!lblPower(3).Caption
          Print #1, frmBubblePower!lblPowerLabel(4).Caption; Tab(VALUE_TAB); frmBubblePower!lblPower(4).Caption
          Print #1, frmBubblePower!lblPowerLabel(5).Caption; Tab(VALUE_TAB); frmBubblePower!txtPower(5).Text
          Print #1, frmBubblePower!lblPowerLabel(6).Caption; Tab(VALUE_TAB); frmBubblePower!lblPower(6).Caption
          Print #1,
          Print #1,
          Print #1, "Contaminant Glossary:"
          Print #1,
          For i = 1 To bub.NumChemical
              Print #1, Format$(i, "0"); " = "; Trim$(bub.Contaminant(i).Name)
          Next i

          Print #1,
          Print #1,
          Print #1,
          Print #1, "Contaminant Properties:"
          Print #1,
          Print #1, "Con.:"; Tab(MWT_TAB); "MWT"; Tab(HC_TAB); "HC"; Tab(VB_TAB); "Vb"; Tab(DIFL_TAB); "DIFL"; Tab(MTCOEFF_TAB); "MT Coeff."; Tab(STANTON_TAB); "St."
          Print #1,
          For i = 1 To bub.NumChemical
              If bub.DesignContaminant.Name = bub.Contaminant(i).Name Then
                 ContaminantMTCoeff(i) = bub.ContaminantMassTransferCoefficient.value
                 StantonNumber(i) = bub.StantonNumber.value
              Else
                 Call KLABUB(ContaminantMTCoeff(i), bub.Oxygen.MassTransferCoefficient.value, bub.Contaminant(i).LiquidDiffusivity.value, bub.Oxygen.LiquidDiffusivity.value, bub.N_for_Finding_KLa.value, bub.kgOVERkl_for_Finding_KLa.value, bub.Contaminant(i).HenrysConstant.value)
                 Call GETPHIB(StantonNumber(i), ContaminantMTCoeff(i), bub.TankVolume.value, bub.Contaminant(i).HenrysConstant.value, bub.AirFlowRate.value)
              End If
              Print #1, Format$(i, "0"); Tab(MWT_TAB); Format$(bub.Contaminant(i).MolecularWeight.value, "0.00"); Tab(HC_TAB); Format$(bub.Contaminant(i).HenrysConstant.value, GetTheFormat(bub.Contaminant(i).HenrysConstant.value)); Tab(VB_TAB); Format$(bub.Contaminant(i).MolarVolume.value, GetTheFormat(bub.Contaminant(i).MolarVolume.value)); Tab(DIFL_TAB); Format$(bub.Contaminant(i).LiquidDiffusivity.value, GetTheFormat(bub.Contaminant(i).LiquidDiffusivity.value)); Tab(MTCOEFF_TAB); Format$(ContaminantMTCoeff(i), GetTheFormat(ContaminantMTCoeff(i))); Tab(STANTON_TAB); Format$(StantonNumber(i), GetTheFormat(StantonNumber(i)))
          Next i
          Print #1,
          Print #1, "Glossary:"
          Print #1, "Con. = Contaminant Number (see Contaminant Glossary on page 1)"
          Print #1, "MWT = Molecular Weight (kg/kmol)"
          Print #1, "HC = Henry's Constant (-)"
          Print #1, "Vb = Molar Volume (m" & Chr$(179) & "/kmol)"
          Print #1, "DIFL = Liquid Diffusivity (m" & Chr$(178) & "/sec)"
          Print #1, "MT Coeff. = Mass Transfer Coeff. (1/sec)"
          Print #1, "St. = Stanton Number (-)"
          Print #1,
          Print #1,
          Print #1,
          Print #1, "Contaminant Concentration Results:"
          Print #1,
          Print #1, "Con.:"; Tab(MWT_TAB); "Cinf"; Tab(HC_TAB); "Cto"; Tab(VB_TAB); "De. % Rem."; Tab(DIFL_TAB); "Ceff"; Tab(MTCOEFF_TAB); "Ach. % Rem."
          Print #1,
          For i = 1 To bub.NumChemical
              If bub.DesignContaminant.Name = bub.Contaminant(i).Name Then
                 DesiredPercentRemoval(i) = bub.DesiredPercentRemoval
                 bub.Contaminant(i).Effluent(0) = bub.DesignContaminant.Effluent(0)
                 For j = 1 To bub.NumberOfTanks.value
                     bub.Contaminant(i).Effluent(j) = bub.DesignContaminant.Effluent(j)
                     bub.Contaminant(i).GasEffluent(j) = bub.DesignContaminant.GasEffluent(j)
                 Next j
                 AchievedPercentRemoval(i) = bub.AchievedPercentRemoval
              Else
                 Call REMOVBUB(DesiredPercentRemoval(i), bub.Contaminant(i).Influent.value, bub.Contaminant(i).TreatmentObjective.value)
                 Call EFFLBUB(Effluent(0), GasEffluent(1), bub.Contaminant(i).HenrysConstant.value, bub.Contaminant(i).Influent.value, bub.AirToWaterRatio.value, bub.NumberOfTanks.value, StantonNumber(i))
                 bub.Contaminant(i).Effluent(0) = Effluent(0)
                 For j = 1 To bub.NumberOfTanks.value
                     bub.Contaminant(i).Effluent(j) = Effluent(j)
                     bub.Contaminant(i).GasEffluent(j) = GasEffluent(j)
                 Next j
                 Call REMOVBUB(AchievedPercentRemoval(i), bub.Contaminant(i).Influent.value, bub.Contaminant(i).Effluent(bub.NumberOfTanks.value))
              End If
              Print #1, Format$(i, "0"); Tab(MWT_TAB); Format$(bub.Contaminant(i).Influent.value, GetTheFormat(bub.Contaminant(i).Influent.value)); Tab(HC_TAB); Format$(bub.Contaminant(i).TreatmentObjective.value, GetTheFormat(bub.Contaminant(i).TreatmentObjective.value)); Tab(VB_TAB); Format$(DesiredPercentRemoval(i), GetTheFormat(DesiredPercentRemoval(i))); Tab(DIFL_TAB); Format$(bub.Contaminant(i).Effluent(bub.NumberOfTanks.value), GetTheFormat(bub.Contaminant(i).Effluent(bub.NumberOfTanks.value))); Tab(MTCOEFF_TAB); Format$(AchievedPercentRemoval(i), GetTheFormat(AchievedPercentRemoval(i)))
          Next i
          Print #1,
          Print #1, "Glossary:"
          Print #1, "Con. = Contaminant Number (see Contaminant Glossary on page 1)"
          Print #1, "Cinf = "; "Liquid Phase " & frmBubble!lblDesignConcentration(0).Caption
          Print #1, "Cto = "; frmBubble!lblDesignConcentration(1).Caption
          Print #1, "De. % Rem. = "; frmBubble!lblDesignConcentration(2).Caption
          Print #1, "Ceff = "; "Liquid Phase Effluent from Last Tank (" & Chr$(181) & "g/L)"
          Print #1, "Ach. % Rem. = "; frmBubble!lblConcentrationResultsLabel(4).Caption
          Print #1,
          Print #1,
          Print #1,
          Print #1, "Liquid Phase Effluent Concentrations from Each Tank in " & Chr$(181) & "g/L:"
          Print #1,
          Print #1,
          Print #1, Tab(MWT_TAB); "Contaminant Number:"
          Print #1,
          Select Case bub.NumChemical
             Case 1
                Print #1, "Tank:"; Tab(MWT_TAB); "1:"
             Case 2
                Print #1, "Tank:"; Tab(MWT_TAB); "1:"; Tab(HC_TAB); "2:"
             Case 3
                Print #1, "Tank:"; Tab(MWT_TAB); "1:"; Tab(HC_TAB); "2:"; Tab(VB_TAB); "3:"
             Case 4
                Print #1, "Tank:"; Tab(MWT_TAB); "1:"; Tab(HC_TAB); "2:"; Tab(VB_TAB); "3:"; Tab(DIFL_TAB); "4:"
             Case 5
                Print #1, "Tank:"; Tab(MWT_TAB); "1:"; Tab(HC_TAB); "2:"; Tab(VB_TAB); "3:"; Tab(DIFL_TAB); "4:"; Tab(MTCOEFF_TAB); "5:"
             Case Else
                Print #1, "Tank:"; Tab(MWT_TAB); "1:"; Tab(HC_TAB); "2:"; Tab(VB_TAB); "3:"; Tab(DIFL_TAB); "4:"; Tab(MTCOEFF_TAB); "5:"; Tab(STANTON_TAB); "6:"
          End Select
          Print #1,

          'Print Liquid Phase Influent Concentrations of Each Contaminant
          Print #1, "Cinf";
          For j = 1 To 6
              If bub.NumChemical < j Then
                 Exit For
              End If
              Select Case j
                 Case 1
                    Print #1, Tab(MWT_TAB);
                 Case 2
                    Print #1, Tab(HC_TAB);
                 Case 3
                    Print #1, Tab(VB_TAB);
                 Case 4
                    Print #1, Tab(DIFL_TAB);
                 Case 5
                    Print #1, Tab(MTCOEFF_TAB);
                 Case 6
                    Print #1, Tab(STANTON_TAB);
              End Select
              Print #1, Format$(bub.Contaminant(j).Influent.value, GetTheFormat(bub.Contaminant(j).Influent.value));
          Next j
          Print #1,

          'Print Liquid Phase Effluent Concentrations from each tank for each contaminant
          For i = 1 To bub.NumberOfTanks.value
              Print #1, Format$(i, "0");
              For j = 1 To 6
                  If bub.NumChemical < j Then
                     Exit For
                  End If
                  Select Case j
                     Case 1
                        Print #1, Tab(MWT_TAB);
                     Case 2
                        Print #1, Tab(HC_TAB);
                     Case 3
                        Print #1, Tab(VB_TAB);
                     Case 4
                        Print #1, Tab(DIFL_TAB);
                     Case 5
                        Print #1, Tab(MTCOEFF_TAB);
                     Case 6
                        Print #1, Tab(STANTON_TAB);
                  End Select
                  Print #1, Format$(bub.Contaminant(j).Effluent(i), GetTheFormat(bub.Contaminant(j).Effluent(i)));
             Next j
             Print #1,
          Next i
          
          If bub.NumChemical < 7 Then
             Print #1,
             Print #1, "Glossary:"
             Print #1, "Cinf = Liquid Phase Influent Concentration to Tank 1 (" & Chr$(181) & "g/L)"
             GoTo AfterLiquidEffluentsFile
          End If
          Print #1,
          Print #1,
          Print #1, Tab(MWT_TAB); "Contaminant Number:"
          Print #1,
          Select Case bub.NumChemical
             Case 7
                Print #1, "Tank:"; Tab(MWT_TAB); "7:"
             Case 8
                Print #1, "Tank:"; Tab(MWT_TAB); "7:"; Tab(HC_TAB); "8:"
             Case 9
                Print #1, "Tank:"; Tab(MWT_TAB); "7:"; Tab(HC_TAB); "8:"; Tab(VB_TAB); "9:"
             Case 10
                Print #1, "Tank:"; Tab(MWT_TAB); "7:"; Tab(HC_TAB); "8:"; Tab(VB_TAB); "9:"; Tab(DIFL_TAB); "10:"
          End Select
          Print #1,

          'Print Liquid Phase Influent Concentrations of Each Contaminant
          Print #1, "Cinf";
          For j = 7 To 10
              If bub.NumChemical < j Then
                 Exit For
              End If
              Select Case j
                 Case 7
                    Print #1, Tab(MWT_TAB);
                 Case 8
                    Print #1, Tab(HC_TAB);
                 Case 9
                    Print #1, Tab(VB_TAB);
                 Case 10
                    Print #1, Tab(DIFL_TAB);
              End Select
              Print #1, Format$(bub.Contaminant(j).Influent.value, GetTheFormat(bub.Contaminant(j).Influent.value));
          Next j
          Print #1,

          'Print Liquid Phase Effluent Concentrations from each tank for each contaminant
          For i = 1 To bub.NumberOfTanks.value
              Print #1, Format$(i, "0");
              For j = 7 To 10
                  If bub.NumChemical < j Then
                     Exit For
                  End If
                  Select Case j
                     Case 7
                        Print #1, Tab(MWT_TAB);
                     Case 8
                        Print #1, Tab(HC_TAB);
                     Case 9
                        Print #1, Tab(VB_TAB);
                     Case 10
                        Print #1, Tab(DIFL_TAB);
                  End Select
                  Print #1, Format$(bub.Contaminant(j).Effluent(i), GetTheFormat(bub.Contaminant(j).Effluent(i)));
             Next j
             Print #1,
          Next i
          
          Print #1,
          Print #1, "Glossary:"
          Print #1, "Cinf = Liquid Phase Influent Concentration to Tank 1 (" & Chr$(181) & "g/L)"

AfterLiquidEffluentsFile:

             Print #1,
             Print #1,
             Print #1,
             Print #1,
          Print #1, "Gas Phase Effluent Concentrations from Each Tank in " & Chr$(181) & "g/L:"
          Print #1,
          Print #1,
          Print #1, Tab(MWT_TAB); "Contaminant Number:"
          Print #1,
          
          Select Case bub.NumChemical
             Case 1
                Print #1, "Tank:"; Tab(MWT_TAB); "1:"
             Case 2
                Print #1, "Tank:"; Tab(MWT_TAB); "1:"; Tab(HC_TAB); "2:"
             Case 3
                Print #1, "Tank:"; Tab(MWT_TAB); "1:"; Tab(HC_TAB); "2:"; Tab(VB_TAB); "3:"
             Case 4
                Print #1, "Tank:"; Tab(MWT_TAB); "1:"; Tab(HC_TAB); "2:"; Tab(VB_TAB); "3:"; Tab(DIFL_TAB); "4:"
             Case 5
                Print #1, "Tank:"; Tab(MWT_TAB); "1:"; Tab(HC_TAB); "2:"; Tab(VB_TAB); "3:"; Tab(DIFL_TAB); "4:"; Tab(MTCOEFF_TAB); "5:"
             Case Else
                Print #1, "Tank:"; Tab(MWT_TAB); "1:"; Tab(HC_TAB); "2:"; Tab(VB_TAB); "3:"; Tab(DIFL_TAB); "4:"; Tab(MTCOEFF_TAB); "5:"; Tab(STANTON_TAB); "6:"
          End Select
          Print #1,

          'Print Gas Phase Influent Concentrations of Each Contaminant
          Print #1, "Yinf";
          For j = 1 To 6
              If bub.NumChemical < j Then
                 Exit For
              End If
              Select Case j
                 Case 1
                    Print #1, Tab(MWT_TAB);
                 Case 2
                    Print #1, Tab(HC_TAB);
                 Case 3
                    Print #1, Tab(VB_TAB);
                 Case 4
                    Print #1, Tab(DIFL_TAB);
                 Case 5
                    Print #1, Tab(MTCOEFF_TAB);
                 Case 6
                    Print #1, Tab(STANTON_TAB);
              End Select
              Print #1, "0";
          Next j
          Print #1,

          'Print Gas Phase Effluent Concentrations from each tank for each contaminant
          For i = 1 To bub.NumberOfTanks.value
              Print #1, Format$(i, "0");
              For j = 1 To 6
                  If bub.NumChemical < j Then
                     Exit For
                  End If
                  Select Case j
                     Case 1
                        Print #1, Tab(MWT_TAB);
                     Case 2
                        Print #1, Tab(HC_TAB);
                     Case 3
                        Print #1, Tab(VB_TAB);
                     Case 4
                        Print #1, Tab(DIFL_TAB);
                     Case 5
                        Print #1, Tab(MTCOEFF_TAB);
                     Case 6
                        Print #1, Tab(STANTON_TAB);
                  End Select
                  Print #1, Format$(bub.Contaminant(j).GasEffluent(i), GetTheFormat(bub.Contaminant(j).GasEffluent(i)));
             Next j
             Print #1,
          Next i
          
          If bub.NumChemical < 7 Then
             Print #1,
             Print #1, "Glossary:"
             Print #1, "Yinf = Gas Phase Influent Concentration to All Tanks (" & Chr$(181) & "g/L)"
             GoTo AfterGasEffluentsFile
          End If
          Print #1,
          Print #1,
          Print #1, Tab(MWT_TAB); "Contaminant Number:"
          Print #1,
       
          Select Case bub.NumChemical
             Case 7
                Print #1, "Tank:"; Tab(MWT_TAB); "7:"
             Case 8
                Print #1, "Tank:"; Tab(MWT_TAB); "7:"; Tab(HC_TAB); "8:"
             Case 9
                Print #1, "Tank:"; Tab(MWT_TAB); "7:"; Tab(HC_TAB); "8:"; Tab(VB_TAB); "9:"
             Case 10
                Print #1, "Tank:"; Tab(MWT_TAB); "7:"; Tab(HC_TAB); "8:"; Tab(VB_TAB); "9:"; Tab(DIFL_TAB); "10:"
          End Select
          Print #1,

          'Print Gas Phase Influent Concentrations of Each Contaminant
          Print #1, "Yinf";
          For j = 7 To 10
              If bub.NumChemical < j Then
                 Exit For
              End If
              Select Case j
                 Case 7
                    Print #1, Tab(MWT_TAB);
                 Case 8
                    Print #1, Tab(HC_TAB);
                 Case 9
                    Print #1, Tab(VB_TAB);
                 Case 10
                    Print #1, Tab(DIFL_TAB);
              End Select
              Print #1, "0";
          Next j
          Print #1,

          'Print Gas Phase Effluent Concentrations from each tank for each contaminant
          For i = 1 To bub.NumberOfTanks.value
              Print #1, Format$(i, "0");
              For j = 7 To 10
                  If bub.NumChemical < j Then
                     Exit For
                  End If
                  Select Case j
                     Case 7
                        Print #1, Tab(MWT_TAB);
                     Case 8
                        Print #1, Tab(HC_TAB);
                     Case 9
                        Print #1, Tab(VB_TAB);
                     Case 10
                        Print #1, Tab(DIFL_TAB);
                  End Select
                  Print #1, Format$(bub.Contaminant(j).GasEffluent(i), GetTheFormat(bub.Contaminant(j).GasEffluent(i)));
             Next j
             Print #1,
          Next i

          Print #1,
          Print #1, "Glossary:"
          Print #1, "Yinf = Gas Phase Influent Concentration to All Tanks (" & Chr$(181) & "g/L)"
          

AfterGasEffluentsFile:

    End Select

    Close #1

End Sub

Sub savebubble()
Dim FileID As String
Dim i As Integer
Dim xu As rec_Units_frmContaminantPropertyEdit
Dim TransferTestDummy As Integer

  If (IsThisADemo() = True) Then
    Call Demo_ShowError("Saving is not allowed in the demonstration version.")
    Exit Sub
  End If
    
    If Right$(frmBubble.Caption, 14) = "(untitled.bub)" Then
       Call savefilebubble(Filename)
    End If

    If Filename$ <> "" Then
       FileID = BUBBLE_FILEID
       Open Filename$ For Output As #1
       
       Write #1, FileID
      
       Write #1, BubbleAerationMode, ""

       Write #1, bub.OperatingPressure.value, ""
       Write #1, bub.operatingtemperature.value, ""

       
       If bub.Oxygen.KLaMethod = KLA_METHOD_USER_INPUT Then
          Write #1, bub.Oxygen.KLaMethod, ""
          Write #1, bub.Oxygen.MassTransferCoefficient.value, ""

       ElseIf bub.Oxygen.KLaMethod = KLA_METHOD_CWO2_TRANSFER_TEST Then
          Write #1, bub.Oxygen.KLaMethod, ""
          If frmOxygenMassTransferCoeff!optDataAvailable(0).value = True Then
             TransferTestDummy = 1
          Else
             TransferTestDummy = 2
          End If
          
          If TransferTestDummy = 1 Then       'SOTR vs. QAIR data available
             Write #1, TransferTestDummy, ""
             Write #1, bub.Oxygen.CWO2TestData.SOTR.value, ""
             Write #1, bub.Oxygen.CWO2TestData.AirFlowRate_QAIR.value, ""
             Write #1, bub.Oxygen.CWO2TestData.BarometricPressure_PB.value, ""
             Write #1, bub.Oxygen.CWO2TestData.WaterDepth_DEPTHW.value, ""
             Write #1, bub.Oxygen.CWO2TestData.WaterVolumePerTank_VM3.value, ""

          ElseIf TransferTestDummy = 2 Then   'SOTE vs. QAIR data available
             Write #1, TransferTestDummy, ""
             Write #1, bub.Oxygen.CWO2TestData.SOTE.value, ""
             Write #1, bub.Oxygen.CWO2TestData.AirFlowRate_QAIR.value, ""
             Write #1, bub.Oxygen.CWO2TestData.BarometricPressure_PB.value, ""
             Write #1, bub.Oxygen.CWO2TestData.WaterDepth_DEPTHW.value, ""
             Write #1, bub.Oxygen.CWO2TestData.WaterVolumePerTank_VM3.value, ""
                       
          End If
       End If

       Write #1, bub.NumChemical, ""
       For i = 1 To bub.NumChemical
           Write #1, bub.Contaminant(i).Pressure, bub.Contaminant(i).Temperature, bub.Contaminant(i).Name, bub.Contaminant(i).MolecularWeight.value, bub.Contaminant(i).HenrysConstant.value, bub.Contaminant(i).MolarVolume.value, bub.Contaminant(i).LiquidDiffusivity.value, bub.Contaminant(i).Influent.value, bub.Contaminant(i).TreatmentObjective.value
       Next i
       Write #1, bub.DesignContaminant.Name, ""

       Write #1, bub.WaterFlowRate.value, ""
       Write #1, bub.AirToWaterRatio.value, bub.AirToWaterRatio.UserInput, ""
       Write #1, bub.AirFlowRate.value, bub.AirFlowRate.UserInput, ""
       Write #1, bub.NumberOfTanks.value, ""

       Write #1, bub.CodeForTausAndTankVolumes, ""
       Select Case bub.CodeForTausAndTankVolumes
          Case 1   'Write Hydraulic Retention Time for 1 Tank
             Write #1, bub.TankHydraulicRetentionTime.value, ""
          Case 2   'Write Hydraulic Retention Time for All Tanks
             Write #1, bub.TotalHydraulicRetentionTime.value, ""
          Case 3   'Write Volume of Each Tank
             Write #1, bub.TankVolume.value, ""
          Case 4   'Write Volume of All Tanks
             Write #1, bub.TotalTankVolume.value, ""
       End Select

       Write #1, bub.Power.BlowerEfficiency, ""
       Write #1, bub.Power.TankWaterDepth, ""
       Write #1, bub.Power.NumberOfBlowersinEachTank, ""
       
       'Output the units of this screen.
       Write #1, GetUnits(frmBubble!UnitsOpCond(0)), GetUnits(frmBubble!UnitsOpCond(1))
       Write #1, GetUnits(frmBubble!UnitsOxygenRef(1)), GetUnits(frmBubble!UnitsOxygenRef(2))
       Write #1, GetUnits(frmBubble!UnitsDesignContam(0)), GetUnits(frmBubble!UnitsDesignContam(1)), GetUnits(frmBubble!UnitsDesignContam(3))
       Write #1, GetUnits(frmBubble!UnitsFlowParam(0)), GetUnits(frmBubble!UnitsFlowParam(3))
       Write #1, GetUnits(frmBubble!UnitsTankParam(1)), GetUnits(frmBubble!UnitsTankParam(2)), GetUnits(frmBubble!UnitsTankParam(3)), GetUnits(frmBubble!UnitsTankParam(4))
       Write #1, GetUnits(frmBubble!UnitsConcResults(1)), GetUnits(frmBubble!UnitsConcResults(2)), GetUnits(frmBubble!UnitsConcResults(3))

       'Output the units of frmContaminantPropertyEdit.
       xu = Units_frmContaminantPropertyEdit
       Write #1, xu.UnitsProp(0), xu.UnitsProp(2), xu.UnitsProp(3), xu.UnitsProp(4), xu.UnitsProp(5)
       Write #1, xu.UnitsConc(0), xu.UnitsConc(1)

       Close #1

       If BubbleAerationMode = DESIGN_MODE Then
          frmBubble.Caption = "Bubble Aeration - Design Mode"
       Else
          frmBubble.Caption = "Bubble Aeration - Rating Mode"
       End If

       frmBubble.Caption = frmBubble.Caption & " (" & Filename & ")"

    End If

End Sub

Sub SaveContaminantListBubble()
    Dim FileID As String
    Dim i As Integer

    Call SaveFile(Filename)

    If Filename$ <> "" Then
       FileID = CONTAMINANTS_BUBBLE_FILEID
       Open Filename$ For Output As #1
       
       Write #1, FileID
      
       For i = 1 To bub.NumChemical
           Write #1, bub.Contaminant(i).Pressure, bub.Contaminant(i).Temperature, bub.Contaminant(i).Name, bub.Contaminant(i).MolecularWeight.value, bub.Contaminant(i).HenrysConstant.value, bub.Contaminant(i).MolarVolume.value, bub.Contaminant(i).LiquidDiffusivity.value, bub.Contaminant(i).Influent.value, bub.Contaminant(i).TreatmentObjective.value
       Next i

       Close #1

    End If

End Sub

Sub savefilebubble(Filename As String)
Dim Ctl As Control
Set Ctl = frmBubble.CommonDialog1

    On Error Resume Next
    'frmBubble!CMDialog1.DefaultExt = "bub"
    'frmBubble!CMDialog1.Filter = "Bubble Aeration Files (*.bub)|*.bub"
    'frmBubble!CMDialog1.DialogTitle = "Save Bubble Aeration File"
    'frmBubble!CMDialog1.flags = OFN_OVERWRITEPROMPT Or OFN_PATHMUSTEXIST
    'frmBubble!CMDialog1.Action = 2
    'Filename$ = frmBubble!CMDialog1.Filename
    Ctl.DefaultExt = "bub"
    Ctl.Filter = "Bubble Aeration Files (*.bub)|*.bub"
    Ctl.DialogTitle = "Save Bubble Aeration File"
    Ctl.flags = OFN_OVERWRITEPROMPT Or OFN_PATHMUSTEXIST
    Ctl.Action = 2
    Filename$ = Ctl.Filename
    If Err = 32755 Then   'Cancel selected by user
       Filename$ = ""
    End If

End Sub

Sub SetDesignContaminantEnabledBubble(NumInList As Integer)
    Dim i As Integer

    If NumInList = 0 Then
       frmBubble!mnuFile(4).Enabled = False
       frmBubble!mnuFile(5).Enabled = False
       frmBubble!mnuOptions(0).Enabled = False
       'frmBubble!fraDesignContaminant.Enabled = False
       frmBubble!cboDesignContaminant.Enabled = False
       frmBubble!fraFlowParameters.Enabled = False
       frmBubble!fraTankParameters.Enabled = False
       frmBubble!fraConcentrationResults.Enabled = False
       frmBubble!mnuPower(0).Enabled = False
       For i = 0 To 2
           frmBubble!lblDesignConcentrationValue(i).Caption = ""
       Next i
       frmBubble!txtDesignConcentrationValue(3).Text = ""
       frmBubble!lblFlowParameters(1).Caption = ""
       frmBubble!lblStanton.Caption = ""
       If BubbleAerationMode = DESIGN_MODE Then
          frmBubble.txtTankParameters(1).Text = ""
          frmBubble.txtTankParameters(2).Text = ""
          frmBubble.txtTankParameters(3).Text = ""
          frmBubble.txtTankParameters(4).Text = ""
       End If

       frmBubble!lblConcentrationResults(0).Caption = ""
       frmBubble!lblConcentrationResults(1).Caption = ""
       frmBubble!lblConcentrationResults(3).Caption = ""
       frmBubble!lblConcentrationResults(4).Caption = ""

    Else
     
       frmBubble!mnuFile(4).Enabled = True
       frmBubble!mnuFile(5).Enabled = True
     
       frmBubble!mnuOptions(0).Enabled = True
       frmBubble!mnuPower(0).Enabled = True

       'frmBubble!fraDesignContaminant.Enabled = True
       frmBubble!cboDesignContaminant.Enabled = True
       frmBubble!fraTankParameters.Enabled = True
       frmBubble!fraFlowParameters.Enabled = True
       frmBubble!fraConcentrationResults.Enabled = True

    End If
    Call frmBubble.LOCAL___Reset_DemoVersionDisablings
End Sub

Sub SetPowerBubble()
          bub.Power.InletAirTemperature = bub.operatingtemperature.value - 273.15
          Call CalculatePowerBubble
          
             frmBubblePower!txtPower(0).Text = Format$(bub.Power.InletAirTemperature, GetTheFormat(bub.Power.InletAirTemperature))
             frmBubblePower!txtPower(1).Text = Format$(bub.Power.BlowerEfficiency, GetTheFormat(bub.Power.BlowerEfficiency))
             frmBubblePower!txtPower(2).Text = Format$(bub.Power.TankWaterDepth, GetTheFormat(bub.Power.TankWaterDepth))
             frmBubblePower!lblPower(3).Caption = Format$(bub.Power.BlowerBrakePower, GetTheFormat(bub.Power.BlowerBrakePower))
             frmBubblePower!lblPower(4).Caption = Format$(bub.NumberOfTanks.value, "0")
             frmBubblePower!txtPower(5).Text = Format$(bub.Power.NumberOfBlowersinEachTank, "0")
             frmBubblePower!lblPower(6).Caption = Format$(bub.Power.TotalBrakePower, GetTheFormat(bub.Power.TotalBrakePower))

End Sub

Function StartBubbleDefaultCase() As Boolean

    Filename = "TheDefaultCaseBubble"
    StartBubbleDefaultCase = loadbubble("")

End Function

