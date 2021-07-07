Attribute VB_Name = "SurfaceMod"
Option Explicit
Global Const KLA_METHOD_SURFACE_ROBERTS_CORRELATION = 1
Global Const KLA_METHOD_SURFACE_USER_INPUT = 2
Global Const SURFACE_FILEID = "Properties_Surface_Aeration"
Global Const CONTAMINANTS_SURFACE_FILEID = "Contaminants_Surface_Aeration"

Global SurfaceAerationMode As Integer

Type SurfaceInformationType
     value As Double
     UserInput As Integer
     ValChanged As Integer
End Type

Type SurfaceInformationType2
     value As Long
     UserInput As Integer
     ValChanged As Integer
End Type

Type OxygenInformationType_Surface
     LiquidDiffusivity As SurfaceInformationType
     KLaMethod As Integer
     MassTransferCoefficient As SurfaceInformationType
End Type

Type SurfaceContaminantPropertyType
     Pressure As Double
     Temperature As Double
     Name As String
     MolecularWeight As SurfaceInformationType
     HenrysConstant As SurfaceInformationType
     MolarVolume As SurfaceInformationType
     LiquidDiffusivity As SurfaceInformationType
     Influent As SurfaceInformationType
     TreatmentObjective As SurfaceInformationType
     Effluent(0 To MAXIMUM_TANKS) As Double
End Type

Type PowerTypeSurface
     AeratorMotorEfficiency As Double
     PowerForEachTank As Double
     TotalPowerForAllTanks As Double
End Type

Type SurfaceType
     OperatingPressure As SurfaceInformationType
     operatingtemperature As SurfaceInformationType
     WaterDensity As SurfaceInformationType
     WaterViscosity As SurfaceInformationType
     PowerInput_PoverV As SurfaceInformationType
     N_for_Finding_KLa As SurfaceInformationType
     kgOVERkl_for_Finding_KLa As SurfaceInformationType
     ContaminantMassTransferCoefficient As SurfaceInformationType
     WaterFlowRate As SurfaceInformationType
     TankHydraulicRetentionTime As SurfaceInformationType
     TotalHydraulicRetentionTime As SurfaceInformationType
     TankVolume As SurfaceInformationType
     TotalTankVolume As SurfaceInformationType
     
     NumberOfTanks As SurfaceInformationType2
     
     CodeForTausAndTankVolumes As Long
     DesiredPercentRemoval As Double
     AchievedPercentRemoval As Double
     
     Power As PowerTypeSurface
     
     Oxygen As OxygenInformationType_Surface
     
     NumChemical As Integer
     Chemical As Integer
     Contaminant(1 To MAXCHEMICAL) As SurfaceContaminantPropertyType
     DesignContaminant As SurfaceContaminantPropertyType
End Type

Global sur As SurfaceType

Global ErrorFlagSur As Long   'Error Flag passed to Sub VOLBUB

Sub CalculateOxygenMTCoeffSurface()

  Call KLAO2SUR(sur.Oxygen.MassTransferCoefficient.value, sur.PowerInput_PoverV.value)
  'frmSurface!txtOxygen(2).Text = Format$(sur.Oxygen.MassTransferCoefficient.Value, GetTheFormat(sur.Oxygen.MassTransferCoefficient.Value))
  Call Unitted_NumberUpdate(frmSurface!UnitsOxygenRef(2))
  sur.Oxygen.MassTransferCoefficient.UserInput = False

End Sub

Sub CalculatePowerSurface()
Dim Dummy As Double

  Call PCALCSUR(sur.Power.TotalPowerForAllTanks, sur.Power.PowerForEachTank, sur.PowerInput_PoverV.value, sur.TotalTankVolume.value, sur.NumberOfTanks.value, sur.Power.AeratorMotorEfficiency)
  
  'UPDATED_UNITS.
  'Update Power Calculation | Power Required per Tank.
  'frmSurface!lblPowerCalculation(1).Caption = Format$(sur.Power.PowerForEachTank, GetTheFormat(sur.Power.PowerForEachTank))
  Call Unitted_NumberUpdate(frmSurface!UnitsPowerCalc(1))
 
  'UPDATED_UNITS.
  'Update Power Calculation | Total Power Required.
  'frmSurface!lblPowerCalculation(2).Caption = Format$(sur.Power.TotalPowerForAllTanks, GetTheFormat(sur.Power.TotalPowerForAllTanks))
  Call Unitted_NumberUpdate(frmSurface!UnitsPowerCalc(2))

End Sub

Sub CalculateRetentionTimeSurface()

  Call TAUISURF(sur.TankHydraulicRetentionTime.value, sur.DesignContaminant.Influent.value, sur.DesignContaminant.TreatmentObjective.value, sur.NumberOfTanks.value, sur.ContaminantMassTransferCoefficient.value)
  'frmSurface!txtTankParameters(1).Text = Format$(sur.TankHydraulicRetentionTime.Value, GetTheFormat(sur.TankHydraulicRetentionTime.Value))
  Call Unitted_NumberUpdate(frmSurface!UnitsTankParam(1))
  sur.TankHydraulicRetentionTime.UserInput = False

End Sub

Sub CalculateTausAndTankVolumesSurface()

  Call TAUSVOLS(sur.TotalHydraulicRetentionTime.value, sur.NumberOfTanks.value, sur.TankHydraulicRetentionTime.value, sur.TankVolume.value, sur.TotalTankVolume.value, sur.WaterFlowRate.value, sur.CodeForTausAndTankVolumes)

  Select Case sur.CodeForTausAndTankVolumes
    Case 1   'Input Fluid Residence Time of Each Tank
      'frmSurface!txtTankParameters(2).Text = Format$(sur.TotalHydraulicRetentionTime.Value, GetTheFormat(sur.TotalHydraulicRetentionTime.Value))
      'frmSurface!txtTankParameters(3).Text = Format$(sur.TankVolume.Value, GetTheFormat(sur.TankVolume.Value))
      'frmSurface!txtTankParameters(4).Text = Format$(sur.TotalTankVolume.Value, GetTheFormat(sur.TotalTankVolume.Value))
      sur.TotalHydraulicRetentionTime.UserInput = False
      sur.TankVolume.UserInput = False
      sur.TotalTankVolume.UserInput = False
      Call Unitted_NumberUpdate(frmSurface!UnitsTankParam(2))
      Call Unitted_NumberUpdate(frmSurface!UnitsTankParam(3))
      Call Unitted_NumberUpdate(frmSurface!UnitsTankParam(4))
    Case 2   'Input Total Fluid Residence Time
      'frmSurface!txtTankParameters(1).Text = Format$(sur.TankHydraulicRetentionTime.Value, GetTheFormat(sur.TankHydraulicRetentionTime.Value))
      'frmSurface!txtTankParameters(3).Text = Format$(sur.TankVolume.Value, GetTheFormat(sur.TankVolume.Value))
      'frmSurface!txtTankParameters(4).Text = Format$(sur.TotalTankVolume.Value, GetTheFormat(sur.TotalTankVolume.Value))
      sur.TankHydraulicRetentionTime.UserInput = False
      sur.TankVolume.UserInput = False
      sur.TotalTankVolume.UserInput = False
      Call Unitted_NumberUpdate(frmSurface!UnitsTankParam(1))
      Call Unitted_NumberUpdate(frmSurface!UnitsTankParam(3))
      Call Unitted_NumberUpdate(frmSurface!UnitsTankParam(4))
    Case 3   'Input Volume of Each Tank
      'frmSurface!txtTankParameters(1).Text = Format$(sur.TankHydraulicRetentionTime.Value, GetTheFormat(sur.TankHydraulicRetentionTime.Value))
      'frmSurface!txtTankParameters(2).Text = Format$(sur.TotalHydraulicRetentionTime.Value, GetTheFormat(sur.TotalHydraulicRetentionTime.Value))
      'frmSurface!txtTankParameters(4).Text = Format$(sur.TotalTankVolume.Value, GetTheFormat(sur.TotalTankVolume.Value))
      sur.TankHydraulicRetentionTime.UserInput = False
      sur.TotalHydraulicRetentionTime.UserInput = False
      sur.TotalTankVolume.UserInput = False
      Call Unitted_NumberUpdate(frmSurface!UnitsTankParam(1))
      Call Unitted_NumberUpdate(frmSurface!UnitsTankParam(2))
      Call Unitted_NumberUpdate(frmSurface!UnitsTankParam(4))
    Case 4   'Input Total Volume of All Tanks
      'frmSurface!txtTankParameters(1).Text = Format$(sur.TankHydraulicRetentionTime.Value, GetTheFormat(sur.TankHydraulicRetentionTime.Value))
      'frmSurface!txtTankParameters(2).Text = Format$(sur.TotalHydraulicRetentionTime.Value, GetTheFormat(sur.TotalHydraulicRetentionTime.Value))
      'frmSurface!txtTankParameters(3).Text = Format$(sur.TankVolume.Value, GetTheFormat(sur.TankVolume.Value))
      sur.TankHydraulicRetentionTime.UserInput = False
      sur.TotalHydraulicRetentionTime.UserInput = False
      sur.TankVolume.UserInput = False
      Call Unitted_NumberUpdate(frmSurface!UnitsTankParam(1))
      Call Unitted_NumberUpdate(frmSurface!UnitsTankParam(2))
      Call Unitted_NumberUpdate(frmSurface!UnitsTankParam(3))
  End Select

End Sub

Sub CalculateWaterPropertiesSurface()
    Dim Pressure As Double
    Dim Temperature As Double
    Dim WaterDensity As Double
    Dim WaterViscosity As Double
    Dim i As Integer
    
    
       Pressure = sur.OperatingPressure.value
       Temperature = sur.operatingtemperature.value

       For i = 0 To 1
           If frmWaterPropertiesSurface!chkUpdateValues(i).value = True Then
              Select Case i
                 Case 0
                    If HaveValue(Temperature) Then
                       Call H2ODENS(WaterDensity, Temperature)
                       sur.WaterDensity.value = WaterDensity
                       sur.WaterDensity.UserInput = False
                       sur.WaterDensity.ValChanged = True
                       frmWaterPropertiesSurface.txtAirWaterProperties(0).Text = Format$(WaterDensity, "###0.00")
                       frmWaterPropertiesSurface.lblValueSource(0).Caption = "Correlation"
                    End If
                 Case 1
                    If HaveValue(Temperature) Then
                       Call H2OVISC(WaterViscosity, Temperature)
                       sur.WaterViscosity.value = WaterViscosity
                       sur.WaterViscosity.UserInput = False
                       sur.WaterViscosity.ValChanged = True
                       frmWaterPropertiesSurface.txtAirWaterProperties(1).Text = Format$(WaterViscosity, "0.000E+##")
                       frmWaterPropertiesSurface.lblValueSource(1).Caption = "Correlation"
                    End If
              End Select
          End If
       Next i
    

End Sub

Sub InitializeOxygenMTCoeff_Surface()

    frmSurface!cboOxygen.ListIndex = 1   'User input
    sur.Oxygen.KLaMethod = KLA_METHOD_SURFACE_USER_INPUT
    sur.Oxygen.MassTransferCoefficient.value = 0.0046
    frmSurface!txtOxygen(2).Text = "0.0046"

End Sub

Sub InitializePressureTemperatureSurface()
    
  '*****************************************************
  '*                                                   *
  '* Initialize Pressure and Temperature to defaults:  *
  '*                                                   *
  '*  Operating Pressure = 1 atm                       *
  '*  Operating Temperature = 10.0 C                   *
  '*                                                   *
  '*****************************************************

  sur.OperatingPressure.value = 1#
  sur.OperatingPressure.ValChanged = True
  sur.operatingtemperature.value = 293.15
  sur.operatingtemperature.ValChanged = True

  frmSurface.txtOperatingPressure.Text = "101325.0"
  frmSurface.txtOperatingTemperature.Text = "20.00"

  Call CalculateWaterPropertiesSurface
  Call CalculateOxygenLiquidDiffSurface

End Sub

Sub LoadContaminantListSurface()
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

       'frmListContaminantSurface.ListContaminants.Clear
       frmSurface!cboDesignContaminant.Clear

       i = 0
       NotSpecifiedAtOperatingTemperature = False
       NotSpecifiedAtOperatingPressure = False
       Do Until EOF(1)
          i = i + 1
          If FileID = CONTAMINANTS_BUBBLE_FILEID Or FileID = CONTAMINANTS_SURFACE_FILEID Then
             Input #1, sur.Contaminant(i).Pressure, sur.Contaminant(i).Temperature, sur.Contaminant(i).Name, sur.Contaminant(i).MolecularWeight.value, sur.Contaminant(i).HenrysConstant.value, sur.Contaminant(i).MolarVolume.value, sur.Contaminant(i).LiquidDiffusivity.value, sur.Contaminant(i).Influent.value, sur.Contaminant(i).TreatmentObjective.value
          Else
             Input #1, sur.Contaminant(i).Pressure, sur.Contaminant(i).Temperature, sur.Contaminant(i).Name, sur.Contaminant(i).MolecularWeight.value, sur.Contaminant(i).HenrysConstant.value, sur.Contaminant(i).MolarVolume.value, NormalBoilingPoint, sur.Contaminant(i).LiquidDiffusivity.value, GasDiffusivity, sur.Contaminant(i).Influent.value, sur.Contaminant(i).TreatmentObjective.value
          End If
          'frmListContaminantSurface.ListContaminants.AddItem sur.Contaminant(i).Name
          frmSurface!cboDesignContaminant.AddItem sur.Contaminant(i).Name

          If Not NotSpecifiedAtOperatingTemperature Then
             If Abs(sur.Contaminant(i).Temperature - sur.operatingtemperature.value) > TOLERANCE Then
                NotSpecifiedAtOperatingTemperature = True
             End If
          End If
          If Not NotSpecifiedAtOperatingPressure Then
             If Abs(sur.Contaminant(i).Pressure - sur.OperatingPressure.value) > TOLERANCE Then
                NotSpecifiedAtOperatingPressure = True
             End If
          End If

       Loop
       sur.NumChemical = i
          
       Close #1

       'If frmListContaminantSurface.mnuOptionsManipulateContaminant(1).Enabled = False Then
       '   frmListContaminantSurface.mnuOptionsManipulateContaminant(1).Enabled = True
       '   frmListContaminantSurface.mnuOptionsManipulateContaminant(3).Enabled = True
       '   frmListContaminantSurface.mnuOptionsManipulateContaminant(4).Enabled = True
       '   frmListContaminantSurface.mnuOptionsSave.Enabled = True
       '   frmListContaminantSurface.mnuOptionsView.Enabled = True
       'End If

       'frmListContaminantSurface.ListContaminants.Selected(0) = True

       If NotSpecifiedAtOperatingPressure And NotSpecifiedAtOperatingTemperature Then
          MsgBox "For one or more contaminants, the temperature and pressure at which the contaminant properties are specified differs from the operating temperature and pressure.", MB_ICONINFORMATION, "Warning"
       ElseIf NotSpecifiedAtOperatingTemperature Then
          MsgBox "For one or more contaminants, the temperature at which the contaminant properties are specified differs from the operating temperature.", MB_ICONINFORMATION, "Warning"
       ElseIf NotSpecifiedAtOperatingPressure Then
          MsgBox "For one or more contaminants, the pressure at which the contaminant properties are specified differs from the operating pressure.", MB_ICONINFORMATION, "Warning"
       End If

    End If

End Sub

Sub LoadFileSurface(Filename As String)
Dim Ctl As Control
Set Ctl = frmSurface.CommonDialog1

    On Error Resume Next
    'frmSurface!CMDialog1.DefaultExt = "sur"
    'frmSurface!CMDialog1.Filter = "Surface Aeration Files (*.sur)|*.sur"
    'frmSurface!CMDialog1.DialogTitle = "Load Surface Aeration File"
    'frmSurface!CMDialog1.flags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST
    'frmSurface!CMDialog1.Action = 1
    'Filename$ = frmSurface!CMDialog1.Filename
    Ctl.DefaultExt = "sur"
    Ctl.Filter = "Surface Aeration Files (*.sur)|*.sur"
    Ctl.DialogTitle = "Load Surface Aeration File"
    Ctl.flags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST
    Ctl.Action = 1
    Filename$ = Ctl.Filename
    If Err = 32755 Then   'Cancel selected by user
       Filename$ = ""
    End If

End Sub

Function loadsurface(OverrideFilename As String) As Boolean
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
      If Filename = "TheDefaultCaseSurface" Then
        If SurfaceAerationMode = DESIGN_MODE Then
          Filename = App.Path & "\dbase\defltdes.sur"
        Else
          Filename = App.Path & "\dbase\defltrat.sur"
        End If
      Else
        Call LoadFileSurface(Filename)
      End If
    End If
    
    If Filename$ <> "" Then
       FileID = ""
       If (fileexists(Filename) = False) Then
         Call Error_Unavailable_File( _
            Filename, _
            IIf(SurfaceAerationMode = DESIGN_MODE, _
                "Surface Aeration Design Mode", _
                "Surface Aeration Rating Mode"))
         loadsurface = False
         Exit Function
       End If
       Open Filename$ For Input As #1
       On Error Resume Next
       Input #1, FileID
       If FileID <> SURFACE_FILEID Then
          msg = "Invalid Design File"
          MsgBox msg, 48, "Error"
          Close #1
          Exit Function
       End If

       'frmListContaminantSurface.ListContaminants.Clear
       frmSurface!cboDesignContaminant.Clear

       Input #1, SurfaceAerationMode, CommentDummy
       If SurfaceAerationMode = DESIGN_MODE Then
          frmSurface.Caption = "Surface Aeration - Design Mode"
          frmSurface!mnuFile(0).Caption = "Switch to &Rating Mode"
       ElseIf SurfaceAerationMode = RATING_MODE Then
          frmSurface.Caption = "Surface Aeration - Rating Mode"
          frmSurface!mnuFile(0).Caption = "Switch to &Design Mode"
       End If

       Input #1, sur.OperatingPressure.value, CommentDummy
       frmSurface!txtOperatingPressure.Text = Format$(sur.OperatingPressure.value * 101325# / 1#, "0.00")

       Input #1, sur.operatingtemperature.value, CommentDummy
       frmSurface!txtOperatingTemperature.Text = Format$(sur.operatingtemperature.value - 273.15, "0.0")

       Call CalculateWaterPropertiesSurface

       Input #1, sur.PowerInput_PoverV.value, CommentDummy
       frmSurface!txtPowerInput.Text = Format$(sur.PowerInput_PoverV.value, GetTheFormat(sur.PowerInput_PoverV.value))

       Call CalculateOxygenLiquidDiffSurface

       Input #1, sur.Oxygen.KLaMethod, CommentDummy
       If sur.Oxygen.KLaMethod = KLA_METHOD_SURFACE_USER_INPUT Then
          Input #1, sur.Oxygen.MassTransferCoefficient.value, CommentDummy
          frmSurface!cboOxygen.ListIndex = 1
          frmSurface!txtOxygen(2).Text = Trim$(Str$(sur.Oxygen.MassTransferCoefficient.value))

       ElseIf sur.Oxygen.KLaMethod = KLA_METHOD_SURFACE_ROBERTS_CORRELATION Then
          frmSurface!cboOxygen.ListIndex = 0
          Call CalculateOxygenMTCoeffSurface
       End If

       Input #1, sur.NumChemical, CommentDummy
       For i = 1 To sur.NumChemical
           Input #1, sur.Contaminant(i).Pressure, sur.Contaminant(i).Temperature, sur.Contaminant(i).Name, sur.Contaminant(i).MolecularWeight.value, sur.Contaminant(i).HenrysConstant.value, sur.Contaminant(i).MolarVolume.value, sur.Contaminant(i).LiquidDiffusivity.value, sur.Contaminant(i).Influent.value, sur.Contaminant(i).TreatmentObjective.value
           'frmListContaminantSurface.ListContaminants.AddItem sur.Contaminant(i).Name
           frmSurface!cboDesignContaminant.AddItem sur.Contaminant(i).Name
       Next i

       Input #1, sur.DesignContaminant.Name, CommentDummy

       Call SetDesignContaminantEnabledSurface(CInt(frmSurface!cboDesignContaminant.ListCount))

       For i = 1 To sur.NumChemical
           If sur.DesignContaminant.Name = sur.Contaminant(i).Name Then
              sur.DesignContaminant = sur.Contaminant(i)
              'frmListContaminantSurface!ListContaminants.Selected(i - 1) = True
              SelectedContaminant = i - 1
              Exit For
           End If
       Next i

       'If frmListContaminantSurface.mnuOptionsManipulateContaminant(1).Enabled = False Then
       '   frmListContaminantSurface.mnuOptionsManipulateContaminant(1).Enabled = True
       '   frmListContaminantSurface.mnuOptionsManipulateContaminant(3).Enabled = True
       '   frmListContaminantSurface.mnuOptionsManipulateContaminant(4).Enabled = True
       '   frmListContaminantSurface.mnuOptionsSave.Enabled = True
       '   frmListContaminantSurface.mnuOptionsView.Enabled = True
       '   frmSurface!mnuFile(4).Enabled = True
       '   frmSurface!mnuFile(5).Enabled = True
       '   frmSurface!mnuOptions(0).Enabled = True
       'End If

       Call CalculateContaminantMTCoeffSurface

       Input #1, sur.WaterFlowRate.value, CommentDummy
       frmSurface!txtFlowParameters(0).Text = Format$(sur.WaterFlowRate.value, GetTheFormat(sur.WaterFlowRate.value))

       Input #1, sur.NumberOfTanks.value, CommentDummy
       frmSurface!txtTankParameters(0).Text = Format$(sur.NumberOfTanks.value, "0")

       Input #1, sur.CodeForTausAndTankVolumes, CommentDummy

          Select Case sur.CodeForTausAndTankVolumes
             Case 1   'Input Hydraulic Retention Time for 1 Tank
                Input #1, sur.TankHydraulicRetentionTime.value, CommentDummy
                frmSurface!txtTankParameters(1).Text = Format$(sur.TankHydraulicRetentionTime.value, GetTheFormat(sur.TankHydraulicRetentionTime.value))
                sur.TankHydraulicRetentionTime.UserInput = True
             Case 2   'Input Hydraulic Retention Time for All Tanks
                Input #1, sur.TotalHydraulicRetentionTime.value, CommentDummy
                frmSurface!txtTankParameters(2).Text = Format$(sur.TotalHydraulicRetentionTime.value, GetTheFormat(sur.TotalHydraulicRetentionTime.value))
                sur.TotalHydraulicRetentionTime.UserInput = True
             Case 3   'Input Volume of Each Tank
                Input #1, sur.TankVolume.value, CommentDummy
                frmSurface!txtTankParameters(3).Text = Format$(sur.TankVolume.value, GetTheFormat(sur.TankVolume.value))
                sur.TankVolume.UserInput = True
             Case 4   'Input Volume of All Tanks
                Input #1, sur.TotalTankVolume.value, CommentDummy
                frmSurface!txtTankParameters(4).Text = Format$(sur.TotalTankVolume.value, GetTheFormat(sur.TotalTankVolume.value))
                sur.TotalTankVolume.UserInput = True
          End Select

       Input #1, sur.Power.AeratorMotorEfficiency, CommentDummy
       frmSurface!txtPowerCalculation(0).Text = Format$(sur.Power.AeratorMotorEfficiency, "0.0")

       If SurfaceAerationMode = DESIGN_MODE Then
          sur.CodeForTausAndTankVolumes = 1
          Call CalculateRetentionTimeSurface
          For i = 1 To 4
              frmSurface!txtTankParameters(i).Enabled = False
          Next i
       Else
          For i = 1 To 4
              frmSurface!txtTankParameters(i).Enabled = True
          Next i
       End If

       Call CalculateTausAndTankVolumesSurface
       frmSurface.cboDesignContaminant.ListIndex = SelectedContaminant
       Call CalculatePowerSurface

       'Input the units of this screen.
       Input #1, u(1), u(2)
       Call SetUnits(frmSurface!UnitsOpCond(0), u(1))
       Call SetUnits(frmSurface!UnitsOpCond(1), u(2))
     
       Input #1, u(1)
       Call SetUnits(frmSurface!UnitsPowerInput, u(1))
     
       Input #1, u(1), u(2)
       Call SetUnits(frmSurface!UnitsOxygenRef(1), u(1))
       Call SetUnits(frmSurface!UnitsOxygenRef(2), u(2))
     
       Input #1, u(1), u(2), u(3)
       Call SetUnits(frmSurface!UnitsDesignContam(0), u(1))
       Call SetUnits(frmSurface!UnitsDesignContam(1), u(2))
       Call SetUnits(frmSurface!UnitsDesignContam(3), u(3))
     
       Input #1, u(1)
       Call SetUnits(frmSurface!UnitsFlowParam(0), u(1))
     
       Input #1, u(1), u(2), u(3), u(4)
       Call SetUnits(frmSurface!UnitsTankParam(1), u(1))
       Call SetUnits(frmSurface!UnitsTankParam(2), u(2))
       Call SetUnits(frmSurface!UnitsTankParam(3), u(3))
       Call SetUnits(frmSurface!UnitsTankParam(4), u(4))
     
       Input #1, u(1), u(2)
       Call SetUnits(frmSurface!UnitsConcResults(1), u(1))
       Call SetUnits(frmSurface!UnitsConcResults(3), u(2))
     
       Input #1, u(1), u(2)
       Call SetUnits(frmSurface!UnitsPowerCalc(1), u(1))
       Call SetUnits(frmSurface!UnitsPowerCalc(2), u(2))
     
       'Input the units of frmContaminantPropertyEdit.
       xu = Units_frmContaminantPropertyEdit
       Input #1, xu.UnitsProp(0), xu.UnitsProp(2), xu.UnitsProp(3), xu.UnitsProp(4), xu.UnitsProp(5)
       Input #1, xu.UnitsConc(0), xu.UnitsConc(1)
       Units_frmContaminantPropertyEdit = xu
     
       Close #1

       If Right$(Filename, 12) = "defltdes.sur" Or Right$(Filename, 12) = "defltrat.sur" Then
          frmSurface.Caption = frmSurface.Caption & " (" & "untitled.sur" & ")"
       Else
          frmSurface.Caption = frmSurface.Caption & " (" & Filename & ")"
       End If

       'Add this file to the last-few-files list.
       Call LastFewFiles_MoveFilenameToTop(Filename)

    End If

    loadsurface = True
    
End Function

Sub CalculateContaminantMTCoeffSurface()
Dim Dummy As Double

  Call KLASURF(sur.ContaminantMassTransferCoefficient.value, sur.Oxygen.MassTransferCoefficient.value, sur.DesignContaminant.LiquidDiffusivity.value, sur.Oxygen.LiquidDiffusivity.value, sur.N_for_Finding_KLa.value, sur.kgOVERkl_for_Finding_KLa.value, sur.DesignContaminant.HenrysConstant.value)
   
  'UPDATED_UNITS.
  'frmSurface!txtDesignConcentrationValue(3).Text = Format$(sur.ContaminantMassTransferCoefficient.Value, GetTheFormat(sur.ContaminantMassTransferCoefficient.Value))
  Call Unitted_NumberUpdate(frmSurface!UnitsDesignContam(3))

  sur.ContaminantMassTransferCoefficient.UserInput = False

End Sub

Sub CalculateEffluentConcentrationsSurface()
ReDim Effluent(0 To MAXIMUM_TANKS) As Double
Dim i As Integer
Dim SaveOldUnit As Integer
Dim Dummy As Double

  Call SEFFL(Effluent(1), sur.AchievedPercentRemoval, sur.DesignContaminant.Influent.value, sur.ContaminantMassTransferCoefficient.value, sur.TankHydraulicRetentionTime.value, sur.NumberOfTanks.value)
  For i = 1 To sur.NumberOfTanks.value
    sur.DesignContaminant.Effluent(i) = Effluent(i)
  Next i
  sur.DesignContaminant.Effluent(0) = sur.DesignContaminant.Influent.value

  'frmSurface!lblConcentrationResults(3).Caption = Format$(sur.DesignContaminant.Effluent(sur.NumberOfTanks.Value), GetTheFormat(sur.DesignContaminant.Effluent(sur.NumberOfTanks.Value)))
  'Dummy = sur.DesignContaminant.Effluent(sur.NumberOfTanks.Value)
  Call Unitted_NumberUpdate(frmSurface!UnitsConcResults(3))
  
  For i = 1 To sur.NumberOfTanks.value
    frmSurfaceEffluentConcentrations!lblTankNumber(i).Visible = True
    frmSurfaceEffluentConcentrations!lblLiquidEffluent(i).Visible = True
    frmSurfaceEffluentConcentrations!lblTankNumber(i).Caption = Trim$(Str$(i))
    frmSurfaceEffluentConcentrations!lblLiquidEffluent(i).Caption = Format$(sur.DesignContaminant.Effluent(i), GetTheFormat(sur.DesignContaminant.Effluent(i)))
  Next i
  
  For i = (sur.NumberOfTanks.value + 1) To MAXIMUM_TANKS
    frmSurfaceEffluentConcentrations!lblTankNumber(i).Visible = False
    frmSurfaceEffluentConcentrations!lblLiquidEffluent(i).Visible = False
  Next i

  i = sur.NumberOfTanks.value
  frmSurfaceEffluentConcentrations!cmdOK.Top = frmSurfaceEffluentConcentrations!lblTankNumber(i).Top + frmSurfaceEffluentConcentrations!lblTankNumber(i).Height + 300
  frmSurfaceEffluentConcentrations.Height = frmSurfaceEffluentConcentrations!cmdOK.Top + frmSurfaceEffluentConcentrations!cmdOK.Height + 500
  frmSurfaceEffluentConcentrations!cmdOK.Left = frmSurfaceEffluentConcentrations.Width / 2 - frmSurfaceEffluentConcentrations!cmdOK.Width / 2
        
  frmSurface!lblConcentrationResults(4).Caption = Format$(sur.AchievedPercentRemoval, GetTheFormat(sur.AchievedPercentRemoval))

End Sub

Sub CalculateOxygenLiquidDiffSurface()

  Call DIFO2(sur.Oxygen.LiquidDiffusivity.value, sur.operatingtemperature.value)
  'frmSurface!txtOxygen(1).Text = Format$(sur.Oxygen.LiquidDiffusivity.Value, GetTheFormat(sur.Oxygen.LiquidDiffusivity.Value))
  Call Unitted_NumberUpdate(frmSurface!UnitsOxygenRef(1))
  sur.Oxygen.LiquidDiffusivity.UserInput = False

End Sub

Sub PrintSurfaceToFile()
Dim i As Integer, j As Integer
Dim xu As rec_Units_frmContaminantPropertyEdit

    ReDim ContaminantMTCoeff(1 To MAXCHEMICAL) As Double
    ReDim DesiredPercentRemoval(1 To MAXCHEMICAL) As Double
    ReDim AchievedPercentRemoval(1 To MAXCHEMICAL) As Double
    ReDim Effluent(0 To MAXIMUM_TANKS) As Double
    ReDim GasEffluent(1 To MAXIMUM_TANKS) As Double

    xu = Units_frmContaminantPropertyEdit

        Call GetPrintFileName(PrintFileName)
        If PrintFileName$ = "" Then Exit Sub

        Open PrintFileName$ For Output As #1

    Select Case SurfaceAerationMode
       Case DESIGN_MODE

          Print #1, "Surface Aeration - Design Mode"
          Print #1,
          Print #1,
          'Printer.FontUnderline = True
          'Printer.FontSize = 10
          Print #1, "Property:"; Tab(VALUE_TAB); "Value:"
          Print #1,
          'Printer.FontUnderline = False
          'Printer.FontBold = False
          Print #1, "Operating Pressure (" & frmSurface!UnitsOpCond(0) & ")"; Tab(VALUE_TAB); frmSurface!txtOperatingPressure.Text
          Print #1, "Operating Temperature (" & frmSurface!UnitsOpCond(1) & ")"; Tab(VALUE_TAB); frmSurface!txtOperatingTemperature.Text
          Print #1, frmWaterPropertiesSurface!lblAirWaterProperties(0).Caption; Tab(VALUE_TAB); frmWaterPropertiesSurface!txtAirWaterProperties(0).Text
          Print #1, frmWaterPropertiesSurface!lblAirWaterProperties(1).Caption; Tab(VALUE_TAB); frmWaterPropertiesSurface!txtAirWaterProperties(1).Text
          Print #1,
          Print #1, frmSurface!lblPowerInputLabel.Caption & " (" & frmSurface!UnitsPowerInput & ")"; Tab(VALUE_TAB); frmSurface!txtPowerInput.Text
          Print #1,
          Print #1, "Oxygen " & frmSurface!lblOxygenLabel(1).Caption & " (" & frmSurface!UnitsOxygenRef(1) & ")"; Tab(VALUE_TAB); frmSurface!txtOxygen(1).Text
          Print #1, "Method to Find Oxygen KLa"; Tab(VALUE_TAB); frmSurface!cboOxygen.Text
          Print #1, "Oxygen " & frmSurface!lblOxygenLabel(2).Caption & " (" & frmSurface!UnitsOxygenRef(2) & ")"; Tab(VALUE_TAB); frmSurface!txtOxygen(2).Text
          Print #1,
          Print #1, "Design Contaminant:  "; frmSurface!cboDesignContaminant.Text
          Print #1, "Molecular Weight" & " (" & xu.UnitsProp(0) & ")"; Tab(VALUE_TAB); Format$(sur.DesignContaminant.MolecularWeight.value, "0.00")
          Print #1, "Henry's Constant (-)"; Tab(VALUE_TAB); Format$(sur.DesignContaminant.HenrysConstant.value, GetTheFormat(sur.DesignContaminant.HenrysConstant.value))
          Print #1, "Molar Volume" & " (" & xu.UnitsProp(2) & ")"; Tab(VALUE_TAB); Format$(sur.DesignContaminant.MolarVolume.value, GetTheFormat(sur.DesignContaminant.MolarVolume.value))
          Print #1, "Liquid Diffusivity" & " (" & xu.UnitsProp(4) & ")"; Tab(VALUE_TAB); Format$(sur.DesignContaminant.LiquidDiffusivity.value, GetTheFormat(sur.DesignContaminant.LiquidDiffusivity.value))
          Print #1, frmSurface!lblDesignConcentration(0).Caption & " (" & frmSurface!UnitsDesignContam(0) & ")"; Tab(VALUE_TAB); frmSurface!lblDesignConcentrationValue(0).Caption
          Print #1, frmSurface!lblDesignConcentration(1).Caption & " (" & frmSurface!UnitsDesignContam(1) & ")"; Tab(VALUE_TAB); frmSurface!lblDesignConcentrationValue(1).Caption
          Print #1, frmSurface!lblDesignConcentration(2).Caption & " (%)"; Tab(VALUE_TAB); frmSurface!lblDesignConcentrationValue(2).Caption
          Print #1, frmSurface!lblDesignConcentration(3).Caption & " (" & frmSurface!UnitsDesignContam(3) & ")"; Tab(VALUE_TAB); frmSurface!txtDesignConcentrationValue(3).Text
          Print #1,
          Print #1, frmSurface!lblFlowParametersLabel(0).Caption & " (" & frmSurface!UnitsFlowParam(0) & ")"; Tab(VALUE_TAB); frmSurface!txtFlowParameters(0).Text
          Print #1,
          Print #1, frmSurface!lblTankParametersLabel(0).Caption; Tab(VALUE_TAB); frmSurface!txtTankParameters(0).Text
          Print #1, frmSurface!lblTankParametersLabel(1).Caption & " (" & frmSurface!UnitsTankParam(1) & ")"; Tab(VALUE_TAB); frmSurface!txtTankParameters(1).Text
          Print #1, frmSurface!lblTankParametersLabel(2).Caption & " (" & frmSurface!UnitsTankParam(2) & ")"; Tab(VALUE_TAB); frmSurface!txtTankParameters(2).Text
          Print #1, frmSurface!lblTankParametersLabel(3).Caption & " (" & frmSurface!UnitsTankParam(3) & ")"; Tab(VALUE_TAB); frmSurface!txtTankParameters(3).Text
          Print #1, frmSurface!lblTankParametersLabel(4).Caption & " (" & frmSurface!UnitsTankParam(4) & ")"; Tab(VALUE_TAB); frmSurface!txtTankParameters(4).Text
          Print #1,
          Print #1, "Design Contaminant:  "; frmSurface!lblConcentrationResults(0).Caption
          Print #1, "Liquid Phase Influent Concentration to Tank 1" & " (" & frmSurface!UnitsConcResults(1) & ")"; Tab(VALUE_TAB); frmSurface!lblConcentrationResults(1).Caption
          Print #1, "Liquid Phase Effluent from Last Tank" & " (" & frmSurface!UnitsConcResults(3) & ")"; Tab(VALUE_TAB); frmSurface!lblConcentrationResults(3).Caption
          Print #1, "Achieved Percent Removal (%)"; Tab(VALUE_TAB); frmSurface!lblConcentrationResults(4).Caption
          Print #1,
          Print #1, "Effluent Concentrations from Each Tank in " & Chr$(181) & "g/L:"
          Print #1,
          Print #1, "Tank:"; Tab(LIQUID_EFFLUENT_TAB); "Effluent Conc."
          Print #1,
          For i = 1 To sur.NumberOfTanks.value
              Print #1, Format$(i, "0"); Tab(LIQUID_EFFLUENT_TAB); Format$(sur.DesignContaminant.Effluent(i), GetTheFormat(sur.DesignContaminant.Effluent(i)))
          Next i
             Print #1,
             Print #1,
          Print #1, "Power Calculation:"
          Print #1,
          Print #1, "Property:"; Tab(VALUE_TAB); "Value:"
          Print #1,
          Print #1, frmSurface!lblPowerCalculationLabel(0).Caption & " (%)"; Tab(VALUE_TAB); frmSurface!txtPowerCalculation(0).Text
          Print #1, frmSurface!lblPowerCalculationLabel(1).Caption & " (" & frmSurface!UnitsPowerCalc(1) & ")"; Tab(VALUE_TAB); frmSurface!lblPowerCalculation(1).Caption
          Print #1, frmSurface!lblPowerCalculationLabel(2).Caption & " (" & frmSurface!UnitsPowerCalc(2) & ")"; Tab(VALUE_TAB); frmSurface!lblPowerCalculation(2).Caption

       Case RATING_MODE
          Print #1, "Surface Aeration - Rating Mode"
          Print #1,
          Print #1,
          'Printer.FontUnderline = True
          'Printer.FontSize = 10
          Print #1, "Property:"; Tab(VALUE_TAB); "Value:"
          Print #1,
          'Printer.FontUnderline = False
          'Printer.FontBold = False
          Print #1, "Operating Pressure (" & frmSurface!UnitsOpCond(0) & ")"; Tab(VALUE_TAB); frmSurface!txtOperatingPressure.Text
          Print #1, "Operating Temperature (" & frmSurface!UnitsOpCond(1) & ")"; Tab(VALUE_TAB); frmSurface!txtOperatingTemperature.Text
          Print #1, frmWaterPropertiesSurface!lblAirWaterProperties(0).Caption; Tab(VALUE_TAB); frmWaterPropertiesSurface!txtAirWaterProperties(0).Text
          Print #1, frmWaterPropertiesSurface!lblAirWaterProperties(1).Caption; Tab(VALUE_TAB); frmWaterPropertiesSurface!txtAirWaterProperties(1).Text
          Print #1,
          Print #1, frmSurface!lblPowerInputLabel.Caption & " (" & frmSurface!UnitsPowerInput & ")"; Tab(VALUE_TAB); frmSurface!txtPowerInput.Text
          Print #1,
          Print #1, "Oxygen " & frmSurface!lblOxygenLabel(1).Caption & " (" & frmSurface!UnitsOxygenRef(1) & ")"; Tab(VALUE_TAB); frmSurface!txtOxygen(1).Text
          Print #1, "Method to Find Oxygen KLa"; Tab(VALUE_TAB); frmSurface!cboOxygen.Text
          Print #1, "Oxygen " & frmSurface!lblOxygenLabel(2).Caption & " (" & frmSurface!UnitsOxygenRef(2) & ")"; Tab(VALUE_TAB); frmSurface!txtOxygen(2).Text
          Print #1,
          Print #1, frmSurface!lblFlowParametersLabel(0).Caption & " (" & frmSurface!UnitsFlowParam(0) & ")"; Tab(VALUE_TAB); frmSurface!txtFlowParameters(0).Text
          Print #1,
          Print #1, frmSurface!lblTankParametersLabel(0).Caption; Tab(VALUE_TAB); frmSurface!txtTankParameters(0).Text
          Print #1, frmSurface!lblTankParametersLabel(1).Caption & " (" & frmSurface!UnitsTankParam(1) & ")"; Tab(VALUE_TAB); frmSurface!txtTankParameters(1).Text
          Print #1, frmSurface!lblTankParametersLabel(2).Caption & " (" & frmSurface!UnitsTankParam(2) & ")"; Tab(VALUE_TAB); frmSurface!txtTankParameters(2).Text
          Print #1, frmSurface!lblTankParametersLabel(3).Caption & " (" & frmSurface!UnitsTankParam(3) & ")"; Tab(VALUE_TAB); frmSurface!txtTankParameters(3).Text
          Print #1, frmSurface!lblTankParametersLabel(4).Caption & " (" & frmSurface!UnitsTankParam(4) & ")"; Tab(VALUE_TAB); frmSurface!txtTankParameters(4).Text
          Print #1,
          Print #1,
          'Printer.FontBold = True
          Print #1, "Power Calculation:"
          'Printer.FontUnderline = True
          Print #1,
          Print #1, "Property:"; Tab(VALUE_TAB); "Value:"
          'Printer.FontBold = False
          'Printer.FontUnderline = False
          Print #1,
          Print #1, frmSurface!lblPowerCalculationLabel(0).Caption & " (%)"; Tab(VALUE_TAB); frmSurface!txtPowerCalculation(0).Text
          Print #1, frmSurface!lblPowerCalculationLabel(1).Caption & " (" & frmSurface!UnitsPowerCalc(1) & ")"; Tab(VALUE_TAB); frmSurface!lblPowerCalculation(1).Caption
          Print #1, frmSurface!lblPowerCalculationLabel(2).Caption & " (" & frmSurface!UnitsPowerCalc(2) & ")"; Tab(VALUE_TAB); frmSurface!lblPowerCalculation(2).Caption
          
          Print #1,
          Print #1,
          Print #1, "Contaminant Glossary:"
          For i = 1 To sur.NumChemical
              Print #1, Format$(i, "0"); " = "; Trim$(sur.Contaminant(i).Name)
          Next i
             Print #1,
             Print #1,
             Print #1,
          Print #1, "Contaminant Properties:"
          Print #1,
         
          Print #1, "Con.:"; Tab(MWT_TAB); "MWT"; Tab(HC_TAB); "HC"; Tab(VB_TAB); "Vb"; Tab(DIFL_TAB); "DIFL"; Tab(MTCOEFF_TAB); "MT Coeff."
          Print #1,
          For i = 1 To sur.NumChemical
              If sur.DesignContaminant.Name = sur.Contaminant(i).Name Then
                 ContaminantMTCoeff(i) = sur.ContaminantMassTransferCoefficient.value
              Else
                 Call KLASURF(ContaminantMTCoeff(i), sur.Oxygen.MassTransferCoefficient.value, sur.Contaminant(i).LiquidDiffusivity.value, sur.Oxygen.LiquidDiffusivity.value, sur.N_for_Finding_KLa.value, sur.kgOVERkl_for_Finding_KLa.value, sur.Contaminant(i).HenrysConstant.value)
              End If
              Print #1, Format$(i, "0"); Tab(MWT_TAB); Format$(sur.Contaminant(i).MolecularWeight.value, "0.00"); Tab(HC_TAB); Format$(sur.Contaminant(i).HenrysConstant.value, GetTheFormat(sur.Contaminant(i).HenrysConstant.value)); Tab(VB_TAB); Format$(sur.Contaminant(i).MolarVolume.value, GetTheFormat(sur.Contaminant(i).MolarVolume.value)); Tab(DIFL_TAB); Format$(sur.Contaminant(i).LiquidDiffusivity.value, GetTheFormat(sur.Contaminant(i).LiquidDiffusivity.value)); Tab(MTCOEFF_TAB); Format$(ContaminantMTCoeff(i), GetTheFormat(ContaminantMTCoeff(i)))
          Next i
          Print #1,
         
          Print #1, "Glossary:"
       
          Print #1, "Con. = Contaminant Number (see Contaminant Glossary on page 1)"
          Print #1, "MWT = Molecular Weight (kg/kmol)"
          Print #1, "HC = Henry's Constant (-)"
          Print #1, "Vb = Molar Volume (m" & Chr$(179) & "/kmol)"
          Print #1, "DIFL = Liquid Diffusivity (m" & Chr$(178) & "/sec)"
          Print #1, "MT Coeff. = Mass Transfer Coeff. (1/sec)"
             Print #1,
             Print #1,
             Print #1,
          Print #1, "Contaminant Concentration Results:"
          Print #1,
         
          Print #1, "Con.:"; Tab(MWT_TAB); "Cinf"; Tab(HC_TAB); "Cto"; Tab(VB_TAB); "De. % Rem."; Tab(DIFL_TAB); "Ceff"; Tab(MTCOEFF_TAB); "Ach. % Rem."
          Print #1,
          For i = 1 To sur.NumChemical
              If sur.DesignContaminant.Name = sur.Contaminant(i).Name Then
                 DesiredPercentRemoval(i) = sur.DesiredPercentRemoval
                 sur.Contaminant(i).Effluent(0) = sur.DesignContaminant.Effluent(0)
                 For j = 1 To sur.NumberOfTanks.value
                     sur.Contaminant(i).Effluent(j) = sur.DesignContaminant.Effluent(j)
                 Next j
                 AchievedPercentRemoval(i) = sur.AchievedPercentRemoval
              Else
                 Call REMOVBUB(DesiredPercentRemoval(i), sur.Contaminant(i).Influent.value, sur.Contaminant(i).TreatmentObjective.value)
                 Effluent(0) = sur.Contaminant(i).Influent.value
                 Call SEFFL(Effluent(1), AchievedPercentRemoval(i), sur.Contaminant(i).Influent.value, ContaminantMTCoeff(i), sur.TankHydraulicRetentionTime.value, sur.NumberOfTanks.value)
                 sur.Contaminant(i).Effluent(0) = Effluent(0)
                 For j = 0 To sur.NumberOfTanks.value
                     sur.Contaminant(i).Effluent(j) = Effluent(j)
                 Next j
              End If
              Print #1, Format$(i, "0"); Tab(MWT_TAB); Format$(sur.Contaminant(i).Influent.value, GetTheFormat(sur.Contaminant(i).Influent.value)); Tab(HC_TAB); Format$(sur.Contaminant(i).TreatmentObjective.value, GetTheFormat(sur.Contaminant(i).TreatmentObjective.value)); Tab(VB_TAB); Format$(DesiredPercentRemoval(i), GetTheFormat(DesiredPercentRemoval(i))); Tab(DIFL_TAB); Format$(sur.Contaminant(i).Effluent(sur.NumberOfTanks.value), GetTheFormat(sur.Contaminant(i).Effluent(sur.NumberOfTanks.value))); Tab(MTCOEFF_TAB); Format$(AchievedPercentRemoval(i), GetTheFormat(AchievedPercentRemoval(i)))
          Next i
          Print #1,
        
          Print #1, "Glossary:"
      
          Print #1, "Con. = Contaminant Number (see Contaminant Glossary on page 1)"
          Print #1, "Cinf = "; "Liquid Phase " & frmSurface!lblDesignConcentration(0).Caption & " (" & Chr$(181) & "g/L)"
          Print #1, "Cto = "; frmSurface!lblDesignConcentration(1).Caption & " (" & Chr$(181) & "g/L)"
          Print #1, "De. % Rem. = "; frmSurface!lblDesignConcentration(2).Caption
          Print #1, "Ceff = "; "Liquid Phase Effluent from Last Tank (" & Chr$(181) & "g/L)"
          Print #1, "Ach. % Rem. = "; frmSurface!lblConcentrationResultsLabel(4).Caption
             Print #1,
             Print #1,
             Print #1,
          Print #1, "Liquid Phase Effluent Concentrations from Each Tank in " & Chr$(181) & "g/L:"
          Print #1,
          Print #1,
          Print #1, Tab(MWT_TAB); "Contaminant Number:"
          Print #1,
       
          Select Case sur.NumChemical
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
              If sur.NumChemical < j Then
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
              Print #1, Format$(sur.Contaminant(j).Influent.value, GetTheFormat(sur.Contaminant(j).Influent.value));
          Next j
          Print #1,

          'Print Liquid Phase Effluent Concentrations from each tank for each contaminant
          For i = 1 To sur.NumberOfTanks.value
              Print #1, Format$(i, "0");
              For j = 1 To 6
                  If sur.NumChemical < j Then
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
                  Print #1, Format$(sur.Contaminant(j).Effluent(i), GetTheFormat(sur.Contaminant(j).Effluent(i)));
             Next j
             Print #1,
          Next i
          
          If sur.NumChemical < 7 Then
             Print #1,
             Print #1, "Glossary:"
             Print #1, "Cinf = Liquid Phase Influent Concentration to Tank 1 (" & Chr$(181) & "g/L)"
             GoTo AfterLiquidEffluentsSurface
          End If
          Print #1,
          Print #1,
          Print #1, Tab(MWT_TAB); "Contaminant Number:"
          Print #1,
          Select Case sur.NumChemical
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
              If sur.NumChemical < j Then
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
              Print #1, Format$(sur.Contaminant(j).Influent.value, GetTheFormat(sur.Contaminant(j).Influent.value));
          Next j
          Print #1,

          'Print Liquid Phase Effluent Concentrations from each tank for each contaminant
          For i = 1 To sur.NumberOfTanks.value
              Print #1, Format$(i, "0");
              For j = 7 To 10
                  If sur.NumChemical < j Then
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
                  Print #1, Format$(sur.Contaminant(j).Effluent(i), GetTheFormat(sur.Contaminant(j).Effluent(i)));
             Next j
             Print #1,
          Next i
          
          Print #1,
         
          Print #1, "Glossary:"
         
          Print #1, "Cinf = Liquid Phase Influent Concentration to Tank 1 (" & Chr$(181) & "g/L)"

AfterLiquidEffluentsSurface:

    End Select

    Close #1

End Sub

Sub SaveContaminantListSurface()
    Dim FileID As String
    Dim i As Integer

    Call SaveFile(Filename)

    If Filename$ <> "" Then
       FileID = CONTAMINANTS_SURFACE_FILEID
       Open Filename$ For Output As #1
       
       Write #1, FileID
      
       For i = 1 To sur.NumChemical
           Write #1, sur.Contaminant(i).Pressure, sur.Contaminant(i).Temperature, sur.Contaminant(i).Name, sur.Contaminant(i).MolecularWeight.value, sur.Contaminant(i).HenrysConstant.value, sur.Contaminant(i).MolarVolume.value, sur.Contaminant(i).LiquidDiffusivity.value, sur.Contaminant(i).Influent.value, sur.Contaminant(i).TreatmentObjective.value
       Next i

       Close #1

    End If

End Sub

Sub savefilesurface(Filename As String)
Dim Ctl As Control
Set Ctl = frmSurface.CommonDialog1

    On Error Resume Next
    'frmSurface!CMDialog1.DefaultExt = "sur"
    'frmSurface!CMDialog1.Filter = "Surface Aeration Files (*.sur)|*.sur"
    'frmSurface!CMDialog1.DialogTitle = "Save Surface Aeration File"
    'frmSurface!CMDialog1.flags = OFN_OVERWRITEPROMPT Or OFN_PATHMUSTEXIST
    'frmSurface!CMDialog1.Action = 2
    'Filename$ = frmSurface!CMDialog1.Filename
    Ctl.DefaultExt = "sur"
    Ctl.Filter = "Surface Aeration Files (*.sur)|*.sur"
    Ctl.DialogTitle = "Save Surface Aeration File"
    Ctl.flags = OFN_OVERWRITEPROMPT Or OFN_PATHMUSTEXIST
    Ctl.Action = 2
    Filename$ = Ctl.Filename
    If Err = 32755 Then   'Cancel selected by user
       Filename$ = ""
    End If

End Sub

Sub SaveSurface()
Dim FileID As String
Dim i As Integer
Dim xu As rec_Units_frmContaminantPropertyEdit
 
  If (IsThisADemo() = True) Then
    Call Demo_ShowError("Saving is not allowed in the demonstration version.")
    Exit Sub
  End If
    
    If Right$(frmSurface.Caption, 14) = "(untitled.sur)" Then
       Call savefilesurface(Filename)
    End If
    If Filename$ <> "" Then
       FileID = SURFACE_FILEID
       Open Filename$ For Output As #1
       
       Write #1, FileID
      
       Write #1, SurfaceAerationMode, ""

       Write #1, sur.OperatingPressure.value, ""
       Write #1, sur.operatingtemperature.value, ""

       Write #1, sur.PowerInput_PoverV.value, ""

       If sur.Oxygen.KLaMethod = KLA_METHOD_USER_INPUT Then
          Write #1, sur.Oxygen.KLaMethod, "KLaMethod:  2 = User Input"
          Write #1, sur.Oxygen.MassTransferCoefficient.value, ""
       ElseIf sur.Oxygen.KLaMethod = KLA_METHOD_CWO2_TRANSFER_TEST Then
          Write #1, sur.Oxygen.KLaMethod, ""
       End If

       Write #1, sur.NumChemical, ""
       For i = 1 To sur.NumChemical
           Write #1, sur.Contaminant(i).Pressure, sur.Contaminant(i).Temperature, sur.Contaminant(i).Name, sur.Contaminant(i).MolecularWeight.value, sur.Contaminant(i).HenrysConstant.value, sur.Contaminant(i).MolarVolume.value, sur.Contaminant(i).LiquidDiffusivity.value, sur.Contaminant(i).Influent.value, sur.Contaminant(i).TreatmentObjective.value
       Next i
       Write #1, sur.DesignContaminant.Name, ""

       Write #1, sur.WaterFlowRate.value, ""
       Write #1, sur.NumberOfTanks.value, ""

       Write #1, sur.CodeForTausAndTankVolumes, ""
       Select Case sur.CodeForTausAndTankVolumes
          Case 1   'Write Hydraulic Retention Time for 1 Tank
             Write #1, sur.TankHydraulicRetentionTime.value, ""
          Case 2   'Write Hydraulic Retention Time for All Tanks
             Write #1, sur.TotalHydraulicRetentionTime.value, ""
          Case 3   'Write Volume of Each Tank
             Write #1, sur.TankVolume.value, ""
          Case 4   'Write Volume of All Tanks
             Write #1, sur.TotalTankVolume.value, ""
       End Select

       Write #1, sur.Power.AeratorMotorEfficiency, ""

       'Output the units of this screen.
       Write #1, GetUnits(frmSurface!UnitsOpCond(0)), GetUnits(frmSurface!UnitsOpCond(1))
       Write #1, GetUnits(frmSurface!UnitsPowerInput)
       Write #1, GetUnits(frmSurface!UnitsOxygenRef(1)), GetUnits(frmSurface!UnitsOxygenRef(2))
       Write #1, GetUnits(frmSurface!UnitsDesignContam(0)), GetUnits(frmSurface!UnitsDesignContam(1)), GetUnits(frmSurface!UnitsDesignContam(3))
       Write #1, GetUnits(frmSurface!UnitsFlowParam(0))
       Write #1, GetUnits(frmSurface!UnitsTankParam(1)), GetUnits(frmSurface!UnitsTankParam(2)), GetUnits(frmSurface!UnitsTankParam(3)), GetUnits(frmSurface!UnitsTankParam(4))
       Write #1, GetUnits(frmSurface!UnitsConcResults(1)), GetUnits(frmSurface!UnitsConcResults(3))
       Write #1, GetUnits(frmSurface!UnitsPowerCalc(1)), GetUnits(frmSurface!UnitsPowerCalc(2))
       
       'Output the units of frmContaminantPropertyEdit.
       xu = Units_frmContaminantPropertyEdit
       Write #1, xu.UnitsProp(0), xu.UnitsProp(2), xu.UnitsProp(3), xu.UnitsProp(4), xu.UnitsProp(5)
       Write #1, xu.UnitsConc(0), xu.UnitsConc(1)
       
       Close #1

       If SurfaceAerationMode = DESIGN_MODE Then
          frmSurface.Caption = "Surface Aeration - Design Mode"
       Else
          frmSurface.Caption = "Surface Aeration - Rating Mode"
       End If

       frmSurface.Caption = frmSurface.Caption & " (" & Filename & ")"

    End If

End Sub

Sub SetDesignContaminantEnabledSurface(NumInList As Integer)
    Dim i As Integer

    If NumInList = 0 Then
       frmSurface!mnuFile(4).Enabled = False
       frmSurface!mnuFile(5).Enabled = False
       frmSurface!mnuOptions(0).Enabled = False
       'frmSurface!fraDesignContaminant.Enabled = False
       frmSurface!cboDesignContaminant.Enabled = False
       frmSurface!fraTankParameters.Enabled = False
       frmSurface!fraConcentrationResults.Enabled = False
       frmSurface!fraPower.Enabled = False

       frmSurface!lblPowerCalculation(1).Caption = ""
       frmSurface!lblPowerCalculation(2).Caption = ""
       For i = 0 To 2
           frmSurface!lblDesignConcentrationValue(i).Caption = ""
       Next i
       frmSurface!txtDesignConcentrationValue(3).Text = ""
       If SurfaceAerationMode = DESIGN_MODE Then
          frmSurface.txtTankParameters(1).Text = ""
          frmSurface.txtTankParameters(2).Text = ""
          frmSurface.txtTankParameters(3).Text = ""
          frmSurface.txtTankParameters(4).Text = ""
       End If

       frmSurface!lblConcentrationResults(0).Caption = ""
       frmSurface!lblConcentrationResults(1).Caption = ""
       frmSurface!lblConcentrationResults(3).Caption = ""
       frmSurface!lblConcentrationResults(4).Caption = ""

    Else
       
       frmSurface!mnuFile(4).Enabled = True
       frmSurface!mnuFile(5).Enabled = True
       
       frmSurface!mnuOptions(0).Enabled = True

       'frmSurface!fraDesignContaminant.Enabled = True
       frmSurface!cboDesignContaminant.Enabled = True
       frmSurface!fraTankParameters.Enabled = True
       frmSurface!fraConcentrationResults.Enabled = True
       frmSurface!fraPower.Enabled = True

    End If
    Call frmSurface.LOCAL___Reset_DemoVersionDisablings
End Sub

Sub NewPageSurface()

          Printer.NewPage
          Printer.FontSize = 12
          Printer.FontBold = True
          If SurfaceAerationMode = DESIGN_MODE Then
             Printer.Print "Surface Aeration - Design Mode (continued)"
          Else
             Printer.Print "Surface Aeration - Rating Mode (continued)"
          End If
          Printer.Print
          Printer.Print
          Printer.FontSize = 10
          Printer.FontBold = False

End Sub

Sub PrintSurface()
Dim i As Integer, j As Integer
ReDim ContaminantMTCoeff(1 To MAXCHEMICAL) As Double
ReDim DesiredPercentRemoval(1 To MAXCHEMICAL) As Double
ReDim AchievedPercentRemoval(1 To MAXCHEMICAL) As Double
ReDim Effluent(0 To MAXIMUM_TANKS) As Double
ReDim GasEffluent(1 To MAXIMUM_TANKS) As Double
Dim xu As rec_Units_frmContaminantPropertyEdit

    xu = Units_frmContaminantPropertyEdit

    On Error GoTo PrinterError

    Select Case SurfaceAerationMode
       Case DESIGN_MODE

          Printer.ScaleLeft = -1440
          Printer.ScaleTop = -1440
          Printer.CurrentX = 0
          Printer.CurrentY = 0
          Printer.FontSize = 12
          Printer.FontBold = True
          
          Printer.Print "Surface Aeration - Design Mode"
          Printer.Print
          Printer.Print
          Printer.FontUnderline = True
          Printer.FontSize = 10
          Printer.Print "Property:"; Tab(VALUE_TAB); "Value:"
          Printer.Print
          Printer.FontUnderline = False
          Printer.FontBold = False
          Printer.Print "Operating Pressure (" & frmSurface!UnitsOpCond(0) & ")"; Tab(VALUE_TAB); frmSurface!txtOperatingPressure.Text
          Printer.Print "Operating Temperature (" & frmSurface!UnitsOpCond(1) & ")"; Tab(VALUE_TAB); frmSurface!txtOperatingTemperature.Text
          Printer.Print frmWaterPropertiesSurface!lblAirWaterProperties(0).Caption; Tab(VALUE_TAB); frmWaterPropertiesSurface!txtAirWaterProperties(0).Text
          Printer.Print frmWaterPropertiesSurface!lblAirWaterProperties(1).Caption; Tab(VALUE_TAB); frmWaterPropertiesSurface!txtAirWaterProperties(1).Text
          Printer.Print
          Printer.Print frmSurface!lblPowerInputLabel.Caption & " (" & frmSurface!UnitsPowerInput & ")"; Tab(VALUE_TAB); frmSurface!txtPowerInput.Text
          Printer.Print
          Printer.Print "Oxygen " & frmSurface!lblOxygenLabel(1).Caption & " (" & frmSurface!UnitsOxygenRef(1) & ")"; Tab(VALUE_TAB); frmSurface!txtOxygen(1).Text
          Printer.Print "Method to Find Oxygen KLa"; Tab(VALUE_TAB); frmSurface!cboOxygen.Text
          Printer.Print "Oxygen " & frmSurface!lblOxygenLabel(2).Caption & " (" & frmSurface!UnitsOxygenRef(2) & ")"; Tab(VALUE_TAB); frmSurface!txtOxygen(2).Text
          Printer.Print
          Printer.Print "Design Contaminant:  "; frmSurface!cboDesignContaminant.Text
          Printer.Print "Molecular Weight" & " (" & xu.UnitsProp(0) & ")"; Tab(VALUE_TAB); Format$(sur.DesignContaminant.MolecularWeight.value, "0.00")
          Printer.Print "Henry's Constant (-)"; Tab(VALUE_TAB); Format$(sur.DesignContaminant.HenrysConstant.value, GetTheFormat(sur.DesignContaminant.HenrysConstant.value))
          Printer.Print "Molar Volume" & " (" & xu.UnitsProp(2) & ")"; Tab(VALUE_TAB); Format$(sur.DesignContaminant.MolarVolume.value, GetTheFormat(sur.DesignContaminant.MolarVolume.value))
          Printer.Print "Liquid Diffusivity" & " (" & xu.UnitsProp(4) & ")"; Tab(VALUE_TAB); Format$(sur.DesignContaminant.LiquidDiffusivity.value, GetTheFormat(sur.DesignContaminant.LiquidDiffusivity.value))
          Printer.Print frmSurface!lblDesignConcentration(0).Caption & " (" & frmSurface!UnitsDesignContam(0) & ")"; Tab(VALUE_TAB); frmSurface!lblDesignConcentrationValue(0).Caption
          Printer.Print frmSurface!lblDesignConcentration(1).Caption & " (" & frmSurface!UnitsDesignContam(1) & ")"; Tab(VALUE_TAB); frmSurface!lblDesignConcentrationValue(1).Caption
          Printer.Print frmSurface!lblDesignConcentration(2).Caption & " (%)"; Tab(VALUE_TAB); frmSurface!lblDesignConcentrationValue(2).Caption
          Printer.Print frmSurface!lblDesignConcentration(3).Caption & " (" & frmSurface!UnitsDesignContam(3) & ")"; Tab(VALUE_TAB); frmSurface!txtDesignConcentrationValue(3).Text
          Printer.Print
          Printer.Print frmSurface!lblFlowParametersLabel(0).Caption & " (" & frmSurface!UnitsFlowParam(0) & ")"; Tab(VALUE_TAB); frmSurface!txtFlowParameters(0).Text
          Printer.Print
          Printer.Print frmSurface!lblTankParametersLabel(0).Caption; Tab(VALUE_TAB); frmSurface!txtTankParameters(0).Text
          Printer.Print frmSurface!lblTankParametersLabel(1).Caption & " (" & frmSurface!UnitsTankParam(1) & ")"; Tab(VALUE_TAB); frmSurface!txtTankParameters(1).Text
          Printer.Print frmSurface!lblTankParametersLabel(2).Caption & " (" & frmSurface!UnitsTankParam(2) & ")"; Tab(VALUE_TAB); frmSurface!txtTankParameters(2).Text
          Printer.Print frmSurface!lblTankParametersLabel(3).Caption & " (" & frmSurface!UnitsTankParam(3) & ")"; Tab(VALUE_TAB); frmSurface!txtTankParameters(3).Text
          Printer.Print frmSurface!lblTankParametersLabel(4).Caption & " (" & frmSurface!UnitsTankParam(4) & ")"; Tab(VALUE_TAB); frmSurface!txtTankParameters(4).Text
          Printer.Print
          Printer.Print "Design Contaminant:  "; frmSurface!lblConcentrationResults(0).Caption
          Printer.Print "Liquid Phase Influent Concentration to Tank 1" & " (" & frmSurface!UnitsConcResults(1) & ")"; Tab(VALUE_TAB); frmSurface!lblConcentrationResults(1).Caption
          Printer.Print "Liquid Phase Effluent from Last Tank" & " (" & frmSurface!UnitsConcResults(3) & ")"; Tab(VALUE_TAB); frmSurface!lblConcentrationResults(3).Caption
          Printer.Print "Achieved Percent Removal (%)"; Tab(VALUE_TAB); frmSurface!lblConcentrationResults(4).Caption
          Printer.Print
          Printer.FontBold = True
          Printer.Print "Effluent Concentrations from Each Tank in " & Chr$(181) & "g/L:"
          Printer.Print
          Printer.FontUnderline = True
          Printer.Print "Tank:"; Tab(LIQUID_EFFLUENT_TAB); "Effluent Conc."
          Printer.Print
          Printer.FontBold = False
          Printer.FontUnderline = False
          For i = 1 To sur.NumberOfTanks.value
              Printer.Print Format$(i, "0"); Tab(LIQUID_EFFLUENT_TAB); Format$(sur.DesignContaminant.Effluent(i), GetTheFormat(sur.DesignContaminant.Effluent(i)))
          Next i
          If sur.NumberOfTanks.value > 8 Then
             Call NewPageSurface
          Else
             Printer.Print
             Printer.Print
          End If
          Printer.FontBold = True
          Printer.Print "Power Calculation:"
          Printer.FontUnderline = True
          Printer.Print
          Printer.Print "Property:"; Tab(VALUE_TAB); "Value:"
          Printer.FontBold = False
          Printer.FontUnderline = False
          Printer.Print
          Printer.Print frmSurface!lblPowerCalculationLabel(0).Caption; Tab(VALUE_TAB); frmSurface!txtPowerCalculation(0).Text
          Printer.Print frmSurface!lblPowerCalculationLabel(1).Caption; Tab(VALUE_TAB); frmSurface!lblPowerCalculation(1).Caption
          Printer.Print frmSurface!lblPowerCalculationLabel(2).Caption; Tab(VALUE_TAB); frmSurface!lblPowerCalculation(2).Caption

       Case RATING_MODE
          Printer.ScaleLeft = -1440
          Printer.ScaleTop = -1440
          Printer.CurrentX = 0
          Printer.CurrentY = 0
          Printer.FontSize = 12
          Printer.FontBold = True
          Printer.Print "Surface Aeration - Rating Mode"
          Printer.Print
          Printer.Print
          Printer.FontUnderline = True
          Printer.FontSize = 10
          Printer.Print "Property:"; Tab(VALUE_TAB); "Value:"
          Printer.Print
          Printer.FontUnderline = False
          Printer.FontBold = False
          Printer.Print "Operating Pressure (" & frmSurface!UnitsOpCond(0) & ")"; Tab(VALUE_TAB); frmSurface!txtOperatingPressure.Text
          Printer.Print "Operating Temperature (" & frmSurface!UnitsOpCond(1) & ")"; Tab(VALUE_TAB); frmSurface!txtOperatingTemperature.Text
          Printer.Print frmWaterPropertiesSurface!lblAirWaterProperties(0).Caption; Tab(VALUE_TAB); frmWaterPropertiesSurface!txtAirWaterProperties(0).Text
          Printer.Print frmWaterPropertiesSurface!lblAirWaterProperties(1).Caption; Tab(VALUE_TAB); frmWaterPropertiesSurface!txtAirWaterProperties(1).Text
          Printer.Print
          Printer.Print frmSurface!lblPowerInputLabel.Caption & " (" & frmSurface!UnitsPowerInput & ")"; Tab(VALUE_TAB); frmSurface!txtPowerInput.Text
          Printer.Print
          Printer.Print "Oxygen " & frmSurface!lblOxygenLabel(1).Caption & " (" & frmSurface!UnitsOxygenRef(1) & ")"; Tab(VALUE_TAB); frmSurface!txtOxygen(1).Text
          Printer.Print "Method to Find Oxygen KLa"; Tab(VALUE_TAB); frmSurface!cboOxygen.Text
          Printer.Print "Oxygen " & frmSurface!lblOxygenLabel(2).Caption & " (" & frmSurface!UnitsOxygenRef(2) & ")"; Tab(VALUE_TAB); frmSurface!txtOxygen(2).Text
          Printer.Print
          Printer.Print frmSurface!lblFlowParametersLabel(0).Caption & " (" & frmSurface!UnitsFlowParam(0) & ")"; Tab(VALUE_TAB); frmSurface!txtFlowParameters(0).Text
          Printer.Print
          Printer.Print frmSurface!lblTankParametersLabel(0).Caption; Tab(VALUE_TAB); frmSurface!txtTankParameters(0).Text
          Printer.Print frmSurface!lblTankParametersLabel(1).Caption & " (" & frmSurface!UnitsTankParam(1) & ")"; Tab(VALUE_TAB); frmSurface!txtTankParameters(1).Text
          Printer.Print frmSurface!lblTankParametersLabel(2).Caption & " (" & frmSurface!UnitsTankParam(2) & ")"; Tab(VALUE_TAB); frmSurface!txtTankParameters(2).Text
          Printer.Print frmSurface!lblTankParametersLabel(3).Caption & " (" & frmSurface!UnitsTankParam(3) & ")"; Tab(VALUE_TAB); frmSurface!txtTankParameters(3).Text
          Printer.Print frmSurface!lblTankParametersLabel(4).Caption & " (" & frmSurface!UnitsTankParam(4) & ")"; Tab(VALUE_TAB); frmSurface!txtTankParameters(4).Text
          Printer.Print
          Printer.Print
          Printer.FontBold = True
          Printer.Print "Power Calculation:"
          Printer.FontUnderline = True
          Printer.Print
          Printer.Print "Property:"; Tab(VALUE_TAB); "Value:"
          Printer.FontBold = False
          Printer.FontUnderline = False
          Printer.Print
          Printer.Print frmSurface!lblPowerCalculationLabel(0).Caption & " (%)"; Tab(VALUE_TAB); frmSurface!txtPowerCalculation(0).Text
          Printer.Print frmSurface!lblPowerCalculationLabel(1).Caption & " (" & frmSurface!UnitsPowerCalc(1) & ")"; Tab(VALUE_TAB); frmSurface!lblPowerCalculation(1).Caption
          Printer.Print frmSurface!lblPowerCalculationLabel(2).Caption & " (" & frmSurface!UnitsPowerCalc(2) & ")"; Tab(VALUE_TAB); frmSurface!lblPowerCalculation(2).Caption
          Printer.Print
          Printer.Print
          Printer.FontUnderline = True
          Printer.Print "Contaminant Glossary:"
          Printer.FontUnderline = False
          For i = 1 To sur.NumChemical
              Printer.Print Format$(i, "0"); " = "; Trim$(sur.Contaminant(i).Name)
          Next i

          If sur.NumChemical > 6 Then
             Call NewPageSurface
          Else
             Printer.Print
             Printer.Print
          End If
          Printer.FontBold = True
          Printer.Print "Contaminant Properties:"
          Printer.Print
          Printer.FontUnderline = True
          Printer.Print "Con.:"; Tab(MWT_TAB); "MWT"; Tab(HC_TAB); "HC"; Tab(VB_TAB); "Vb"; Tab(DIFL_TAB); "DIFL"; Tab(MTCOEFF_TAB); "MT Coeff."
          Printer.Print
          Printer.FontUnderline = False
          Printer.FontBold = False
          For i = 1 To sur.NumChemical
              If sur.DesignContaminant.Name = sur.Contaminant(i).Name Then
                 ContaminantMTCoeff(i) = sur.ContaminantMassTransferCoefficient.value
              Else
                 Call KLASURF(ContaminantMTCoeff(i), sur.Oxygen.MassTransferCoefficient.value, sur.Contaminant(i).LiquidDiffusivity.value, sur.Oxygen.LiquidDiffusivity.value, sur.N_for_Finding_KLa.value, sur.kgOVERkl_for_Finding_KLa.value, sur.Contaminant(i).HenrysConstant.value)
              End If
              Printer.Print Format$(i, "0"); Tab(MWT_TAB); Format$(sur.Contaminant(i).MolecularWeight.value, "0.00"); Tab(HC_TAB); Format$(sur.Contaminant(i).HenrysConstant.value, GetTheFormat(sur.Contaminant(i).HenrysConstant.value)); Tab(VB_TAB); Format$(sur.Contaminant(i).MolarVolume.value, GetTheFormat(sur.Contaminant(i).MolarVolume.value)); Tab(DIFL_TAB); Format$(sur.Contaminant(i).LiquidDiffusivity.value, GetTheFormat(sur.Contaminant(i).LiquidDiffusivity.value)); Tab(MTCOEFF_TAB); Format$(ContaminantMTCoeff(i), GetTheFormat(ContaminantMTCoeff(i)))
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
          If sur.NumChemical <= 6 Then
             Call NewPageSurface
          Else
             Printer.Print
             Printer.Print
             Printer.Print
          End If
          Printer.FontBold = True
          Printer.Print "Contaminant Concentration Results:"
          Printer.Print
          Printer.FontUnderline = True
          Printer.Print "Con.:"; Tab(MWT_TAB); "Cinf"; Tab(HC_TAB); "Cto"; Tab(VB_TAB); "De. % Rem."; Tab(DIFL_TAB); "Ceff"; Tab(MTCOEFF_TAB); "Ach. % Rem."
          Printer.Print
          Printer.FontUnderline = False
          Printer.FontBold = False
          For i = 1 To sur.NumChemical
              If sur.DesignContaminant.Name = sur.Contaminant(i).Name Then
                 DesiredPercentRemoval(i) = sur.DesiredPercentRemoval
                 sur.Contaminant(i).Effluent(0) = sur.DesignContaminant.Effluent(0)
                 For j = 1 To sur.NumberOfTanks.value
                     sur.Contaminant(i).Effluent(j) = sur.DesignContaminant.Effluent(j)
                 Next j
                 AchievedPercentRemoval(i) = sur.AchievedPercentRemoval
              Else
                 Call REMOVBUB(DesiredPercentRemoval(i), sur.Contaminant(i).Influent.value, sur.Contaminant(i).TreatmentObjective.value)
                 Effluent(0) = sur.Contaminant(i).Influent.value
                 Call SEFFL(Effluent(1), AchievedPercentRemoval(i), sur.Contaminant(i).Influent.value, ContaminantMTCoeff(i), sur.TankHydraulicRetentionTime.value, sur.NumberOfTanks.value)
                 sur.Contaminant(i).Effluent(0) = Effluent(0)
                 For j = 0 To sur.NumberOfTanks.value
                     sur.Contaminant(i).Effluent(j) = Effluent(j)
                 Next j
              End If
              Printer.Print Format$(i, "0"); Tab(MWT_TAB); Format$(sur.Contaminant(i).Influent.value, GetTheFormat(sur.Contaminant(i).Influent.value)); Tab(HC_TAB); Format$(sur.Contaminant(i).TreatmentObjective.value, GetTheFormat(sur.Contaminant(i).TreatmentObjective.value)); Tab(VB_TAB); Format$(DesiredPercentRemoval(i), GetTheFormat(DesiredPercentRemoval(i))); Tab(DIFL_TAB); Format$(sur.Contaminant(i).Effluent(sur.NumberOfTanks.value), GetTheFormat(sur.Contaminant(i).Effluent(sur.NumberOfTanks.value))); Tab(MTCOEFF_TAB); Format$(AchievedPercentRemoval(i), GetTheFormat(AchievedPercentRemoval(i)))
          Next i
          Printer.Print
          Printer.FontUnderline = True
          Printer.Print "Glossary:"
          Printer.FontUnderline = False
          Printer.Print "Con. = Contaminant Number (see Contaminant Glossary on page 1)"
          Printer.Print "Cinf = "; "Liquid Phase " & frmSurface!lblDesignConcentration(0).Caption & " (" & Chr$(181) & "g/L)"
          Printer.Print "Cto = "; frmSurface!lblDesignConcentration(1).Caption & " (" & Chr$(181) & "g/L)"
          Printer.Print "De. % Rem. = "; frmSurface!lblDesignConcentration(2).Caption
          Printer.Print "Ceff = "; "Liquid Phase Effluent from Last Tank (" & Chr$(181) & "g/L)"
          Printer.Print "Ach. % Rem. = "; frmSurface!lblConcentrationResultsLabel(4).Caption
          If sur.NumChemical > 6 Then
             Call NewPageSurface
          Else
             Printer.Print
             Printer.Print
             Printer.Print
          End If
          Printer.FontBold = True
          Printer.Print "Liquid Phase Effluent Concentrations from Each Tank in " & Chr$(181) & "g/L:"
          Printer.Print
          Printer.Print
          Printer.Print Tab(MWT_TAB); "Contaminant Number:"
          Printer.Print
          Printer.FontUnderline = True
          Select Case sur.NumChemical
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
              If sur.NumChemical < j Then
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
              Printer.Print Format$(sur.Contaminant(j).Influent.value, GetTheFormat(sur.Contaminant(j).Influent.value));
          Next j
          Printer.Print

          'Print Liquid Phase Effluent Concentrations from each tank for each contaminant
          For i = 1 To sur.NumberOfTanks.value
              Printer.Print Format$(i, "0");
              For j = 1 To 6
                  If sur.NumChemical < j Then
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
                  Printer.Print Format$(sur.Contaminant(j).Effluent(i), GetTheFormat(sur.Contaminant(j).Effluent(i)));
             Next j
             Printer.Print
          Next i
          
          If sur.NumChemical < 7 Then
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
          Select Case sur.NumChemical
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
              If sur.NumChemical < j Then
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
              Printer.Print Format$(sur.Contaminant(j).Influent.value, GetTheFormat(sur.Contaminant(j).Influent.value));
          Next j
          Printer.Print

          'Print Liquid Phase Effluent Concentrations from each tank for each contaminant
          For i = 1 To sur.NumberOfTanks.value
              Printer.Print Format$(i, "0");
              For j = 7 To 10
                  If sur.NumChemical < j Then
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
                  Printer.Print Format$(sur.Contaminant(j).Effluent(i), GetTheFormat(sur.Contaminant(j).Effluent(i)));
             Next j
             Printer.Print
          Next i
          
          Printer.Print
          Printer.FontUnderline = True
          Printer.Print "Glossary:"
          Printer.FontUnderline = False
          Printer.Print "Cinf = Liquid Phase Influent Concentration to Tank 1 (" & Chr$(181) & "g/L)"

AfterLiquidEffluents:

    End Select

    Printer.EndDoc

    Exit Sub

PrinterError:
    MsgBox error$(Err)
    Resume ExitPrint:

ExitPrint:

End Sub

Function StartSurfaceDefaultCase() As Boolean

    Filename = "TheDefaultCaseSurface"
    StartSurfaceDefaultCase = loadsurface("")

End Function

Sub surface_results()
    Dim i As Integer, j As Integer
    ReDim ContaminantMTCoeff(1 To MAXCHEMICAL) As Double
    ReDim DesiredPercentRemoval(1 To MAXCHEMICAL) As Double
    ReDim AchievedPercentRemoval(1 To MAXCHEMICAL) As Double
    ReDim Effluent(0 To MAXIMUM_TANKS) As Double
    Dim ContaminantGlossaryBottom As Integer, GlossaryBottom As Integer

          For i = 1 To sur.NumChemical
              If sur.DesignContaminant.Name = sur.Contaminant(i).Name Then
                 DesiredPercentRemoval(i) = sur.DesiredPercentRemoval
                 ContaminantMTCoeff(i) = sur.ContaminantMassTransferCoefficient.value
                 sur.Contaminant(i).Effluent(0) = sur.DesignContaminant.Effluent(0)
                 For j = 1 To sur.NumberOfTanks.value
                     sur.Contaminant(i).Effluent(j) = sur.DesignContaminant.Effluent(j)
                 Next j
                 AchievedPercentRemoval(i) = sur.AchievedPercentRemoval
              Else
                 Call REMOVBUB(DesiredPercentRemoval(i), sur.Contaminant(i).Influent.value, sur.Contaminant(i).TreatmentObjective.value)
                 Call KLASURF(ContaminantMTCoeff(i), sur.Oxygen.MassTransferCoefficient.value, sur.Contaminant(i).LiquidDiffusivity.value, sur.Oxygen.LiquidDiffusivity.value, sur.N_for_Finding_KLa.value, sur.kgOVERkl_for_Finding_KLa.value, sur.Contaminant(i).HenrysConstant.value)
                 Effluent(0) = sur.Contaminant(i).Influent.value
                 Call SEFFL(Effluent(1), AchievedPercentRemoval(i), sur.Contaminant(i).Influent.value, ContaminantMTCoeff(i), sur.TankHydraulicRetentionTime.value, sur.NumberOfTanks.value)
                 sur.Contaminant(i).Effluent(0) = Effluent(0)
                 For j = 0 To sur.NumberOfTanks.value
                     sur.Contaminant(i).Effluent(j) = Effluent(j)
                 Next j
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

    For i = 1 To sur.NumChemical
        frmViewEffluentConcentrationsASAP!lblContaminantNumber(i - 1).Visible = True
        frmViewEffluentConcentrationsASAP!lblContaminantNumber(i + 10 - 1).Visible = True
        frmViewEffluentConcentrationsASAP!lblInfluentConcentration(i - 1).Visible = True
        frmViewEffluentConcentrationsASAP!lblTreatmentObjective(i - 1).Visible = True
        frmViewEffluentConcentrationsASAP!lblDesiredPercentRemoval(i - 1).Visible = True
        frmViewEffluentConcentrationsASAP!lblEffluentConcentration(i - 1).Visible = True
        frmViewEffluentConcentrationsASAP!lblAchievedPercentRemoval(i - 1).Visible = True
        frmViewEffluentConcentrationsASAP!lblContaminantName(i - 1).Visible = True

        frmViewEffluentConcentrationsASAP!lblInfluentConcentration(i - 1).Caption = Format$(sur.Contaminant(i).Influent.value, GetTheFormat(sur.Contaminant(i).Influent.value))
        frmViewEffluentConcentrationsASAP!lblTreatmentObjective(i - 1).Caption = Format$(sur.Contaminant(i).TreatmentObjective.value, GetTheFormat(sur.Contaminant(i).TreatmentObjective.value))
        frmViewEffluentConcentrationsASAP!lblDesiredPercentRemoval(i - 1).Caption = Format$(DesiredPercentRemoval(i), "0.0")
        frmViewEffluentConcentrationsASAP!lblEffluentConcentration(i - 1).Caption = Format$(sur.Contaminant(i).Effluent(sur.NumberOfTanks.value), GetTheFormat(sur.Contaminant(i).Effluent(sur.NumberOfTanks.value)))
        frmViewEffluentConcentrationsASAP!lblAchievedPercentRemoval(i - 1).Caption = Format$(AchievedPercentRemoval(i), "0.0")
        frmViewEffluentConcentrationsASAP!lblContaminantName(i - 1).Caption = Trim$(LCase$(sur.Contaminant(i).Name))

    Next i

    frmViewEffluentConcentrationsASAP!fraConcentrationResults.Height = frmViewEffluentConcentrationsASAP!lblContaminantNumber(sur.NumChemical - 1).Top + frmViewEffluentConcentrationsASAP!lblContaminantNumber(sur.NumChemical - 1).Height + 120
    frmViewEffluentConcentrationsASAP!fraContaminantGlossary.Height = frmViewEffluentConcentrationsASAP!lblContaminantNumber(sur.NumChemical + 10 - 1).Top + frmViewEffluentConcentrationsASAP!lblContaminantNumber(sur.NumChemical + 10 - 1).Height + 120
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

Function surface_savechanges() As Integer
Dim i As Integer
Dim msg As String, Response As Integer

msg = "Would you like to save the parameters "
msg = msg + "for this surface aeration design case to a file?" & Chr$(13) & Chr$(13)
msg = msg + "Note:  Any information not saved will be permanently lost."
Response = MsgBox(msg, MB_ICONquestion + MB_YESNOCANCEL, "Save Current Design")
                
If Response = IDCANCEL Then
 Screen.MousePointer = 0
 surface_savechanges = 1
 Exit Function
End If

If Response = IDYES Then
  Call SaveSurface
   If StrComp(Filename, "") = 0 Then Response = 5
      Do While Response = 5
         msg = "Would you like to save the parameters "
         msg = msg + "for this surface aeration design case to a file?" & Chr$(13) & Chr$(13)
         msg = msg + "Note:  Any information not saved will be permanently lost."
           Response = MsgBox(msg, MB_ICONquestion + MB_YESNOCANCEL, "Save Current Design")
                           
           If Response = IDCANCEL Then
           Screen.MousePointer = 0
           surface_savechanges = 1
           Exit Function
           End If
                           
           If Response = IDYES Then Call SaveSurface
           If StrComp(Filename, "") = 0 And Response <> IDNO Then Response = 5
      Loop
   End If

surface_savechanges = 0
End Function

