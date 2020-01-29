Attribute VB_Name = "ASAP_Mod"
Option Explicit

''''''''''''''''''''''''''''
' Visual Basic global constant file. This file can be loaded
' into a code module.
'
' Some constants are commented out because they have
' duplicates (e.g., NONE appears several places).
'
' If you are updating a Visual Basic application written with
' an older version, you should replace your global constants
' with the constants in this file.
'
''''''''''''''''''''''''''''

' General
Global Const Application_Name = "ASAP"
Global NL As String
Global AddFlag As Integer

' ErrNum (LinkError)
Global Const WRONG_FORMAT = 1
Global Const DDE_SOURCE_CLOSED = 6
Global Const TOO_MANY_LINKS = 7
Global Const DATA_TRANSFER_FAILED = 8

' Enumerated Types

' Align (picture box)
Global Const NONE = 0
Global Const ALIGN_TOP = 1
Global Const ALIGN_BOTTOM = 2

' BorderStyle (form)
Global Const FIXED_SINGLE = 1   ' 1 - Fixed Single
Global Const SIZABLE = 2        ' 2 - Sizable (Forms only)
Global Const FIXED_DOUBLE = 3   ' 3 - Fixed Double (Forms only)

' LinkMode (forms and controls)
' Global Const NONE = 0         ' 0 - None
Global Const LINK_SOURCE = 1    ' 1 - Source (forms only)
Global Const LINK_AUTOMATIC = 1 ' 1 - Automatic (controls only)
Global Const LINK_MANUAL = 2    ' 2 - Manual (controls only)
Global Const LINK_NOTIFY = 3    ' 3 - Notify (controls only)

' ScaleMode
Global Const TWIPS = 1       ' 1 - Twip
Global Const PIXELS = 3      ' 3 - Pixel

' Function Parameters
' MsgBox parameters
Global Const MB_OK = 0                 ' OK button only
Global Const MB_OKCANCEL = 1           ' OK and Cancel buttons
Global Const MB_ABORTRETRYIGNORE = 2   ' Abort, Retry, and Ignore buttons
Global Const MB_YESNOCANCEL = 3        ' Yes, No, and Cancel buttons
Global Const MB_YESNO = 4              ' Yes and No buttons
Global Const MB_RETRYCANCEL = 5        ' Retry and Cancel buttons

Global Const MB_ICONSTOP = 16          ' Critical message
Global Const MB_ICONquestion = 32      ' Warning query
Global Const MB_ICONEXCLAMATION = 48   ' Warning message
Global Const MB_ICONINFORMATION = 64   ' Information message

' MsgBox return values
Global Const IDOK = 1                  ' OK button pressed
Global Const IDCANCEL = 2              ' Cancel button pressed
Global Const IDABORT = 3               ' Abort button pressed
Global Const IDRETRY = 4               ' Retry button pressed
Global Const IDIGNORE = 5              ' Ignore button pressed
Global Const IDYES = 6                 ' Yes button pressed
Global Const IDNO = 7                  ' No button pressed

' SetAttr, Dir, GetAttr functions
Global Const ATTR_NORMAL = 0
Global Const ATTR_READONLY = 1
Global Const ATTR_HIDDEN = 2
Global Const ATTR_SYSTEM = 4
Global Const ATTR_VOLUME = 8
Global Const ATTR_DIRECTORY = 16
Global Const ATTR_ARCHIVE = 32


'File Open/Save Dialog Flags
Global Const OFN_READONLY = &H1&
Global Const OFN_OVERWRITEPROMPT = &H2&
Global Const OFN_HIDEREADONLY = &H4&
Global Const OFN_NOCHANGEDIR = &H8&
Global Const OFN_SHOWHELP = &H10&
Global Const OFN_NOVALIDATE = &H100&
Global Const OFN_ALLOWMULTISELECT = &H200&
Global Const OFN_EXTENSIONDIFFERENT = &H400&
Global Const OFN_PATHMUSTEXIST = &H800&
Global Const OFN_FILEMUSTEXIST = &H1000&
Global Const OFN_CREATEPROMPT = &H2000&
Global Const OFN_SHAREAWARE = &H4000&
Global Const OFN_NOREADONLYRETURN = &H8000&


'Printer Dialog Flags
Global Const PD_PRINTSETUP = &H40&

Global ContaminantData(0 To MAXCHEMICAL) As StrippingContaminantProperties

Global Const TOLERANCE = 0.00000000000001
Global Const NUMBER_CHANGING_CRITERIA = 1E-17

Global PressureChanged As Integer
Global TemperatureChanged As Integer

Global ShownForm As Integer

Global scr1 As SCR  'Data structure for screen 1 of PTAD
Global Scr2 As SCR  'Data structure for screen 2 of PTAD

Global Temp_Text As String

Global IsError As Integer

Global ListContaminantMenuOptionsIndex As Integer

Global OriginalProperties As ContaminantPropertyType

Global Const CONTAMINANTS_PTAD_FILEID = "Contaminant_Properties_PTAD"
Global Const SCREEN1_PTAD1_FILEID = "DesignProperties_PTAD_Screen1"
Global Const SCREEN2_PTAD2_FILEID = "RatingProperties_PTAD_Screen2"

Global Filename As String
Global PrintFileName As String

Global ErrorFlag As Long

Global CurrentScreen As SCR   'Current screen user is
                              'manipulating - Screen 1 or
                              'Screen 2 in PTAD

Global ScreenNumber As Integer


Global DefaultPacking As PackingDataType

'Screen Sizes for a Standard VGA Screen
Global Const SCREEN_WIDTH_STANDARD = 9600
Global Const SCREEN_HEIGHT_STANDARD = 7200

Sub CalculateAirWaterProperties()
    Dim Pressure As Double
    Dim Temperature As Double
    Dim WaterDensity As Double
    Dim WaterViscosity As Double
    Dim WaterSurfaceTension As Double
    Dim AirDensity As Double
    Dim AirViscosity As Double
    Dim i As Integer
    
    If scr1.OperatingPressure.ValChanged Or scr1.operatingtemperature.ValChanged Then
       Pressure = scr1.OperatingPressure.value
       Temperature = scr1.operatingtemperature.value

       For i = 0 To 4
           If frmAirWaterProperties!chkUpdateValues(i).value = True Then
              Select Case i
                 Case 0
                    If HaveValue(Temperature) Then
                       Call H2ODENS(WaterDensity, Temperature)
                       scr1.WaterDensity.value = WaterDensity
                       scr1.WaterDensity.UserInput = False
                       scr1.WaterDensity.ValChanged = True
                       frmAirWaterProperties.txtAirWaterProperties(0).Text = Format$(WaterDensity, "0.00")
                       frmAirWaterProperties.lblValueSource(0).Caption = "Correlation"
                    End If
                 Case 1
                    If HaveValue(Temperature) Then
                       Call H2OVISC(WaterViscosity, Temperature)
                       scr1.WaterViscosity.value = WaterViscosity
                       scr1.WaterViscosity.UserInput = False
                       scr1.WaterViscosity.ValChanged = True
                       frmAirWaterProperties.txtAirWaterProperties(1).Text = Format$(WaterViscosity, GetTheFormat(WaterViscosity))
                       frmAirWaterProperties.lblValueSource(1).Caption = "Correlation"
                    End If
                 Case 2
                    If HaveValue(Temperature) Then
                       Call H2OST(WaterSurfaceTension, Temperature)
                       scr1.WaterSurfaceTension.value = WaterSurfaceTension
                       scr1.WaterSurfaceTension.UserInput = False
                       scr1.WaterSurfaceTension.ValChanged = True
                       frmAirWaterProperties.txtAirWaterProperties(2).Text = Format$(WaterSurfaceTension, GetTheFormat(WaterSurfaceTension))
                       frmAirWaterProperties.lblValueSource(2).Caption = "Correlation"
                    End If
                 Case 3
                    If HaveValue(Temperature) And HaveValue(Pressure) Then
                       Call AIRDENS(AirDensity, Temperature, Pressure)
                       scr1.AirDensity.value = AirDensity
                       scr1.AirDensity.UserInput = False
                       scr1.AirDensity.ValChanged = True
                       frmAirWaterProperties.txtAirWaterProperties(3).Text = Format$(AirDensity, GetTheFormat(AirDensity))
                       frmAirWaterProperties.lblValueSource(3).Caption = "Correlation"
                    End If
                 Case 4
                    If HaveValue(Temperature) Then
                       Call AIRVISC(AirViscosity, Temperature)
                       scr1.AirViscosity.value = AirViscosity
                       scr1.AirViscosity.UserInput = False
                       scr1.AirViscosity.ValChanged = True
                       frmAirWaterProperties.txtAirWaterProperties(4).Text = Format$(AirViscosity, GetTheFormat(AirViscosity))
                       frmAirWaterProperties.lblValueSource(4).Caption = "Correlation"
                    End If
              End Select
          End If
       Next i
    End If
End Sub

Sub CalculatePowerScreen1(CalculatedPower As Integer)
Dim CalculatedBlowerPower As Integer
Dim CalculatedPumpPower As Integer

  CalculatedBlowerPower = False
  If HaveValue(scr1.AirFlowRate.value) And HaveValue(scr1.TowerArea.value) And HaveValue(scr1.OperatingPressure.value) And HaveValue(scr1.AirPressureDrop.value) And HaveValue(scr1.TowerHeight.value) And HaveValue(scr1.AirDensity.value) Then
    Call PBLOWPT(scr1.Power.BlowerBrakePower, scr1.AirFlowRate.value, scr1.TowerArea.value, scr1.OperatingPressure.value, scr1.AirPressureDrop.value, scr1.TowerHeight.value, scr1.AirDensity.value, scr1.Power.InletAirTemperature, scr1.Power.BlowerEfficiency)
    CalculatedBlowerPower = True
  End If

  CalculatedPumpPower = False
  If HaveValue(scr1.WaterDensity.value) And HaveValue(scr1.WaterFlowRate.value) And HaveValue(scr1.TowerHeight.value) Then
    Call PPUMPPT(scr1.Power.PumpBrakePower, scr1.Power.PumpEfficiency, scr1.WaterDensity.value, scr1.WaterFlowRate.value, scr1.TowerHeight.value)
    CalculatedPumpPower = True
  End If

  If CalculatedBlowerPower And CalculatedPumpPower Then
    Call PTOTALPT(scr1.Power.TotalBrakePower, scr1.Power.BlowerBrakePower, scr1.Power.PumpBrakePower)
    CalculatedPower = True
  End If

End Sub

Sub GetDesignKLaOrKLaSafetyFactor()

  If scr1.KLaSafetyFactor.UserInput = True Then
    Call SpecifiedKLaSafetyFactor
  ElseIf scr1.DesignMassTransferCoefficient.UserInput = True Then
    Call SpecifiedDesignMassTransferCoefficient
  End If

End Sub

Sub GetLoadings()

  If HaveValue(scr1.AirPressureDrop.value) And HaveValue(scr1.AirToWaterRatio.value) And HaveValue(scr1.AirDensity.value) And HaveValue(scr1.WaterDensity.value) And HaveValue(scr1.Packing.PackingFactor) And HaveValue(scr1.WaterViscosity.value) Then
    Call PT1LDAIR(scr1.AirLoadingRate.value, scr1.AirPressureDrop.value, scr1.AirToWaterRatio.value, scr1.AirDensity.value, scr1.WaterDensity.value, scr1.Packing.PackingFactor, scr1.WaterViscosity.value)
    scr1.AirLoadingRate.ValChanged = True
    scr1.AirLoadingRate.UserInput = False
    'frmPTADScreen1.lblFlowsLoadings(6).Caption = Format$(Scr1.AirLoadingRate.Value, GetTheFormat(Scr1.AirLoadingRate.Value))
    Call Unitted_NumberUpdate(frmPTADScreen1!lblFlowsUnits(6))
  End If

  If HaveValue(scr1.AirToWaterRatio.value) And HaveValue(scr1.AirDensity.value) And HaveValue(scr1.WaterDensity.value) And HaveValue(scr1.AirLoadingRate.value) Then
    Call PT1LDH2O(scr1.WaterLoadingRate.value, scr1.AirToWaterRatio.value, scr1.AirDensity.value, scr1.WaterDensity.value, scr1.AirLoadingRate.value)
    scr1.WaterLoadingRate.ValChanged = True
    scr1.WaterLoadingRate.UserInput = False
    'frmPTADScreen1.lblFlowsLoadings(7).Caption = Format$(Scr1.WaterLoadingRate.Value, GetTheFormat(Scr1.WaterLoadingRate.Value))
    Call Unitted_NumberUpdate(frmPTADScreen1!lblFlowsUnits(7))
  End If

End Sub

Sub GetOndaMassTransferCoefficient()

  If HaveValue(scr1.Packing.CriticalSurfaceTension) And HaveValue(scr1.WaterSurfaceTension.value) And HaveValue(scr1.WaterLoadingRate.value) And HaveValue(scr1.Packing.SpecificSurfaceArea) And HaveValue(scr1.WaterViscosity.value) And HaveValue(scr1.WaterDensity.value) And HaveValue(scr1.DesignContaminant.LiquidDiffusivity.value) And HaveValue(scr1.Packing.NominalSize) And HaveValue(scr1.AirLoadingRate.value) And HaveValue(scr1.AirViscosity.value) And HaveValue(scr1.AirDensity.value) And HaveValue(scr1.DesignContaminant.GasDiffusivity.value) And HaveValue(scr1.DesignContaminant.HenrysConstant.value) Then
    Call AWCALC(scr1.Packing.OndaWettedSurfaceArea, scr1.Packing.CriticalSurfaceTension, scr1.WaterSurfaceTension.value, scr1.WaterLoadingRate.value, scr1.Packing.SpecificSurfaceArea, scr1.WaterViscosity.value, scr1.WaterDensity.value, scr1.Onda.ReynoldsNumber, scr1.Onda.FroudeNumber, scr1.Onda.WeberNumber)
    Call ONDAKLPT(scr1.Onda.LiquidPhaseMassTransferCoefficient, scr1.WaterLoadingRate.value, scr1.Packing.OndaWettedSurfaceArea, scr1.WaterViscosity.value, scr1.WaterDensity.value, scr1.DesignContaminant.LiquidDiffusivity.value, scr1.Packing.SpecificSurfaceArea, scr1.Packing.NominalSize)
    Call ONDAKGPT(scr1.Onda.GasPhaseMassTransferCoefficient, scr1.AirLoadingRate.value, scr1.Packing.SpecificSurfaceArea, scr1.AirViscosity.value, scr1.AirDensity.value, scr1.DesignContaminant.GasDiffusivity.value, scr1.Packing.NominalSize)
    Call ONDKLAPT(scr1.Onda.OverallMassTransferCoefficient, scr1.Onda.LiquidPhaseMassTransferResistance, scr1.Onda.GasPhaseMassTransferResistance, scr1.Onda.TotalMassTransferResistance, scr1.Onda.LiquidPhaseMassTransferCoefficient, scr1.Packing.OndaWettedSurfaceArea, scr1.Onda.GasPhaseMassTransferCoefficient, scr1.DesignContaminant.HenrysConstant.value)

    'frmPTADScreen1!lblMassTransfer(0).Caption = Format$(Scr1.Onda.OverallMassTransferCoefficient, GetTheFormat(Scr1.Onda.OverallMassTransferCoefficient))
    Call Unitted_NumberUpdate(frmPTADScreen1!UnitsMassTransfer(0))
    scr1.Onda.ValChanged = True

    Call ShowOndaKLaProperties

  End If

End Sub

Sub GetPrintFileName(PrintFileName As String)
Dim Ctl As Control
Set Ctl = frmPTADScreen1.CommonDialog1

  On Error Resume Next
  'frmPTADScreen1!CMDialog1.DefaultExt = "prt"
  'frmPTADScreen1!CMDialog1.Filter = "Print Files (*.prt)|*.prt"
  'frmPTADScreen1!CMDialog1.DialogTitle = "Print ASAP Results To File"
  'frmPTADScreen1!CMDialog1.flags = OFN_OVERWRITEPROMPT Or OFN_PATHMUSTEXIST
  'frmPTADScreen1!CMDialog1.Action = 2
  'PrintFileName$ = frmPTADScreen1!CMDialog1.Filename
  Ctl.DefaultExt = "prt"
  Ctl.Filter = "Print Files (*.prt)|*.prt"
  Ctl.DialogTitle = "Print ASAP Results To File"
  Ctl.flags = OFN_OVERWRITEPROMPT Or OFN_PATHMUSTEXIST
  Ctl.Action = 2
  PrintFileName$ = Ctl.Filename
  If Err = 32755 Then   'Cancel selected by user
    PrintFileName$ = ""
  End If

End Sub

Sub GetTowerAreaAndDiameter()

  If HaveValue(scr1.WaterFlowRate.value) And HaveValue(scr1.WaterDensity.value) And HaveValue(scr1.WaterLoadingRate.value) Then
    Call PT1AREA(scr1.TowerArea.value, scr1.WaterFlowRate.value, scr1.WaterDensity.value, scr1.WaterLoadingRate.value)
    scr1.TowerArea.ValChanged = True
    scr1.TowerArea.UserInput = False
    'frmPTADScreen1!lblTowerParameters(0).Caption = Format$(Scr1.TowerArea.Value, GetTheFormat(Scr1.TowerArea.Value))
    Call Unitted_NumberUpdate(frmPTADScreen1!lblTowerUnits(0))
  End If

  If HaveValue(scr1.TowerArea.value) Then
    Call PT1DTOW(scr1.TowerDiameter.value, scr1.TowerArea.value)
    scr1.TowerDiameter.ValChanged = True
    scr1.TowerDiameter.UserInput = False
    'frmPTADScreen1!lblTowerParameters(1).Caption = Format$(Scr1.TowerDiameter.Value, GetTheFormat(Scr1.TowerDiameter.Value))
    Call Unitted_NumberUpdate(frmPTADScreen1!lblTowerUnits(1))
  End If

End Sub

Sub GetTowerHeightAndVolume()

  If HaveValue(scr1.AirToWaterRatio.value) And HaveValue(scr1.DesignContaminant.HenrysConstant.value) And HaveValue(scr1.DesignContaminant.Influent.value) And HaveValue(scr1.DesignContaminant.TreatmentObjective.value) And HaveValue(scr1.WaterFlowRate.value) And HaveValue(scr1.TowerArea.value) And HaveValue(scr1.DesignMassTransferCoefficient.value) Then
    Call GETCSPT(scr1.DesignContaminant.AirWaterInterfaceConcentration, scr1.AirToWaterRatio.value, scr1.DesignContaminant.HenrysConstant.value, scr1.DesignContaminant.Influent.value, scr1.DesignContaminant.TreatmentObjective.value)
    Call GETHTUPT(scr1.TransferUnitHeight, scr1.WaterFlowRate.value, scr1.TowerArea.value, scr1.DesignMassTransferCoefficient.value)
    Call GETNTUPT(scr1.NumberOfTransferUnits, scr1.DesignContaminant.Influent.value, scr1.DesignContaminant.TreatmentObjective.value, scr1.DesignContaminant.AirWaterInterfaceConcentration)
    Call PT1HTOW(scr1.TowerHeight.value, scr1.TransferUnitHeight, scr1.NumberOfTransferUnits)
    Call PT1TVOL(scr1.TowerVolume.value, scr1.TowerArea.value, scr1.TowerHeight.value)

    'frmPTADScreen1!lblTowerParameters(2).Caption = Format$(Scr1.TowerHeight.Value, GetTheFormat(Scr1.TowerHeight.Value))
    'frmPTADScreen1!lblTowerParameters(3).Caption = Format$(Scr1.TowerVolume.Value, GetTheFormat(Scr1.TowerVolume.Value))
    Call Unitted_NumberUpdate(frmPTADScreen1!lblTowerUnits(2))
    Call Unitted_NumberUpdate(frmPTADScreen1!lblTowerUnits(3))

  End If

End Sub

Sub GetVQmultVQAndAirFlowRate()

  If scr1.MultipleOfMinimumAirToWaterRatio.UserInput = True Then
    Call SpecifiedVQminMultiple
  ElseIf scr1.AirToWaterRatio.UserInput = True Then
    Call SpecifiedAirToWaterRatio
  ElseIf scr1.AirFlowRate.UserInput = True Then
    Call SpecifiedAirFlowRate
  End If

End Sub

Function HaveValue(value As Double) As Integer

  If value > 0# Then HaveValue = True Else HaveValue = False

End Function

Sub InitializeAirPressureDrop()

  scr1.AirPressureDrop.value = 50#
  scr1.AirPressureDrop.UserInput = True
  scr1.AirPressureDrop.ValChanged = True

  'frmPTADScreen1!txtFlowsLoadings(5).Text = Format$(Scr1.AirPressureDrop.Value, GetTheFormat(Scr1.AirPressureDrop.Value))
  Call Unitted_NumberUpdate(frmPTADScreen1!txtFlowsUnits(5))

End Sub

Sub InitializeCalculatedProperties()
Dim i As Integer
    
    'Flow and Loading Properties
 
    frmPTADScreen1.lblFlowsLoadings(1).Caption = "0.0"

    For i = 3 To 4
        frmPTADScreen1.txtFlowsLoadings(i).Text = "0.0"
    Next i

    For i = 6 To 7
        frmPTADScreen1.lblFlowsLoadings(i).Caption = "0.0"
    Next i

    frmPTADScreen1!txtFlowsLoadings(0).Enabled = False
    frmPTADScreen1.lblFlowsLoadings(1).Enabled = False
       
    For i = 2 To 5
        frmPTADScreen1!txtFlowsLoadings(i).Enabled = False
    Next i

    For i = 6 To 7
        frmPTADScreen1!lblFlowsLoadings(i).Enabled = False
    Next i


    'Mass Transfer Properties

    frmPTADScreen1.lblMassTransfer(0).Caption = "0.0"
    frmPTADScreen1.txtMassTransfer(2).Text = "0.0"

    frmPTADScreen1!lblMassTransfer(0).Enabled = False
    For i = 1 To 2
        frmPTADScreen1.txtMassTransfer(i).Enabled = False
    Next i


    'Tower Parameter Properties

    For i = 0 To 3
        frmPTADScreen1!lblTowerParameters(i).Caption = "0.0"
        frmPTADScreen1!lblTowerParameters(i).Enabled = False
    Next i

End Sub

Sub InitializeKLaSafetyFactor()

  scr1.KLaSafetyFactor.value = 1#
  scr1.KLaSafetyFactor.UserInput = True
  scr1.KLaSafetyFactor.ValChanged = True
  scr1.DesignMassTransferCoefficient.UserInput = False

  frmPTADScreen1!txtMassTransfer(1).Text = Format$(scr1.KLaSafetyFactor.value, GetTheFormat(scr1.KLaSafetyFactor.value))

End Sub

Sub InitializePacking()

'*** This subroutine initializes the packing to a default
'*** value.

    Dim i As Integer
    Dim packingname As String

    PackingDatabaseSource = ORIGINALPACKINGDATABASE
    packingname = "Tri-Packs_No.2"

' DEMO MODE CHANGE ::TACK
    If DemoMode% Then
        packingname = "Tri-Packs_No.1"
    End If
' END DEMO CHANGE

    For i = 1 To NumPackingsInDatabase
        If DatabasePacking(i).Name = packingname Then
           scr1.Packing = DatabasePacking(i)
        End If
    Next i

    frmPTADScreen1!lblPackingType.Caption = packingname
    
End Sub

Sub InitializePressureTemperature()
    
    '*****************************************************
    '*                                                   *
    '* Initialize Pressure and Temperature to defaults:  *
    '*                                                   *
    '*  Operating Pressure = 1 atm                       *
    '*  Operating Temperature = 283.15 K                 *
    '*                                                   *
    '*  Note:  Operating Pressure stored as atm but      *
    '*         displayed in Pa.  Operating Temperature   *
    '*         stored as K but displayed in C            *
    '*                                                   *
    '*****************************************************

    scr1.OperatingPressure.value = 1#
    scr1.OperatingPressure.UserInput = True
    scr1.OperatingPressure.ValChanged = True
    scr1.operatingtemperature.value = 283.15
    scr1.operatingtemperature.UserInput = True
    scr1.operatingtemperature.ValChanged = True

    frmPTADScreen1!txtOperatingPressure.Text = "101325.0"
    frmPTADScreen1!txtOperatingTemperature.Text = "10.0"

End Sub

Sub InitializeVQminMultiple()

  scr1.MultipleOfMinimumAirToWaterRatio.value = 3.5
  scr1.MultipleOfMinimumAirToWaterRatio.UserInput = True
  scr1.MultipleOfMinimumAirToWaterRatio.ValChanged = True
  scr1.AirToWaterRatio.UserInput = False
  scr1.AirFlowRate.UserInput = False

  frmPTADScreen1.txtFlowsLoadings(2).Text = Format$(scr1.MultipleOfMinimumAirToWaterRatio.value, GetTheFormat(scr1.MultipleOfMinimumAirToWaterRatio.value))

End Sub

Sub InitializeWaterFlowRate()

  scr1.WaterFlowRate.value = 0.1262 'm3/sec = 2000 gpm
  scr1.WaterFlowRate.UserInput = True
  scr1.WaterFlowRate.ValChanged = True
  frmPTADScreen1.txtFlowsLoadings(0).Text = Format$(scr1.WaterFlowRate.value, GetTheFormat(scr1.WaterFlowRate.value))
    
End Sub

Sub KLaOverSpecificationMessage()
Dim msg As String

  msg = "You may only specify one of these two values:" & Chr$(13) & Chr$(13)
  msg = msg + "     KLa Safety Factor" & Chr$(13)
  msg = msg + "     Design Mass Transfer Coefficient" & Chr$(13) & Chr$(13)
  msg = msg + "Either of the two values that was not just specified will be set to zero."
  MsgBox msg, MB_ICONEXCLAMATION, "Overspecification Error"

End Sub

Sub LoadContaminantList()
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

       'frmListContaminant.ListContaminants.Clear
       frmPTADScreen1!cboSelectCompo.Clear

       i = 0
       NotSpecifiedAtOperatingTemperature = False
       NotSpecifiedAtOperatingPressure = False
       Do Until EOF(1)
          i = i + 1
          Input #1, scr1.Contaminant(i).Pressure, scr1.Contaminant(i).Temperature, scr1.Contaminant(i).Name, scr1.Contaminant(i).MolecularWeight.value, scr1.Contaminant(i).HenrysConstant.value, scr1.Contaminant(i).MolarVolume.value, scr1.Contaminant(i).NormalBoilingPoint.value, scr1.Contaminant(i).LiquidDiffusivity.value, scr1.Contaminant(i).GasDiffusivity.value, scr1.Contaminant(i).Influent.value, scr1.Contaminant(i).TreatmentObjective.value
          'frmListContaminant.ListContaminants.AddItem Scr1.Contaminant(i).Name
          frmPTADScreen1!cboSelectCompo.AddItem scr1.Contaminant(i).Name

          If Not NotSpecifiedAtOperatingTemperature Then
             If Abs(scr1.Contaminant(i).Temperature - scr1.operatingtemperature.value) > TOLERANCE Then
                NotSpecifiedAtOperatingTemperature = True
             End If
          End If
          If Not NotSpecifiedAtOperatingPressure Then
             If Abs(scr1.Contaminant(i).Pressure - scr1.OperatingPressure.value) > TOLERANCE Then
                NotSpecifiedAtOperatingPressure = True
             End If
          End If

       Loop
       scr1.NumChemical = i
          
       Close #1

       'If frmListContaminant.mnuOptionsManipulateContaminant(1).Enabled = False Then
       '   frmListContaminant.mnuOptionsManipulateContaminant(1).Enabled = True
       '   frmListContaminant.mnuOptionsManipulateContaminant(3).Enabled = True
       '   frmListContaminant.mnuOptionsManipulateContaminant(4).Enabled = True
       '   frmListContaminant.mnuOptionsSave.Enabled = True
       '   frmListContaminant.mnuOptionsView.Enabled = True
       'End If

       'frmListContaminant.ListContaminants.Selected(0) = True

       If NotSpecifiedAtOperatingPressure And NotSpecifiedAtOperatingTemperature Then
          MsgBox "For one or more contaminants, the temperature and pressure at which the contaminant properties are specified differs from the operating temperature and pressure.", MB_ICONINFORMATION, "Warning"
       ElseIf NotSpecifiedAtOperatingTemperature Then
          MsgBox "For one or more contaminants, the temperature at which the contaminant properties are specified differs from the operating temperature.", MB_ICONINFORMATION, "Warning"
       ElseIf NotSpecifiedAtOperatingPressure Then
          MsgBox "For one or more contaminants, the pressure at which the contaminant properties are specified differs from the operating pressure.", MB_ICONINFORMATION, "Warning"
       End If

    End If
          
End Sub

Sub LoadFile(Filename As String)

'    'frmFileSelector.Show 1
'    On Error Resume Next
''    frmListContaminant!CMDialog1.Dir = app.path
'    frmListContaminant!CMDialog1.DefaultExt = "con"
'    frmListContaminant!CMDialog1.Filter = "Contaminant Files (*.con)|*.con"
'    frmListContaminant!CMDialog1.DialogTitle = "Load Contaminants"
'    frmListContaminant!CMDialog1.Flags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST
'    frmListContaminant!CMDialog1.Action = 1
'    Filename$ = frmListContaminant!CMDialog1.Filename
'    If Err = 32755 Then  'Cancel selected by user
'       Filename$ = ""
'    End If

End Sub

Sub LoadFileScreen1(Filename As String)
Dim Ctl As Control
Set Ctl = frmPTADScreen1.CommonDialog1
  On Error Resume Next
  'frmPTADScreen1!cmdialog1.DefaultExt = "des"
  'frmPTADScreen1!cmdialog1.Filter = "Design Files (*.des)|*.des"
  'frmPTADScreen1!cmdialog1.DialogTitle = "Load Packed Tower Aeration Design File"
  'frmPTADScreen1!cmdialog1.flags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST
  'frmPTADScreen1!cmdialog1.Action = 1
  'Filename$ = frmPTADScreen1!cmdialog1.Filename
  Ctl.DefaultExt = "des"
  Ctl.Filter = "Design Files (*.des)|*.des"
  Ctl.DialogTitle = "Load Packed Tower Aeration Design File"
  Ctl.flags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST
  Ctl.Action = 1
  Filename$ = Ctl.Filename
  If Err = 32755 Then   'Cancel selected by user
    Filename$ = ""
  End If

End Sub

Function loadscreen1(OverrideFilename As String) As Boolean
Dim FileID As String, msg As String
Dim i As Integer
Dim FoundCurrentPacking As Integer  'Whether packing user specified is currently in the user-modified database or if we have to add it when the database is the user-modified one.
Dim CurrPackingIndex As Integer
ReDim u(10) As String
Dim xu As rec_Units_frmContaminantPropertyEdit

    If (OverrideFilename <> "") Then
      Filename = OverrideFilename
    Else
      If Filename = "TheDefaultCaseScreen1" Then
        Filename = App.Path & "\dbase\default.des"
      Else
        Call LoadFileScreen1(Filename)
      End If
    End If
    
    If Filename$ <> "" Then
       FileID = ""
       If (fileexists(Filename) = False) Then
         Call Error_Unavailable_File( _
            Filename, _
            "Packed Tower Aeration Design Mode")
         loadscreen1 = False
         Exit Function
       End If
       Open Filename$ For Input As #1
       On Error Resume Next
       Input #1, FileID
       If FileID <> SCREEN1_PTAD1_FILEID Then
          msg = "Invalid Design File"
          MsgBox msg, 48, "Error"
          Close #1
          Exit Function
       End If

       frmPTADScreen1!cboSelectCompo.Clear

       Input #1, scr1.OperatingPressure.value
       frmPTADScreen1!txtOperatingPressure.Text = Format$(scr1.OperatingPressure.value * 101325# / 1#, "0.00")
       Scr2.OperatingPressure.ValChanged = True

       Input #1, scr1.operatingtemperature.value
       frmPTADScreen1!txtOperatingTemperature.Text = Format$(scr1.operatingtemperature.value - 273.15, "0.0")
       Scr2.operatingtemperature.ValChanged = True

       Call CalculateAirWaterProperties

       Input #1, scr1.Packing.Name, scr1.Packing.NominalSize, scr1.Packing.PackingFactor, scr1.Packing.SpecificSurfaceArea, scr1.Packing.CriticalSurfaceTension, scr1.Packing.Material, scr1.Packing.source, scr1.Packing.UserInput, scr1.Packing.SourceDatabase
       frmPTADScreen1!lblPackingType.Caption = scr1.Packing.Name

       If PackingDatabaseSource <> scr1.Packing.SourceDatabase Then
          frmSelectPacking!cboSelectPacking.Clear
          If scr1.Packing.SourceDatabase = ORIGINALPACKINGDATABASE Then
             frmSelectPacking!mnuPackDatabase(0).Checked = True
             frmSelectPacking!mnuPackDatabase(1).Checked = False
             frmSelectPacking!mnuPackDatabaseOptions(0).Enabled = False
          
             For i = 1 To NumPackingsInDatabase
                 frmSelectPacking!cboSelectPacking.AddItem DatabasePacking(i).Name
             Next i
             frmSelectPacking!mnuPackDatabase(3).Enabled = False
          ElseIf scr1.Packing.SourceDatabase = USERMODIFIEDPACKINGDATABASE Then
             frmSelectPacking!mnuPackDatabase(0).Checked = False
             frmSelectPacking!mnuPackDatabase(1).Checked = True
       
             For i = 1 To NumUserPackings
                 frmSelectPacking!cboSelectPacking.AddItem UserPacking(i).Name
             Next i
             frmSelectPacking!mnuPackDatabase(3).Enabled = True
          End If
       End If

       If scr1.Packing.SourceDatabase = USERMODIFIEDPACKINGDATABASE Then
             FoundCurrentPacking = False
             For i = 1 To NumUserPackings
                 If UserPacking(i).Name = scr1.Packing.Name Then
                    FoundCurrentPacking = True
                    CurrPackingIndex = i
                 End If
             Next i

             If FoundCurrentPacking Then
                If scr1.Packing.NominalSize <> UserPacking(CurrPackingIndex).NominalSize Or scr1.Packing.PackingFactor <> UserPacking(CurrPackingIndex).PackingFactor Or scr1.Packing.SpecificSurfaceArea <> UserPacking(CurrPackingIndex).SpecificSurfaceArea Or scr1.Packing.CriticalSurfaceTension <> UserPacking(CurrPackingIndex).CriticalSurfaceTension Or scr1.Packing.Material <> UserPacking(CurrPackingIndex).Material Or scr1.Packing.source <> UserPacking(CurrPackingIndex).source Then
                   msg = "Name of packing to be loaded matches the name "
                   msg = msg + "of a packing in the user-modified packing "
                   msg = msg + "database, but the properties of the two "
                   msg = msg + "packings differ." & Chr$(13) & Chr$(13)
                   msg = msg + "The properties of the packing to be loaded "
                   msg = msg + "will overwrite the properties currently "
                   msg = msg + "in the user-modified packing database."
                   MsgBox msg, MB_ICONEXCLAMATION, "Name of Packing Conflict"
                   UserPacking(CurrPackingIndex) = scr1.Packing

                End If
             End If

             If Not FoundCurrentPacking Then
                NumUserPackings = NumUserPackings + 1
                UserPacking(NumUserPackings) = scr1.Packing
                frmSelectPacking!cboSelectPacking.AddItem scr1.Packing.Name
                frmSelectPacking!cboSelectPacking.ListIndex = NumUserPackings - 1
             End If
       End If

       Input #1, scr1.NumChemical
       For i = 1 To scr1.NumChemical
           Input #1, scr1.Contaminant(i).Pressure, scr1.Contaminant(i).Temperature, scr1.Contaminant(i).Name, scr1.Contaminant(i).MolecularWeight.value, scr1.Contaminant(i).HenrysConstant.value, scr1.Contaminant(i).MolarVolume.value, scr1.Contaminant(i).NormalBoilingPoint.value, scr1.Contaminant(i).LiquidDiffusivity.value, scr1.Contaminant(i).GasDiffusivity.value, scr1.Contaminant(i).Influent.value, scr1.Contaminant(i).TreatmentObjective.value
           frmPTADScreen1!cboSelectCompo.AddItem scr1.Contaminant(i).Name
       Next i

Dim Save_DesignContaminant_Name As String
       Input #1, scr1.DesignContaminant.Name
       Save_DesignContaminant_Name = scr1.DesignContaminant.Name
      
       Call SetDesignContaminantEnabled(CInt(frmPTADScreen1!cboSelectCompo.ListCount))

       For i = 1 To scr1.NumChemical
           If Save_DesignContaminant_Name = scr1.Contaminant(i).Name Then
              scr1.DesignContaminant = scr1.Contaminant(i)
              Exit For
           End If
       Next i
       frmPTADScreen1.cboSelectCompo.ListIndex = 0
       For i = 0 To frmPTADScreen1.cboSelectCompo.ListCount
         If (frmPTADScreen1.cboSelectCompo.List(i) = Save_DesignContaminant_Name) Then
           frmPTADScreen1.cboSelectCompo.ListIndex = i
           Exit For
         End If
       Next i
       scr1.DesignContaminant.Name = Save_DesignContaminant_Name

       Input #1, scr1.WaterFlowRate.value
       ''''frmPTADScreen1!txtFlowsLoadings(0).Text = Trim$(Str$(scr1.WaterFlowRate.value))
       frmPTADScreen1!txtFlowsLoadings(0).Text = Format$(scr1.WaterFlowRate.value, GetTheFormat(scr1.WaterFlowRate.value))

       Input #1, scr1.MultipleOfMinimumAirToWaterRatio.value, scr1.MultipleOfMinimumAirToWaterRatio.UserInput
       Input #1, scr1.AirToWaterRatio.value, scr1.AirToWaterRatio.UserInput
       Input #1, scr1.AirFlowRate.value, scr1.AirFlowRate.UserInput
       If scr1.MultipleOfMinimumAirToWaterRatio.UserInput = True Then
          frmPTADScreen1!txtFlowsLoadings(2).Text = Format$(scr1.MultipleOfMinimumAirToWaterRatio.value, GetTheFormat(scr1.MultipleOfMinimumAirToWaterRatio.value))
       ElseIf scr1.AirToWaterRatio.UserInput = True Then
          frmPTADScreen1!txtFlowsLoadings(3).Text = Format$(scr1.AirToWaterRatio.value, GetTheFormat(scr1.AirToWaterRatio.value))
       ElseIf scr1.AirFlowRate.UserInput = True Then
          frmPTADScreen1!txtFlowsLoadings(4).Text = Format$(scr1.AirFlowRate.value, GetTheFormat(scr1.AirFlowRate.value))
       End If

       Input #1, scr1.AirPressureDrop.value
       frmPTADScreen1!txtFlowsLoadings(5).Text = Format$(scr1.AirPressureDrop.value, GetTheFormat(scr1.AirPressureDrop.value))

       Call GetVQmultVQAndAirFlowRate
       Call GetLoadings


       Input #1, scr1.KLaSafetyFactor.value, scr1.KLaSafetyFactor.UserInput
       Input #1, scr1.DesignMassTransferCoefficient.value, scr1.DesignMassTransferCoefficient.UserInput
       If scr1.KLaSafetyFactor.UserInput = True Then
          frmPTADScreen1!txtMassTransfer(1).Text = Format$(scr1.KLaSafetyFactor.value, GetTheFormat(scr1.KLaSafetyFactor.value))
       ElseIf scr1.DesignMassTransferCoefficient.UserInput = True Then
          frmPTADScreen1!txtMassTransfer(2).Text = Format$(scr1.DesignMassTransferCoefficient.value, GetTheFormat(scr1.DesignMassTransferCoefficient.value))
       End If

       'Input the units of this screen.
       Input #1, u(1), u(2)
       Call SetUnits(frmPTADScreen1!txtPUnits, u(1))
       Call SetUnits(frmPTADScreen1!txtTUnits, u(2))
       
       Input #1, u(1), u(2), u(3), u(4), u(5)
       Call SetUnits(frmPTADScreen1!txtFlowsUnits(0), u(1))
       Call SetUnits(frmPTADScreen1!txtFlowsUnits(4), u(2))
       Call SetUnits(frmPTADScreen1!txtFlowsUnits(5), u(3))
       Call SetUnits(frmPTADScreen1!lblFlowsUnits(6), u(4))
       Call SetUnits(frmPTADScreen1!lblFlowsUnits(7), u(5))
       
       Input #1, u(1), u(2)
       Call SetUnits(frmPTADScreen1!UnitsMassTransfer(0), u(1))
       Call SetUnits(frmPTADScreen1!UnitsMassTransfer(2), u(2))
       
       Input #1, u(1), u(2), u(3), u(4)
       Call SetUnits(frmPTADScreen1!lblTowerUnits(0), u(1))
       Call SetUnits(frmPTADScreen1!lblTowerUnits(1), u(2))
       Call SetUnits(frmPTADScreen1!lblTowerUnits(2), u(3))
       Call SetUnits(frmPTADScreen1!lblTowerUnits(3), u(4))
       
       'Input the units of frmContaminantPropertyEdit.
       xu = Units_frmContaminantPropertyEdit
       Input #1, xu.UnitsProp(0), xu.UnitsProp(2), xu.UnitsProp(3), xu.UnitsProp(4), xu.UnitsProp(5)
       Input #1, xu.UnitsConc(0), xu.UnitsConc(1)
       Units_frmContaminantPropertyEdit = xu
       
       Close #1

       Call GetTowerAreaAndDiameter
       Call GetOndaMassTransferCoefficient
       Call GetDesignKLaOrKLaSafetyFactor
       Call GetTowerHeightAndVolume

       frmPTADScreen1.Caption = "Packed Tower Aeration - Design Mode"
       If Right$(Filename, 11) = "default.des" Or Right$(Filename, 11) = "default.rat" Then
          frmPTADScreen1.Caption = frmPTADScreen1.Caption & " (" & "untitled.des" & ")"
       Else
          frmPTADScreen1.Caption = frmPTADScreen1.Caption & " (" & Filename & ")"
       End If

       'Add this file to the last-few-files list.
       Call LastFewFiles_MoveFilenameToTop(Filename)
    
    End If
    loadscreen1 = True
    
End Function

Sub NewPagePTADScreen1()

  Printer.NewPage
  Printer.FontSize = 12
  Printer.FontBold = True
  Printer.Print "Packed Tower Aeration - Design Mode (continued)"
  Printer.Print
  Printer.Print
  Printer.FontSize = 10
  Printer.FontBold = False

End Sub

Sub NumberCheck(KeyAscii As Integer)
    
  If (KeyAscii > Asc("9") Or KeyAscii < Asc("0")) And KeyAscii <> Asc(".") And KeyAscii <> 8 And KeyAscii <> Asc("E") And KeyAscii <> Asc("e") And KeyAscii <> Asc("-") Then
    KeyAscii = 0
    Beep
  End If

End Sub

Sub OptimizeDesignContaminant()
    Dim i As Integer
    Dim msg As String
    ReDim InfluentConcentrations(1 To MAXCHEMICAL) As Double
    ReDim TreatmentObjectives(1 To MAXCHEMICAL) As Double
    ReDim HenrysConstants(1 To MAXCHEMICAL) As Double
    ReDim LiquidDiffusivities(1 To MAXCHEMICAL) As Double
    ReDim GasDiffusivities(1 To MAXCHEMICAL) As Double
    ReDim EffluentConcentrations(1 To MAXCHEMICAL) As Double




       'Create the one-dimensional arrays that will be
       'passed to OPTMAL
       For i = 1 To scr1.NumChemical
           InfluentConcentrations(i) = scr1.Contaminant(i).Influent.value
           TreatmentObjectives(i) = scr1.Contaminant(i).TreatmentObjective.value
           HenrysConstants(i) = scr1.Contaminant(i).HenrysConstant.value
           LiquidDiffusivities(i) = scr1.Contaminant(i).LiquidDiffusivity.value
           GasDiffusivities(i) = scr1.Contaminant(i).GasDiffusivity.value
       Next i

       Call OPTMAL(scr1.WaterDensity.value, scr1.WaterViscosity.value, scr1.WaterSurfaceTension.value, scr1.AirDensity.value, scr1.AirViscosity.value, scr1.WaterFlowRate.value, scr1.Packing.NominalSize, scr1.Packing.PackingFactor, scr1.Packing.CriticalSurfaceTension, scr1.Packing.SpecificSurfaceArea, InfluentConcentrations(1), TreatmentObjectives(1), HenrysConstants(1), scr1.NumChemical, scr1.AirPressureDrop.value, LiquidDiffusivities(1), GasDiffusivities(1), scr1.KLaSafetyFactor.value, scr1.ID_OptimalDesignContaminant, scr1.MultipleOfMinimumAirToWaterRatio.value, EffluentConcentrations(1), ErrorFlag)

       'Copy Effluent Concentrations just calculated into
       'the Scr1 Data Structure
       For i = 1 To scr1.NumChemical
           scr1.Contaminant(i).Effluent.value = EffluentConcentrations(i)
       Next i

       frmPTADScreen1!txtFlowsLoadings(2).Text = Format$(scr1.MultipleOfMinimumAirToWaterRatio.value, GetTheFormat(scr1.MultipleOfMinimumAirToWaterRatio.value))
       scr1.MultipleOfMinimumAirToWaterRatio.UserInput = True
       scr1.AirToWaterRatio.UserInput = False
       scr1.AirFlowRate.UserInput = False


       'Show the effluent concentrations for the compounds
       'on the screen.

       If ErrorFlag = 0 Then
          Call ShowOptimizationConcentrations
       Else
          msg = "Could not converge in the optimization "
          msg = "routine.  Design contaminant chosen "
          msg = "arbitrarily between the contaminants "
          msg = "causing this to occur.  Effluent profile "
          msg = "showing these results will appear next."
          MsgBox msg, MB_ICONEXCLAMATION, "Optimization Error"
          Call ShowOptimizationConcentrations
       End If


       'Calculate values on the screen with the new design
       'contaminant

       frmPTADScreen1!cboSelectCompo.ListIndex = -1
       frmPTADScreen1!cboSelectCompo.ListIndex = scr1.ID_OptimalDesignContaminant - 1
       frmPTADScreen1!cboSelectCompo.SetFocus

End Sub

Sub PrintPTADScreen1()
    Dim CalculatedPower As Integer

    On Error GoTo PrinterError

          Printer.ScaleLeft = -1440
          Printer.ScaleTop = -1440
          Printer.CurrentX = 0
          Printer.CurrentY = 0
          Printer.FontSize = 12
          Printer.FontBold = True
          Printer.Print "Packed Tower Aeration - Design Mode"
          Printer.Print
          Printer.Print
          Printer.FontUnderline = True
          Printer.FontSize = 10
          Printer.Print "Property:"; Tab(VALUE_TAB); "Value:"
          Printer.Print
          Printer.FontUnderline = False
          Printer.FontBold = False
          Printer.Print "Operating Pressure (" & frmPTADScreen1!txtPUnits & ")"; Tab(VALUE_TAB); frmPTADScreen1!txtOperatingPressure.Text
          Printer.Print "Operating Temperature (" & frmPTADScreen1!txtTUnits & ")"; Tab(VALUE_TAB); frmPTADScreen1!txtOperatingTemperature.Text
          Printer.Print frmAirWaterProperties!lblAirWaterProperties(0).Caption; Tab(VALUE_TAB); frmAirWaterProperties!txtAirWaterProperties(0).Text
          Printer.Print frmAirWaterProperties!lblAirWaterProperties(1).Caption; Tab(VALUE_TAB); frmAirWaterProperties!txtAirWaterProperties(1).Text
          Printer.Print frmAirWaterProperties!lblAirWaterProperties(2).Caption; Tab(VALUE_TAB); frmAirWaterProperties!txtAirWaterProperties(2).Text
          Printer.Print frmAirWaterProperties!lblAirWaterProperties(3).Caption; Tab(VALUE_TAB); frmAirWaterProperties!txtAirWaterProperties(3).Text
          Printer.Print frmAirWaterProperties!lblAirWaterProperties(4).Caption; Tab(VALUE_TAB); frmAirWaterProperties!txtAirWaterProperties(4).Text
          Printer.Print
          CurrentScreen = scr1
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
          Printer.Print "Design Contaminant:  "; scr1.DesignContaminant.Name
          Printer.Print "Molecular Weight (g/gmol)"; Tab(VALUE_TAB); Format$(scr1.DesignContaminant.MolecularWeight.value, "0.00")
          Printer.Print "Henry's Constant (dimensionless)"; Tab(VALUE_TAB); Format$(scr1.DesignContaminant.HenrysConstant.value, GetTheFormat(scr1.DesignContaminant.HenrysConstant.value))
          Printer.Print "Molar Volume (m³/kmol)"; Tab(VALUE_TAB); Format$(scr1.DesignContaminant.MolarVolume.value, GetTheFormat(scr1.DesignContaminant.MolarVolume.value))
          Printer.Print "Normal Boiling Point (Celcius)"; Tab(VALUE_TAB); Format$(scr1.DesignContaminant.NormalBoilingPoint.value - 273.15, GetTheFormat(scr1.DesignContaminant.NormalBoilingPoint.value - 273.15))
          Printer.Print "Liquid Diffusivity (m²/s)"; Tab(VALUE_TAB); Format$(scr1.DesignContaminant.LiquidDiffusivity.value, GetTheFormat(scr1.DesignContaminant.LiquidDiffusivity.value))
          Printer.Print "Gas Diffusivity (m²/s)"; Tab(VALUE_TAB); Format$(scr1.DesignContaminant.GasDiffusivity.value, GetTheFormat(scr1.DesignContaminant.GasDiffusivity.value))
          Printer.Print "Influent Concentration (" & Chr$(181) & "g/L)"; Tab(VALUE_TAB); frmPTADScreen1!lblDesignConcentrationValue(0).Caption
          Printer.Print "Treatment Objective (" & Chr$(181) & "g/L)"; Tab(VALUE_TAB); frmPTADScreen1!lblDesignConcentrationValue(1).Caption
          Printer.Print "Percent Removal"; Tab(VALUE_TAB); frmPTADScreen1!lblDesignConcentrationValue(2).Caption
          Printer.Print
          Printer.Print frmPTADScreen1!lblFlowsLoadingsLabel(0).Caption & " (" & frmPTADScreen1!txtFlowsUnits(0) & ")"; Tab(VALUE_TAB); frmPTADScreen1!txtFlowsLoadings(0).Text
          Printer.Print frmPTADScreen1!lblFlowsLoadingsLabel(1).Caption & " (vol/vol)"; Tab(VALUE_TAB); frmPTADScreen1!lblFlowsLoadings(1).Caption
          Printer.Print frmPTADScreen1!lblFlowsLoadingsLabel(2).Caption & " (-)"; Tab(VALUE_TAB); frmPTADScreen1!txtFlowsLoadings(2).Text
          Printer.Print frmPTADScreen1!lblFlowsLoadingsLabel(3).Caption & " (vol/vol)"; Tab(VALUE_TAB); frmPTADScreen1!txtFlowsLoadings(3).Text
          Printer.Print frmPTADScreen1!lblFlowsLoadingsLabel(4).Caption & " (" & frmPTADScreen1!txtFlowsUnits(4) & ")"; Tab(VALUE_TAB); frmPTADScreen1!txtFlowsLoadings(4).Text
          Printer.Print frmPTADScreen1!lblFlowsLoadingsLabel(5).Caption & " (" & frmPTADScreen1!txtFlowsUnits(5) & ")"; Tab(VALUE_TAB); frmPTADScreen1!txtFlowsLoadings(5).Text
          Printer.Print frmPTADScreen1!lblFlowsLoadingsLabel(6).Caption & " (" & frmPTADScreen1!lblFlowsUnits(6) & ")"; Tab(VALUE_TAB); frmPTADScreen1!lblFlowsLoadings(6).Caption
          Printer.Print frmPTADScreen1!lblFlowsLoadingsLabel(7).Caption & " (" & frmPTADScreen1!lblFlowsUnits(7) & ")"; Tab(VALUE_TAB); frmPTADScreen1!lblFlowsLoadings(7).Caption
          Printer.Print
          Printer.Print frmShowOndaKLaProperties!lblOndaPropertiesLabel(0).Caption; Tab(VALUE_TAB); frmShowOndaKLaProperties!lblOndaProperties(0).Caption
          Printer.Print frmShowOndaKLaProperties!lblOndaPropertiesLabel(1).Caption; Tab(VALUE_TAB); frmShowOndaKLaProperties!lblOndaProperties(1).Caption
          Printer.Print frmShowOndaKLaProperties!lblOndaPropertiesLabel(2).Caption; Tab(VALUE_TAB); frmShowOndaKLaProperties!lblOndaProperties(2).Caption
          Printer.Print frmShowOndaKLaProperties!lblOndaPropertiesLabel(3).Caption; Tab(VALUE_TAB); frmShowOndaKLaProperties!lblOndaProperties(3).Caption
          Printer.Print frmShowOndaKLaProperties!lblOndaPropertiesLabel(4).Caption; Tab(VALUE_TAB); frmShowOndaKLaProperties!lblOndaProperties(4).Caption
          Printer.Print frmShowOndaKLaProperties!lblOndaPropertiesLabel(5).Caption; Tab(VALUE_TAB); frmShowOndaKLaProperties!lblOndaProperties(5).Caption
          Printer.Print frmShowOndaKLaProperties!lblOndaPropertiesLabel(6).Caption; Tab(VALUE_TAB); frmShowOndaKLaProperties!lblOndaProperties(6).Caption
          Printer.Print frmShowOndaKLaProperties!lblOndaPropertiesLabel(7).Caption; Tab(VALUE_TAB); frmShowOndaKLaProperties!lblOndaProperties(7).Caption
          Printer.Print frmShowOndaKLaProperties!lblOndaPropertiesLabel(8).Caption; Tab(VALUE_TAB); frmShowOndaKLaProperties!lblOndaProperties(8).Caption
          Printer.Print frmShowOndaKLaProperties!lblOndaPropertiesLabel(9).Caption; Tab(VALUE_TAB); frmShowOndaKLaProperties!lblOndaProperties(9).Caption
          Printer.Print frmPTADScreen1!lblMassTransferLabel(1).Caption & " (-)"; Tab(VALUE_TAB); frmPTADScreen1!txtMassTransfer(1).Text
          Printer.Print frmPTADScreen1!lblMassTransferLabel(2).Caption & " (" & frmPTADScreen1!UnitsMassTransfer(2) & ")"; Tab(VALUE_TAB); frmPTADScreen1!txtMassTransfer(2).Text
          Printer.Print
          Printer.Print frmPTADScreen1!lblTowerParametersLabel(0).Caption & " (" & frmPTADScreen1!lblTowerUnits(0) & ")"; Tab(VALUE_TAB); frmPTADScreen1!lblTowerParameters(0).Caption
          Printer.Print frmPTADScreen1!lblTowerParametersLabel(1).Caption & " (" & frmPTADScreen1!lblTowerUnits(1) & ")"; Tab(VALUE_TAB); frmPTADScreen1!lblTowerParameters(1).Caption
          Printer.Print "Conc. of Design Contaminant at Air-Water Interface (" & Chr$(181) & "g/L)"; Tab(VALUE_TAB); Format$(scr1.DesignContaminant.AirWaterInterfaceConcentration, GetTheFormat(scr1.DesignContaminant.AirWaterInterfaceConcentration))
          Printer.Print "Height of a Transfer Unit (m)"; Tab(VALUE_TAB); Format$(scr1.TransferUnitHeight, GetTheFormat(scr1.TransferUnitHeight))
          Printer.Print "Number of Transfer Units (-)"; Tab(VALUE_TAB); Format$(scr1.NumberOfTransferUnits, GetTheFormat(scr1.NumberOfTransferUnits))
          Printer.Print frmPTADScreen1!lblTowerParametersLabel(2).Caption & " (" & frmPTADScreen1!lblTowerUnits(2) & ")"; Tab(VALUE_TAB); frmPTADScreen1!lblTowerParameters(2).Caption
          Printer.Print frmPTADScreen1!lblTowerParametersLabel(3).Caption & " (" & frmPTADScreen1!lblTowerUnits(3) & ")"; Tab(VALUE_TAB); frmPTADScreen1!lblTowerParameters(3).Caption
          Call NewPagePTADScreen1
          Printer.FontBold = True
          Call SetPowerPTADScreen1(CalculatedPower)
          Printer.Print "Power Calculation:"
          Printer.FontUnderline = True
          Printer.Print
          Printer.Print "Property:"; Tab(VALUE_TAB); "Value:"
          Printer.FontBold = False
          Printer.FontUnderline = False
          Printer.Print
          Printer.Print frmPower!lblPowerLabel(0).Caption; Tab(VALUE_TAB); frmPower!txtPower(0).Text
          Printer.Print frmPower!lblPowerLabel(1).Caption; Tab(VALUE_TAB); frmPower!txtPower(1).Text
          Printer.Print frmPower!lblPowerLabel(2).Caption; Tab(VALUE_TAB); frmPower!lblPower(2).Caption
          Printer.Print frmPower!lblPowerLabel(3).Caption; Tab(VALUE_TAB); frmPower!txtPower(3).Text
          Printer.Print frmPower!lblPowerLabel(4).Caption; Tab(VALUE_TAB); frmPower!lblPower(4).Caption
          Printer.Print frmPower!lblPowerLabel(5).Caption; Tab(VALUE_TAB); frmPower!lblPower(5).Caption
          
          Printer.EndDoc

    Exit Sub

PrinterError:
    MsgBox error$(Err)
    Resume ExitPrint:

ExitPrint:

End Sub

Sub PrintPTADScreen1ToFile()
    Dim CalculatedPower As Integer

        Call GetPrintFileName(PrintFileName)
        If PrintFileName$ = "" Then Exit Sub

        Open PrintFileName$ For Output As #1

          Print #1, "Packed Tower Aeration - Design Mode"
          Print #1,
          Print #1,
          Print #1, "Property:"; Tab(VALUE_TAB); "Value:"
          Print #1,
          Print #1, "Operating Pressure (" & frmPTADScreen1!txtPUnits & ")"; Tab(VALUE_TAB); frmPTADScreen1!txtOperatingPressure.Text
          Print #1, "Operating Temperature (" & frmPTADScreen1!txtTUnits & ")"; Tab(VALUE_TAB); frmPTADScreen1!txtOperatingTemperature.Text
          Print #1, frmAirWaterProperties!lblAirWaterProperties(0).Caption; Tab(VALUE_TAB); frmAirWaterProperties!txtAirWaterProperties(0).Text
          Print #1, frmAirWaterProperties!lblAirWaterProperties(1).Caption; Tab(VALUE_TAB); frmAirWaterProperties!txtAirWaterProperties(1).Text
          Print #1, frmAirWaterProperties!lblAirWaterProperties(2).Caption; Tab(VALUE_TAB); frmAirWaterProperties!txtAirWaterProperties(2).Text
          Print #1, frmAirWaterProperties!lblAirWaterProperties(3).Caption; Tab(VALUE_TAB); frmAirWaterProperties!txtAirWaterProperties(3).Text
          Print #1, frmAirWaterProperties!lblAirWaterProperties(4).Caption; Tab(VALUE_TAB); frmAirWaterProperties!txtAirWaterProperties(4).Text
          Print #1,
          CurrentScreen = scr1
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
          Print #1, "Design Contaminant:  "; scr1.DesignContaminant.Name
          Print #1, "Molecular Weight (g/gmol)"; Tab(VALUE_TAB); Format$(scr1.DesignContaminant.MolecularWeight.value, "0.00")
          Print #1, "Henry's Constant (dimensionless)"; Tab(VALUE_TAB); Format$(scr1.DesignContaminant.HenrysConstant.value, GetTheFormat(scr1.DesignContaminant.HenrysConstant.value))
          Print #1, "Molar Volume (m³/kmol)"; Tab(VALUE_TAB); Format$(scr1.DesignContaminant.MolarVolume.value, GetTheFormat(scr1.DesignContaminant.MolarVolume.value))
          Print #1, "Normal Boiling Point (Celcius)"; Tab(VALUE_TAB); Format$(scr1.DesignContaminant.NormalBoilingPoint.value - 273.15, GetTheFormat(scr1.DesignContaminant.NormalBoilingPoint.value - 273.15))
          Print #1, "Liquid Diffusivity (m²/s)"; Tab(VALUE_TAB); Format$(scr1.DesignContaminant.LiquidDiffusivity.value, GetTheFormat(scr1.DesignContaminant.LiquidDiffusivity.value))
          Print #1, "Gas Diffusivity (m²/s)"; Tab(VALUE_TAB); Format$(scr1.DesignContaminant.GasDiffusivity.value, GetTheFormat(scr1.DesignContaminant.GasDiffusivity.value))
          Print #1, "Influent Concentration (" & Chr$(181) & "g/L)"; Tab(VALUE_TAB); frmPTADScreen1!lblDesignConcentrationValue(0).Caption
          Print #1, "Treatment Objective (" & Chr$(181) & "g/L)"; Tab(VALUE_TAB); frmPTADScreen1!lblDesignConcentrationValue(1).Caption
          Print #1, "Percent Removal"; Tab(VALUE_TAB); frmPTADScreen1!lblDesignConcentrationValue(2).Caption
          Print #1,
          Print #1, frmPTADScreen1!lblFlowsLoadingsLabel(0).Caption & " (" & frmPTADScreen1!txtFlowsUnits(0) & ")"; Tab(VALUE_TAB); frmPTADScreen1!txtFlowsLoadings(0).Text
          Print #1, frmPTADScreen1!lblFlowsLoadingsLabel(1).Caption & " (vol/vol)"; Tab(VALUE_TAB); frmPTADScreen1!lblFlowsLoadings(1).Caption
          Print #1, frmPTADScreen1!lblFlowsLoadingsLabel(2).Caption & " (-)"; Tab(VALUE_TAB); frmPTADScreen1!txtFlowsLoadings(2).Text
          Print #1, frmPTADScreen1!lblFlowsLoadingsLabel(3).Caption & " (vol/vol)"; Tab(VALUE_TAB); frmPTADScreen1!txtFlowsLoadings(3).Text
          Print #1, frmPTADScreen1!lblFlowsLoadingsLabel(4).Caption & " (" & frmPTADScreen1!txtFlowsUnits(4) & ")"; Tab(VALUE_TAB); frmPTADScreen1!txtFlowsLoadings(4).Text
          Print #1, frmPTADScreen1!lblFlowsLoadingsLabel(5).Caption & " (" & frmPTADScreen1!txtFlowsUnits(5) & ")"; Tab(VALUE_TAB); frmPTADScreen1!txtFlowsLoadings(5).Text
          Print #1, frmPTADScreen1!lblFlowsLoadingsLabel(6).Caption & " (" & frmPTADScreen1!lblFlowsUnits(6) & ")"; Tab(VALUE_TAB); frmPTADScreen1!lblFlowsLoadings(6).Caption
          Print #1, frmPTADScreen1!lblFlowsLoadingsLabel(7).Caption & " (" & frmPTADScreen1!lblFlowsUnits(7) & ")"; Tab(VALUE_TAB); frmPTADScreen1!lblFlowsLoadings(7).Caption
          Print #1,
          Print #1, frmShowOndaKLaProperties!lblOndaPropertiesLabel(0).Caption; Tab(VALUE_TAB); frmShowOndaKLaProperties!lblOndaProperties(0).Caption
          Print #1, frmShowOndaKLaProperties!lblOndaPropertiesLabel(1).Caption; Tab(VALUE_TAB); frmShowOndaKLaProperties!lblOndaProperties(1).Caption
          Print #1, frmShowOndaKLaProperties!lblOndaPropertiesLabel(2).Caption; Tab(VALUE_TAB); frmShowOndaKLaProperties!lblOndaProperties(2).Caption
          Print #1, frmShowOndaKLaProperties!lblOndaPropertiesLabel(3).Caption; Tab(VALUE_TAB); frmShowOndaKLaProperties!lblOndaProperties(3).Caption
          Print #1, frmShowOndaKLaProperties!lblOndaPropertiesLabel(4).Caption; Tab(VALUE_TAB); frmShowOndaKLaProperties!lblOndaProperties(4).Caption
          Print #1, frmShowOndaKLaProperties!lblOndaPropertiesLabel(5).Caption; Tab(VALUE_TAB); frmShowOndaKLaProperties!lblOndaProperties(5).Caption
          Print #1, frmShowOndaKLaProperties!lblOndaPropertiesLabel(6).Caption; Tab(VALUE_TAB); frmShowOndaKLaProperties!lblOndaProperties(6).Caption
          Print #1, frmShowOndaKLaProperties!lblOndaPropertiesLabel(7).Caption; Tab(VALUE_TAB); frmShowOndaKLaProperties!lblOndaProperties(7).Caption
          Print #1, frmShowOndaKLaProperties!lblOndaPropertiesLabel(8).Caption; Tab(VALUE_TAB); frmShowOndaKLaProperties!lblOndaProperties(8).Caption
          Print #1, frmShowOndaKLaProperties!lblOndaPropertiesLabel(9).Caption; Tab(VALUE_TAB); frmShowOndaKLaProperties!lblOndaProperties(9).Caption
          Print #1, frmPTADScreen1!lblMassTransferLabel(1).Caption & " (-)"; Tab(VALUE_TAB); frmPTADScreen1!txtMassTransfer(1).Text
          Print #1, frmPTADScreen1!lblMassTransferLabel(2).Caption & " (" & frmPTADScreen1!UnitsMassTransfer(2) & ")"; Tab(VALUE_TAB); frmPTADScreen1!txtMassTransfer(2).Text
          Printer.Print
          Print #1, frmPTADScreen1!lblTowerParametersLabel(0).Caption & " (" & frmPTADScreen1!lblTowerUnits(0) & ")"; Tab(VALUE_TAB); frmPTADScreen1!lblTowerParameters(0).Caption
          Print #1, frmPTADScreen1!lblTowerParametersLabel(1).Caption & " (" & frmPTADScreen1!lblTowerUnits(1) & ")"; Tab(VALUE_TAB); frmPTADScreen1!lblTowerParameters(1).Caption
          Printer.Print "Conc. of Design Contaminant at Air-Water Interface (" & Chr$(181) & "g/L)"; Tab(VALUE_TAB); Format$(scr1.DesignContaminant.AirWaterInterfaceConcentration, GetTheFormat(scr1.DesignContaminant.AirWaterInterfaceConcentration))
          Printer.Print "Height of a Transfer Unit (m)"; Tab(VALUE_TAB); Format$(scr1.TransferUnitHeight, GetTheFormat(scr1.TransferUnitHeight))
          Printer.Print "Number of Transfer Units (-)"; Tab(VALUE_TAB); Format$(scr1.NumberOfTransferUnits, GetTheFormat(scr1.NumberOfTransferUnits))
          Print #1, frmPTADScreen1!lblTowerParametersLabel(2).Caption & " (" & frmPTADScreen1!lblTowerUnits(2) & ")"; Tab(VALUE_TAB); frmPTADScreen1!lblTowerParameters(2).Caption
          Print #1, frmPTADScreen1!lblTowerParametersLabel(3).Caption & " (" & frmPTADScreen1!lblTowerUnits(3) & ")"; Tab(VALUE_TAB); frmPTADScreen1!lblTowerParameters(3).Caption
          Print #1,
          Print #1,
          Call SetPowerPTADScreen1(CalculatedPower)
          Print #1, "Power Calculation:"
          Print #1,
          Print #1, "Property:"; Tab(VALUE_TAB); "Value:"
          Print #1,
          Print #1, frmPower!lblPowerLabel(0).Caption; Tab(VALUE_TAB); frmPower!txtPower(0).Text
          Print #1, frmPower!lblPowerLabel(1).Caption; Tab(VALUE_TAB); frmPower!txtPower(1).Text
          Print #1, frmPower!lblPowerLabel(2).Caption; Tab(VALUE_TAB); frmPower!lblPower(2).Caption
          Print #1, frmPower!lblPowerLabel(3).Caption; Tab(VALUE_TAB); frmPower!txtPower(3).Text
          Print #1, frmPower!lblPowerLabel(4).Caption; Tab(VALUE_TAB); frmPower!lblPower(4).Caption
          Print #1, frmPower!lblPowerLabel(5).Caption; Tab(VALUE_TAB); frmPower!lblPower(5).Caption
          
          Close #1

    Exit Sub

End Sub

Sub ResetVariables()

'*** This subroutine will reset calculated values to 0.0 and
'*** specified values on the right side of the screen to
'*** their defaults when the user empties the contaminant
'*** list

    'Initialize value for Water Flow Rate
    'Call InitializeWaterFlowRate

    'Initialize value for Multiple of Minimum Air to Water Ratio
    Call InitializeVQminMultiple

    'Initialize Value for Air Pressure Drop
    'Call InitializeAirPressureDrop

    'Initialize Value for KLaSafetyFactor
    'Call InitializeKLaSafetyFactor
    If scr1.KLaSafetyFactor.UserInput = False Then
       scr1.KLaSafetyFactor.UserInput = True
       scr1.KLaSafetyFactor.value = 1#
       frmPTADScreen1!txtMassTransfer(1).Text = "1.0"
    End If

    'Initialize calculated properties text boxes to 0
    'and disabled
    scr1.TowerVolume.value = -1#
    Call InitializeCalculatedProperties

End Sub

Sub SaveContaminantList()
    Dim FileID As String
    Dim i As Integer

    Call SaveFile(Filename)

    If Filename$ <> "" Then
       FileID = CONTAMINANTS_PTAD_FILEID
       Open Filename$ For Output As #1
       
       Write #1, FileID
      
       For i = 1 To scr1.NumChemical
           Write #1, scr1.Contaminant(i).Pressure, scr1.Contaminant(i).Temperature, scr1.Contaminant(i).Name, scr1.Contaminant(i).MolecularWeight.value, scr1.Contaminant(i).HenrysConstant.value, scr1.Contaminant(i).MolarVolume.value, scr1.Contaminant(i).NormalBoilingPoint.value, scr1.Contaminant(i).LiquidDiffusivity.value, scr1.Contaminant(i).GasDiffusivity.value, scr1.Contaminant(i).Influent.value, scr1.Contaminant(i).TreatmentObjective.value
       Next i

       Close #1

    End If


End Sub

Sub SaveFile(Filename As String)
    
    'On Error Resume Next
    'frmListContaminant!CMDialog1.DefaultExt = "con"
    'frmListContaminant!CMDialog1.Filter = "Contaminant Files (*.con)|*.con"
    'frmListContaminant!CMDialog1.DialogTitle = "Save Contaminants"
    'frmListContaminant!CMDialog1.Flags = OFN_OVERWRITEPROMPT Or OFN_PATHMUSTEXIST
    'frmListContaminant!CMDialog1.Action = 2
    'Filename$ = frmListContaminant!CMDialog1.Filename
    'If Err = 32755 Then   'Cancel selected by user
    '   Filename$ = ""
    'End If

End Sub

Sub savefilescreen1(Filename As String)
Dim Ctl As Control
Set Ctl = frmPTADScreen1.CommonDialog1

    On Error Resume Next
    'frmPTADScreen1!cmdialog1.DefaultExt = "des"
    'frmPTADScreen1!cmdialog1.Filter = "Design Files (*.des)|*.des"
    'frmPTADScreen1!cmdialog1.DialogTitle = "Save Packed Tower Aeration Design File"
    'frmPTADScreen1!cmdialog1.flags = OFN_OVERWRITEPROMPT Or OFN_PATHMUSTEXIST
    'frmPTADScreen1!cmdialog1.Action = 2
    'Filename$ = frmPTADScreen1!cmdialog1.Filename
    Ctl.DefaultExt = "des"
    Ctl.Filter = "Design Files (*.des)|*.des"
    Ctl.DialogTitle = "Save Packed Tower Aeration Design File"
    Ctl.flags = OFN_OVERWRITEPROMPT Or OFN_PATHMUSTEXIST
    Ctl.Action = 2
    Filename$ = Ctl.Filename
    If Err = 32755 Then   'Cancel selected by user
       Filename$ = ""
    End If

End Sub

Sub SaveScreen1()
Dim FileID As String
Dim i As Integer
Dim xu As rec_Units_frmContaminantPropertyEdit

  If (IsThisADemo() = True) Then
    Call Demo_ShowError("Saving is not allowed in the demonstration version.")
    Exit Sub
  End If
    
    If Right$(frmPTADScreen1.Caption, 14) = "(untitled.des)" Then
       Call savefilescreen1(Filename)
    End If

    If Filename$ <> "" Then
       FileID = SCREEN1_PTAD1_FILEID
       Open Filename$ For Output As #1
       
       Write #1, FileID
      
       Write #1, scr1.OperatingPressure.value
       Write #1, scr1.operatingtemperature.value
       Write #1, scr1.Packing.Name, scr1.Packing.NominalSize, scr1.Packing.PackingFactor, scr1.Packing.SpecificSurfaceArea, scr1.Packing.CriticalSurfaceTension, scr1.Packing.Material, scr1.Packing.source, scr1.Packing.UserInput, scr1.Packing.SourceDatabase

       Write #1, scr1.NumChemical
       For i = 1 To scr1.NumChemical
           Write #1, scr1.Contaminant(i).Pressure, scr1.Contaminant(i).Temperature, scr1.Contaminant(i).Name, scr1.Contaminant(i).MolecularWeight.value, scr1.Contaminant(i).HenrysConstant.value, scr1.Contaminant(i).MolarVolume.value, scr1.Contaminant(i).NormalBoilingPoint.value, scr1.Contaminant(i).LiquidDiffusivity.value, scr1.Contaminant(i).GasDiffusivity.value, scr1.Contaminant(i).Influent.value, scr1.Contaminant(i).TreatmentObjective.value
       Next i
       Write #1, scr1.DesignContaminant.Name

       Write #1, scr1.WaterFlowRate.value
       Write #1, scr1.MultipleOfMinimumAirToWaterRatio.value, scr1.MultipleOfMinimumAirToWaterRatio.UserInput
       Write #1, scr1.AirToWaterRatio.value, scr1.AirToWaterRatio.UserInput
       Write #1, scr1.AirFlowRate.value, scr1.AirFlowRate.UserInput
       Write #1, scr1.AirPressureDrop.value

       Write #1, scr1.KLaSafetyFactor.value, scr1.KLaSafetyFactor.UserInput
       Write #1, scr1.DesignMassTransferCoefficient.value, scr1.DesignMassTransferCoefficient.UserInput

       'Output the units of this screen.
       Write #1, GetUnits(frmPTADScreen1!txtPUnits), GetUnits(frmPTADScreen1!txtTUnits)
       Write #1, GetUnits(frmPTADScreen1!txtFlowsUnits(0)), GetUnits(frmPTADScreen1!txtFlowsUnits(4)), GetUnits(frmPTADScreen1!txtFlowsUnits(5)), GetUnits(frmPTADScreen1!lblFlowsUnits(6)), GetUnits(frmPTADScreen1!lblFlowsUnits(7))
       Write #1, GetUnits(frmPTADScreen1!UnitsMassTransfer(0)), GetUnits(frmPTADScreen1!UnitsMassTransfer(2))
       Write #1, GetUnits(frmPTADScreen1!lblTowerUnits(0)), GetUnits(frmPTADScreen1!lblTowerUnits(1)), GetUnits(frmPTADScreen1!lblTowerUnits(2)), GetUnits(frmPTADScreen1!lblTowerUnits(3))
       
       'Output the units of frmContaminantPropertyEdit.
       xu = Units_frmContaminantPropertyEdit
       Write #1, xu.UnitsProp(0), xu.UnitsProp(2), xu.UnitsProp(3), xu.UnitsProp(4), xu.UnitsProp(5)
       Write #1, xu.UnitsConc(0), xu.UnitsConc(1)
       
       Close #1

       frmPTADScreen1.Caption = "Packed Tower Aeration - Design Mode"
       frmPTADScreen1.Caption = frmPTADScreen1.Caption & " (" & Filename & ")"

    End If

End Sub

Sub screen1_results()
    Dim i As Integer, j As Integer
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
        Dim ContaminantGlossaryBottom As Integer, GlossaryBottom As Integer
    ReDim DesiredPercentRemoval(1 To MAXCHEMICAL) As Double
    ReDim Effluent(1 To MAXCHEMICAL) As Double
    ReDim AchievedPercentRemoval(1 To MAXCHEMICAL) As Double
                                
          '----- View All Concentration Results
         KLaSafetyFactor = scr1.KLaSafetyFactor.value
          For i = 1 To scr1.NumChemical
              If scr1.DesignContaminant.Name = scr1.Contaminant(i).Name Then
                 PackingWettedSurfaceArea(i) = scr1.Packing.OndaWettedSurfaceArea
                 OndaKLa(i) = scr1.Onda.OverallMassTransferCoefficient
                 DesignKLa(i) = scr1.DesignMassTransferCoefficient.value
                 Call REMOVPT(DesiredPercentRemoval(i), scr1.DesignContaminant.Influent.value, scr1.DesignContaminant.TreatmentObjective.value)
                 Effluent(i) = scr1.DesignContaminant.TreatmentObjective.value
                 Call REMOVPT(AchievedPercentRemoval(i), scr1.DesignContaminant.Influent.value, Effluent(i))
              Else
                 Call AWCALC(PackingWettedSurfaceArea(i), scr1.Packing.CriticalSurfaceTension, scr1.WaterSurfaceTension.value, scr1.WaterLoadingRate.value, scr1.Packing.SpecificSurfaceArea, scr1.WaterViscosity.value, scr1.WaterDensity.value, ReynoldsNumber, FroudeNumber, WeberNumber)
                 Call ONDAKLPT(LiquidPhaseMassTransferCoefficient, scr1.WaterLoadingRate.value, PackingWettedSurfaceArea(i), scr1.WaterViscosity.value, scr1.WaterDensity.value, scr1.Contaminant(i).LiquidDiffusivity.value, scr1.Packing.SpecificSurfaceArea, scr1.Packing.NominalSize)
                 Call ONDAKGPT(GasPhaseMassTransferCoefficient, scr1.AirLoadingRate.value, scr1.Packing.SpecificSurfaceArea, scr1.AirViscosity.value, scr1.AirDensity.value, scr1.Contaminant(i).GasDiffusivity.value, scr1.Packing.NominalSize)
                 Call ONDKLAPT(OndaKLa(i), LiquidPhaseMassTransferResistance, GasPhaseMassTransferResistance, TotalMassTransferResistance, LiquidPhaseMassTransferCoefficient, PackingWettedSurfaceArea(i), GasPhaseMassTransferCoefficient, scr1.Contaminant(i).HenrysConstant.value)
                 Call KLACOR(DesignKLa(i), OndaKLa(i), KLaSafetyFactor)
                 Call REMOVPT(DesiredPercentRemoval(i), scr1.Contaminant(i).Influent.value, scr1.Contaminant(i).TreatmentObjective.value)
                 Call EFFLPT2(Effluent(i), scr1.AirToWaterRatio.value, scr1.Contaminant(i).HenrysConstant.value, scr1.WaterFlowRate.value, scr1.TowerArea.value, scr1.TowerHeight.value, DesignKLa(i), scr1.Contaminant(i).Influent.value)
                 Call REMOVPT(AchievedPercentRemoval(i), scr1.Contaminant(i).Influent.value, Effluent(i))
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

    For i = 1 To scr1.NumChemical
        frmViewEffluentConcentrationsASAP!lblContaminantNumber(i - 1).Visible = True
        frmViewEffluentConcentrationsASAP!lblContaminantNumber(i + 10 - 1).Visible = True
        frmViewEffluentConcentrationsASAP!lblInfluentConcentration(i - 1).Visible = True
        frmViewEffluentConcentrationsASAP!lblTreatmentObjective(i - 1).Visible = True
        frmViewEffluentConcentrationsASAP!lblDesiredPercentRemoval(i - 1).Visible = True
        frmViewEffluentConcentrationsASAP!lblEffluentConcentration(i - 1).Visible = True
        frmViewEffluentConcentrationsASAP!lblAchievedPercentRemoval(i - 1).Visible = True
        frmViewEffluentConcentrationsASAP!lblContaminantName(i - 1).Visible = True

        frmViewEffluentConcentrationsASAP!lblInfluentConcentration(i - 1).Caption = Format$(scr1.Contaminant(i).Influent.value, GetTheFormat(scr1.Contaminant(i).Influent.value))
        frmViewEffluentConcentrationsASAP!lblTreatmentObjective(i - 1).Caption = Format$(scr1.Contaminant(i).TreatmentObjective.value, GetTheFormat(scr1.Contaminant(i).TreatmentObjective.value))
        frmViewEffluentConcentrationsASAP!lblDesiredPercentRemoval(i - 1).Caption = Format$(DesiredPercentRemoval(i), "0.0")
        frmViewEffluentConcentrationsASAP!lblEffluentConcentration(i - 1).Caption = Format$(Effluent(i), GetTheFormat(Effluent(i)))
        frmViewEffluentConcentrationsASAP!lblAchievedPercentRemoval(i - 1).Caption = Format$(AchievedPercentRemoval(i), "0.0")
        frmViewEffluentConcentrationsASAP!lblContaminantName(i - 1).Caption = Trim$(LCase$(scr1.Contaminant(i).Name))

    Next i

    frmViewEffluentConcentrationsASAP!fraConcentrationResults.Height = frmViewEffluentConcentrationsASAP!lblContaminantNumber(scr1.NumChemical - 1).Top + frmViewEffluentConcentrationsASAP!lblContaminantNumber(scr1.NumChemical - 1).Height + 120
    frmViewEffluentConcentrationsASAP!fraContaminantGlossary.Height = frmViewEffluentConcentrationsASAP!lblContaminantNumber(scr1.NumChemical + 10 - 1).Top + frmViewEffluentConcentrationsASAP!lblContaminantNumber(scr1.NumChemical + 10 - 1).Height + 120
    frmViewEffluentConcentrationsASAP!fraContaminantGlossary.Top = frmViewEffluentConcentrationsASAP!fraConcentrationResults.Top + frmViewEffluentConcentrationsASAP!fraConcentrationResults.Height + 120
    frmViewEffluentConcentrationsASAP!fraGlossary.Top = frmViewEffluentConcentrationsASAP!fraContaminantGlossary.Top
    ContaminantGlossaryBottom = frmViewEffluentConcentrationsASAP!fraContaminantGlossary.Top + frmViewEffluentConcentrationsASAP!fraContaminantGlossary.Height
    GlossaryBottom = frmViewEffluentConcentrationsASAP!fraGlossary.Top + frmViewEffluentConcentrationsASAP!fraGlossary.Height
    If GlossaryBottom > ContaminantGlossaryBottom Then
       frmViewEffluentConcentrationsASAP!cmdOK.Top = GlossaryBottom + 360
    Else
       frmViewEffluentConcentrationsASAP!cmdOK.Top = ContaminantGlossaryBottom + 360
    End If
    frmViewEffluentConcentrationsASAP.Height = frmViewEffluentConcentrationsASAP!cmdOK.Top + frmViewEffluentConcentrationsASAP!cmdOK.Height + 500  ' 420

    frmViewEffluentConcentrationsASAP.Show 1

End Sub

Function screen1_savechanges() As Integer
Dim i As Integer, Response As Integer
Dim msg As String
                
msg = "Would you like to save the parameters "
msg = msg + "for this design case to a file "
msg = msg + "?" & Chr$(13) & Chr$(13)
msg = msg + "Note:  Any information not saved will be permanently lost."
Response = MsgBox(msg, MB_ICONquestion + MB_YESNOCANCEL, "Save Current Design")
                
If Response = IDCANCEL Then
  Screen.MousePointer = 0
  screen1_savechanges = 1
  Exit Function
End If
                
If Response = IDYES Then
 Call SaveScreen1
 If StrComp(Filename, "") = 0 Then Response = 5
                      
 Do While Response = 5
    msg = "Would you like to save the parameters "
    msg = msg + "for this design case to a file "
    msg = msg + "?" & Chr$(13) & Chr$(13)
    msg = msg + "Note:  Any information not saved will be permanently lost."
    Response = MsgBox(msg, MB_ICONquestion + MB_YESNOCANCEL, "Save Current Design")
                         
    If Response = IDCANCEL Then
       Screen.MousePointer = 0
       screen1_savechanges = 1
       Exit Function
    End If
                         
   If Response = IDYES Then Call SaveScreen1
   If StrComp(Filename, "") = 0 And Response <> IDNO Then Response = 5
  Loop
End If

screen1_savechanges = 0

End Function

Sub SetDesignContaminantEnabled(NumInList As Integer)
    Dim i As Integer

  If NumInList = 0 Then
    frmPTADScreen1!mnuFile(4).Enabled = False
    frmPTADScreen1!mnuFile(5).Enabled = False
    frmPTADScreen1!mnuOptions(0).Enabled = False
    'frmPTADScreen1!fraDesignContaminant.Enabled = False
    'frmPTADScreen1!cboDesignContaminant.Enabled = False
    frmPTADScreen1!cmdDesignContaminant.Enabled = False
    For i = 0 To 2
      frmPTADScreen1!lblDesignConcentrationValue(i).Caption = ""
    Next i
    Call ResetVariables
  Else
    frmPTADScreen1!mnuFile(4).Enabled = True
    frmPTADScreen1!mnuFile(5).Enabled = True
    frmPTADScreen1!mnuOptions(0).Enabled = True
    'frmPTADScreen1!fraDesignContaminant.Enabled = True
    'frmPTADScreen1!cboDesignContaminant.Enabled = True
    If NumInList = 1 Then
      frmPTADScreen1!cmdDesignContaminant.Enabled = False
    Else
      frmPTADScreen1!cmdDesignContaminant.Enabled = True
    End If
    For i = 0 To 0
      frmPTADScreen1!txtFlowsLoadings(i).Enabled = True
    Next i
    For i = 1 To 1
      frmPTADScreen1!lblFlowsLoadings(i).Enabled = True
    Next i
    For i = 2 To 5
      frmPTADScreen1!txtFlowsLoadings(i).Enabled = True
    Next i
    For i = 6 To 7
      frmPTADScreen1!lblFlowsLoadings(i).Enabled = True
    Next i

    frmPTADScreen1!lblMassTransfer(0).Enabled = True
    For i = 1 To 2
      frmPTADScreen1!txtMassTransfer(i).Enabled = True
    Next i
    For i = 0 To 3
      frmPTADScreen1!lblTowerParameters(i).Enabled = True
    Next i
  End If
  Call frmPTADScreen1.LOCAL___Reset_DemoVersionDisablings
End Sub

Sub SetPowerPTADScreen1(CalculatedPower As Integer)

          scr1.Power.InletAirTemperature = scr1.operatingtemperature.value - 273.15
          Call CalculatePowerScreen1(CalculatedPower)
          If CalculatedPower Then
             frmPower!txtPower(0).Text = Format$(scr1.Power.InletAirTemperature, GetTheFormat(scr1.Power.InletAirTemperature))
             frmPower!txtPower(1).Text = Format$(scr1.Power.BlowerEfficiency, GetTheFormat(scr1.Power.BlowerEfficiency))
             frmPower!lblPower(2).Caption = Format$(scr1.Power.BlowerBrakePower, GetTheFormat(scr1.Power.BlowerBrakePower))
             frmPower!txtPower(3).Text = Format$(scr1.Power.PumpEfficiency, GetTheFormat(scr1.Power.PumpEfficiency))
             frmPower!lblPower(4).Caption = Format$(scr1.Power.PumpBrakePower, GetTheFormat(scr1.Power.PumpBrakePower))
             frmPower!lblPower(5).Caption = Format$(scr1.Power.TotalBrakePower, GetTheFormat(scr1.Power.TotalBrakePower))
          End If

End Sub

Sub ShowOndaKLaProperties()
    
  frmShowOndaKLaProperties!lblOndaProperties(0).Caption = Format$(scr1.Onda.ReynoldsNumber, GetTheFormat(scr1.Onda.ReynoldsNumber))
  frmShowOndaKLaProperties!lblOndaProperties(1).Caption = Format$(scr1.Onda.FroudeNumber, GetTheFormat(scr1.Onda.FroudeNumber))
  frmShowOndaKLaProperties!lblOndaProperties(2).Caption = Format$(scr1.Onda.WeberNumber, GetTheFormat(scr1.Onda.WeberNumber))
  frmShowOndaKLaProperties!lblOndaProperties(3).Caption = Format$(scr1.Packing.OndaWettedSurfaceArea, GetTheFormat(scr1.Packing.OndaWettedSurfaceArea))
  frmShowOndaKLaProperties!lblOndaProperties(4).Caption = Format$(scr1.Onda.LiquidPhaseMassTransferResistance, GetTheFormat(scr1.Onda.LiquidPhaseMassTransferResistance))
  frmShowOndaKLaProperties!lblOndaProperties(5).Caption = Format$(scr1.Onda.GasPhaseMassTransferResistance, GetTheFormat(scr1.Onda.GasPhaseMassTransferResistance))
  frmShowOndaKLaProperties!lblOndaProperties(6).Caption = Format$(scr1.Onda.TotalMassTransferResistance, GetTheFormat(scr1.Onda.TotalMassTransferResistance))
  frmShowOndaKLaProperties!lblOndaProperties(7).Caption = Format$(scr1.Onda.LiquidPhaseMassTransferCoefficient, GetTheFormat(scr1.Onda.LiquidPhaseMassTransferCoefficient))
  frmShowOndaKLaProperties!lblOndaProperties(8).Caption = Format$(scr1.Onda.GasPhaseMassTransferCoefficient, GetTheFormat(scr1.Onda.GasPhaseMassTransferCoefficient))
  frmShowOndaKLaProperties!lblOndaProperties(9).Caption = Format$(scr1.Onda.OverallMassTransferCoefficient, GetTheFormat(scr1.Onda.OverallMassTransferCoefficient))

End Sub

Sub ShowOptimizationConcentrations()
Dim i As Integer, Tag As Integer

  Tag = scr1.ID_OptimalDesignContaminant
  frmOptimizeContaminant!lblDesignContaminant(0).Caption = scr1.Contaminant(Tag).Name
  frmOptimizeContaminant!lblDesignContaminant(1).Caption = Format$(scr1.Contaminant(Tag).Influent.value, GetTheFormat(scr1.Contaminant(Tag).Influent.value))
  frmOptimizeContaminant!lblDesignContaminant(2).Caption = Format$(scr1.Contaminant(Tag).TreatmentObjective.value, GetTheFormat(scr1.Contaminant(Tag).TreatmentObjective.value))
  frmOptimizeContaminant!lblDesignContaminant(3).Caption = Format$(scr1.Contaminant(Tag).Effluent.value, GetTheFormat(scr1.Contaminant(Tag).Effluent.value))

  frmOptimizeContaminant!lstOptimizeContaminant.Clear

  For i = 1 To (Tag - 1)
    frmOptimizeContaminant!lstOptimizeContaminant.AddItem scr1.Contaminant(i).Name
  Next i

  For i = (Tag + 1) To scr1.NumChemical
    frmOptimizeContaminant!lstOptimizeContaminant.AddItem scr1.Contaminant(i).Name
  Next i

  frmOptimizeContaminant!lstOptimizeContaminant.ListIndex = 0

  frmOptimizeContaminant.Show 1

End Sub

Sub ShowPackingProperties()

    PackingDatabaseSource = CurrentScreen.Packing.SourceDatabase

    If ScreenNumber = 1 Then
       If frmPTADScreen1.lblPackingType.Caption = "" Then Exit Sub
    ElseIf ScreenNumber = 2 Then
       If frmPTADScreen2.lblPackingType.Caption = "" Then Exit Sub
    End If

    If Not ShownPackingProperties Then   'Set labels on frmShowPackingProperties
       frmShowPackingProperties!lblShowPackingProperties(0).Caption = CurrentScreen.Packing.Name
       frmShowPackingProperties!lblShowPackingProperties(1).Caption = Format$(CurrentScreen.Packing.NominalSize, GetTheFormat(CurrentScreen.Packing.NominalSize))
       frmShowPackingProperties!lblShowPackingProperties(2).Caption = Format$(CurrentScreen.Packing.PackingFactor, GetTheFormat(CurrentScreen.Packing.PackingFactor))
       frmShowPackingProperties!lblShowPackingProperties(3).Caption = Format$(CurrentScreen.Packing.SpecificSurfaceArea, GetTheFormat(CurrentScreen.Packing.SpecificSurfaceArea))
       frmShowPackingProperties!lblShowPackingProperties(4).Caption = Format$(CurrentScreen.Packing.CriticalSurfaceTension, GetTheFormat(CurrentScreen.Packing.CriticalSurfaceTension))
       frmShowPackingProperties!lblShowPackingProperties(5).Caption = CurrentScreen.Packing.Material
       frmShowPackingProperties!lblShowPackingProperties(6).Caption = CurrentScreen.Packing.source

       If PackingDatabaseSource = ORIGINALPACKINGDATABASE Then
          frmShowPackingProperties!lblShowPackingProperties(7).Caption = "Original Packing Database"
       Else
          frmShowPackingProperties.lblShowPackingProperties(7).Caption = "User Input"
       End If

       ShownPackingProperties = True

    End If

    frmShowPackingProperties.Show 1

End Sub

Sub SpecifiedAirFlowRate()

    If HaveValue(scr1.WaterFlowRate.value) Then
       Call VQCALC(scr1.AirToWaterRatio.value, scr1.AirFlowRate.value, scr1.WaterFlowRate.value)
       scr1.AirToWaterRatio.ValChanged = True
       frmPTADScreen1!txtFlowsLoadings(3).Text = Format$(scr1.AirToWaterRatio.value, GetTheFormat(scr1.AirToWaterRatio.value))
       If HaveValue(scr1.MinimumAirToWaterRatio.value) Then
          Call GETMULT(scr1.MultipleOfMinimumAirToWaterRatio.value, scr1.AirToWaterRatio.value, scr1.MinimumAirToWaterRatio.value)
          scr1.MultipleOfMinimumAirToWaterRatio.ValChanged = True
          frmPTADScreen1!txtFlowsLoadings(2).Text = Format$(scr1.MultipleOfMinimumAirToWaterRatio.value, GetTheFormat(scr1.MultipleOfMinimumAirToWaterRatio.value))
       Else
          If (scr1.MultipleOfMinimumAirToWaterRatio.value > 0#) Then
             Call VQOverSpecificationMessage
             scr1.MultipleOfMinimumAirToWaterRatio.value = 0#
             frmPTADScreen1!txtFlowsLoadings(2).Text = "0.0"
          End If
       End If
    Else
       If scr1.MultipleOfMinimumAirToWaterRatio.value > 0# Or scr1.AirToWaterRatio.value > 0# Then
          Call VQOverSpecificationMessage
       End If
       If (scr1.MultipleOfMinimumAirToWaterRatio.value > 0#) Then
          scr1.MultipleOfMinimumAirToWaterRatio.value = 0#
          frmPTADScreen1!txtFlowsLoadings(2).Text = "0.0"
       End If
       If scr1.AirToWaterRatio.value > 0# Then
          scr1.AirToWaterRatio.value = 0#
          frmPTADScreen1!txtFlowsLoadings(3).Text = "0.0"
       End If

    End If

    scr1.MultipleOfMinimumAirToWaterRatio.UserInput = False
    scr1.AirToWaterRatio.UserInput = False

End Sub

Sub SpecifiedAirToWaterRatio()

    If HaveValue(scr1.MinimumAirToWaterRatio.value) Then
       Call GETMULT(scr1.MultipleOfMinimumAirToWaterRatio.value, scr1.AirToWaterRatio.value, scr1.MinimumAirToWaterRatio.value)
       scr1.MultipleOfMinimumAirToWaterRatio.ValChanged = True
       frmPTADScreen1!txtFlowsLoadings(2).Text = Format$(scr1.MultipleOfMinimumAirToWaterRatio.value, GetTheFormat(scr1.MultipleOfMinimumAirToWaterRatio.value))
       If HaveValue(scr1.WaterFlowRate.value) Then
          Call AIRFLO(scr1.AirFlowRate.value, scr1.AirToWaterRatio.value, scr1.WaterFlowRate.value)
          scr1.AirFlowRate.ValChanged = True
          frmPTADScreen1!txtFlowsLoadings(4).Text = Format$(scr1.AirFlowRate.value, GetTheFormat(scr1.AirFlowRate.value))
          Call Units_DoRefresh(frmPTADScreen1.txtFlowsUnits(4))
       ElseIf scr1.AirFlowRate.value > 0# Then
          Call VQOverSpecificationMessage
          scr1.AirFlowRate.value = 0#
          frmPTADScreen1!txtFlowsLoadings(4).Text = "0.0"
          Call Units_DoRefresh(frmPTADScreen1.txtFlowsUnits(4))
       End If
    Else
       If (scr1.MultipleOfMinimumAirToWaterRatio.value > 0#) Or scr1.AirFlowRate.value > 0# Then
          Call VQOverSpecificationMessage
       End If
       If (scr1.MultipleOfMinimumAirToWaterRatio.value > 0#) Then
          scr1.MultipleOfMinimumAirToWaterRatio.value = 0#
          frmPTADScreen1!txtFlowsLoadings(2).Text = "0.0"
       End If
       If scr1.AirFlowRate.value > 0# Then
          scr1.AirFlowRate.value = 0#
          frmPTADScreen1!txtFlowsLoadings(4).Text = "0.0"
          Call Units_DoRefresh(frmPTADScreen1.txtFlowsUnits(4))
       End If

    End If

    scr1.MultipleOfMinimumAirToWaterRatio.UserInput = False
    scr1.AirFlowRate.UserInput = False

End Sub

Sub SpecifiedDesignMassTransferCoefficient()

    If HaveValue(scr1.Onda.OverallMassTransferCoefficient) Then
       Call GETSAF(scr1.KLaSafetyFactor.value, scr1.Onda.OverallMassTransferCoefficient, scr1.DesignMassTransferCoefficient.value)
       scr1.KLaSafetyFactor.ValChanged = True
       frmPTADScreen1!txtMassTransfer(1).Text = Format$(scr1.KLaSafetyFactor.value, GetTheFormat(scr1.KLaSafetyFactor.value))
    ElseIf scr1.KLaSafetyFactor.value > 0# Then
       Call KLaOverSpecificationMessage
       scr1.KLaSafetyFactor.value = 0#
       frmPTADScreen1!txtMassTransfer(1).Text = "0.0"
    End If
    scr1.KLaSafetyFactor.UserInput = False

End Sub

Sub SpecifiedKLaSafetyFactor()

    If HaveValue(scr1.Onda.OverallMassTransferCoefficient) Then
       Call KLACOR(scr1.DesignMassTransferCoefficient.value, scr1.Onda.OverallMassTransferCoefficient, scr1.KLaSafetyFactor.value)
       scr1.DesignMassTransferCoefficient.ValChanged = True
       frmPTADScreen1!txtMassTransfer(2).Text = Format$(scr1.DesignMassTransferCoefficient.value, GetTheFormat(scr1.DesignMassTransferCoefficient.value))
    ElseIf scr1.DesignMassTransferCoefficient.value > 0# Then
       Call KLaOverSpecificationMessage
       scr1.DesignMassTransferCoefficient.value = 0#
       frmPTADScreen1!txtMassTransfer(2).Text = "0.0"
    End If
    scr1.DesignMassTransferCoefficient.UserInput = False

End Sub

Sub SpecifiedVQminMultiple()

'*** This subroutine will calculate air to water ratio
'*** and air flow rate when the user has input a value
'*** for multiple of minimum air to water ratio

    If HaveValue(scr1.MinimumAirToWaterRatio.value) Then
       Call vqmltpt1(scr1.AirToWaterRatio.value, scr1.MinimumAirToWaterRatio.value, scr1.MultipleOfMinimumAirToWaterRatio.value)
       scr1.AirToWaterRatio.ValChanged = True
       frmPTADScreen1!txtFlowsLoadings(3).Text = Format$(scr1.AirToWaterRatio.value, GetTheFormat(scr1.AirToWaterRatio.value))
       If HaveValue(scr1.WaterFlowRate.value) Then
          Call AIRFLO(scr1.AirFlowRate.value, scr1.AirToWaterRatio.value, scr1.WaterFlowRate.value)
          scr1.AirFlowRate.ValChanged = True
          frmPTADScreen1!txtFlowsLoadings(4).Text = Format$(scr1.AirFlowRate.value, GetTheFormat(scr1.AirFlowRate.value))
          Call Units_DoRefresh(frmPTADScreen1.txtFlowsUnits(4))
       ElseIf scr1.AirFlowRate.value > 0# Then
          Call VQOverSpecificationMessage
          scr1.AirFlowRate.value = 0#
          frmPTADScreen1!txtFlowsLoadings(4).Text = "0.0"
          Call Units_DoRefresh(frmPTADScreen1.txtFlowsUnits(4))
       End If
    Else
       If (scr1.AirToWaterRatio.value > 0#) Or (scr1.AirFlowRate.value > 0#) Then
          Call VQOverSpecificationMessage
       End If
       If (scr1.AirToWaterRatio.value > 0#) Then
          scr1.AirToWaterRatio.value = 0#
          frmPTADScreen1!txtFlowsLoadings(3).Text = "0.0"
       End If
       If scr1.AirFlowRate.value > 0# Then
          scr1.AirFlowRate.value = 0#
          frmPTADScreen1!txtFlowsLoadings(4).Text = "0.0"
          Call Units_DoRefresh(frmPTADScreen1.txtFlowsUnits(4))
       End If
    End If

    scr1.AirToWaterRatio.UserInput = False
    scr1.AirFlowRate.UserInput = False

End Sub

Function StartScreen1DefaultCase() As Boolean

    Filename = "TheDefaultCaseScreen1"
    StartScreen1DefaultCase = loadscreen1("")
    
End Function

Sub TextGetFocus(txt As TextBox, Temp_Text As String)
    Temp_Text = txt.Text
    txt.SelStart = 0
    txt.SelLength = Len(txt.Text)

End Sub

Sub TextNumberChanged(ValueChanged As Integer, txt As TextBox, Temp_Text As String)
    Dim Dummy1 As Double, Dummy2 As Double

    If Temp_Text = "" Then
       ValueChanged = True
       Exit Sub
    End If

    Dummy1 = CDbl(txt.Text)
    Dummy2 = CDbl(Temp_Text)
    ValueChanged = True
    If txt.Text = Temp_Text Then ValueChanged = False
    If Abs(Dummy1 - Dummy2) < NUMBER_CHANGING_CRITERIA Then ValueChanged = False

End Sub

Sub TextStringChanged(ValueChanged As Integer, txt As TextBox, Temp_Text As String)
    
    ValueChanged = True
    If txt.Text = Temp_Text Then ValueChanged = False

End Sub

Sub VQOverSpecificationMessage()
    Dim msg As String

    msg = "You may only specify one of these three values:" & Chr$(13) & Chr$(13)
    msg = msg + "     Multiple of Minimum Air To Water Ratio" & Chr$(13)
    msg = msg + "     Air To Water Ratio" & Chr$(13)
    msg = msg + "     Air Flow Rate" & Chr$(13) & Chr$(13)
    msg = msg + "Any of the three values that were not just specified will be set to zero."
    MsgBox msg, MB_ICONEXCLAMATION, "Overspecification Error"

End Sub

