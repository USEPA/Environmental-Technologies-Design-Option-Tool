Attribute VB_Name = "Model_FaVOr"
Option Explicit

'Global Const ModelFAVOR_IN_PathFile = "BFILM1.IN"
Global Const ModelFAVOR_IN_Main1 = "INPUT1.DAT"
Global Const ModelFAVOR_IN_Main2 = "INPUT2.DAT"
'Global Const ModelFAVOR_OUT_SuccessFlag = "BFILM1.OUT"
Global Const ModelFAVOR_OUT_Main1 = "VOC.OUT"
Global Const ModelFAVOR_OUT_Main2 = "OUTPUT.TXT"
'Global Const ModelFAVOR_OUT_CvsT = "BFILM3.OUT"
'Global Const ModelFAVOR_OUT_EndCvsT = "BFILM4.OUT"
'Global Const ModelFAVOR_OUT_CvsT_Details = "BFILM5.OUT"

'Const ModelFAVOR_Version = 1#
Const ModelFAVOR_ExeName = "f32voc.exe"
'Const ModelFAVOR_EofTestMarker = 123456#

'Global Const MODELTYPE_PSDM = 0
'Global Const MODELTYPE_CPHSDM = 1
'Global Const MODELTYPE_ECM = 2
Global Const MODELTYPE_BFILM = 3

'Const ModelFAVOR_NMAX = 20
'Private Type ModelFAVOR_Inputs_Type
'  NX As Integer                                   'DIMENSIONLESS
'  VOID_I As Double                                'DIMENSIONLESS
'  DEN_I As Double                                 'g/cm^3
'  FLRT_I As Double                                'gal/min-ft^2
'  INDEX_IO(1 To ModelFAVOR_NMAX) As Integer          'DIMENSIONLESS
'  XK_I(1 To ModelFAVOR_NMAX) As Double              '(umol/g)*(L/umol)^(1/n)
'  XN_I(1 To ModelFAVOR_NMAX) As Double              'DIMENSIONLESS
'  C0_I(1 To ModelFAVOR_NMAX) As Double              'ug/L
'  XMW_I(1 To ModelFAVOR_NMAX) As Double             'g/gmol
'End Type
'Dim ModelFAVOR_Inputs As ModelFAVOR_Inputs_Type
'Private Type ModelFAVOR_Outputs_Type
'  NX As Integer                                   'DIMENSIONLESS
'  C_O(1 To ModelFAVOR_NMAX, 1 To ModelFAVOR_NMAX) As Double
'  DGY_O(1 To ModelFAVOR_NMAX, 1 To ModelFAVOR_NMAX) As Double
'  FCS_O(1 To ModelFAVOR_NMAX, 1 To ModelFAVOR_NMAX) As Double
'  OATS_O(1 To ModelFAVOR_NMAX) As Double
'  Q_O(1 To ModelFAVOR_NMAX, 1 To ModelFAVOR_NMAX) As Double
'  QAVE_O(1 To ModelFAVOR_NMAX, 1 To ModelFAVOR_NMAX) As Double
'  SSTC_O(1 To ModelFAVOR_NMAX) As Double
'  VW_O(1 To ModelFAVOR_NMAX) As Double
'  ZZZ_O(1 To ModelFAVOR_NMAX) As Double
'  C0_O(1 To ModelFAVOR_NMAX) As Double              'ug/L
'End Type
'Dim ModelFAVOR_Outputs As ModelFAVOR_Outputs_Type

'MISC VARIABLES (TIMER).
Global ModelIO_Timer_TimeStart As String
Global ModelIO_Timer_TimeEnd As String
Global ModelIO_Timer_SummaryMsg As String

'
' MISC VARIABLES (ERROR DESCRIPTION).
'
Global ModelError_NFLAG As Integer
Global ModelError_Description() As String
Global ModelError_Description_LineCount As Integer




Const Model_FaVOr_declarations_end = True


Sub ModelFAVOR_Go()
  On Error GoTo err_ThisSub
  ''''Call ModelFAVOR_WritePathFile
  If (ModelFAVOR_WriteMainFile() = False) Then Exit Sub
  Call ModelFAVOR_CallEXE
  If (ModelFAVOR_ProcessOutput() = False) Then Exit Sub
  If (ModelIO_IsKeepTempFiles() = False) Then
  '  Call ModelFAVOR_RemoveLinkFiles
  End If
exit_normally_ThisSub:
  Calculated_OK = True
  Exit Sub
exit_err_ThisSub:
  Calculated_OK = False
  Exit Sub
err_ThisSub:
  Call Show_Trapped_Error("ModelFAVOR_Go")
  Calculated_OK = False
  GoTo exit_err_ThisSub
End Sub


Sub ModelFAVOR_RemoveLinkFiles()
  Call KillFile_If_Exists(MAIN_EXE_PATH & "\" & ModelFAVOR_IN_Main1)
  Call KillFile_If_Exists(MAIN_EXE_PATH & "\" & ModelFAVOR_IN_Main2)
  Call KillFile_If_Exists(MAIN_EXE_PATH & "\" & ModelFAVOR_OUT_Main1)
  Call KillFile_If_Exists(MAIN_EXE_PATH & "\" & ModelFAVOR_OUT_Main2)
End Sub
Sub ModelFAVOR_CallEXE()
Dim cmdline As String
Dim Test As String
  Call ChangeDir_Exes
  cmdline = ModelFAVOR_ExeName
  Call ModelIO_Timer_Start
  Call FortranLink_ExecAndWaitForProcess(cmdline)
  Call ModelIO_Timer_End
  Call ChangeDir_Main
End Sub
Function ModelFAVOR_ProcessOutput() As Boolean
Dim FileNum As Integer
Dim i%, LineCount%, emsg$
Dim OLine(120) As String
Dim fn_OUT_Main1 As String
Dim fn_OUT_Main2 As String
Dim j As Integer
  On Error GoTo err_ThisSub
  With NowProj
    fn_OUT_Main1 = MAIN_EXE_PATH & "\" & ModelFAVOR_OUT_Main1
    fn_OUT_Main2 = MAIN_EXE_PATH & "\" & ModelFAVOR_OUT_Main2
    If (Not FileExists(fn_OUT_Main1)) Then
      Call Show_Error("Could not open output file `" & _
          fn_OUT_Main1 & "`.")
      GoTo exit_err_ThisSub
    End If

    emsg$ = "Problem Reading Output File"
    FileNum = FreeFile
    Open fn_OUT_Main1 For Input As FileNum
    LineCount% = 0
    Do While (Not EOF(FileNum))
      Line Input #FileNum, OLine(LineCount%)
      If (Len(OLine(LineCount%)) > 10) Then
        LineCount% = LineCount% + 1
      End If
    Loop
    Close FileNum

    emsg$ = "Problem Converting Strings"

    i% = 0
    .OutputRec.TotalAmount.pr_Stripping = CDbl(Trim$(OLine(i%)))
    i% = i% + 1
    .OutputRec.TotalAmount.pr_Volatilization = CDbl(Trim$(OLine(i%)))
    i% = i% + 1
    .OutputRec.TotalAmount.pr_SolidWaste = CDbl(Trim$(OLine(i%)))
    i% = i% + 1
    .OutputRec.TotalAmount.pr_LiquidWaste = CDbl(Trim$(OLine(i%)))
    i% = i% + 1
    .OutputRec.TotalAmount.pr_Biodegredation = CDbl(Trim$(OLine(i%)))
    i% = i% + 1
    .OutputRec.pr_TotalRemoved = CDbl(Trim$(OLine(i%)))
    i% = i% + 1
    
    .OutputRec.TotalAmount.Stripping = CDbl(Trim$(OLine(i%)))
    i% = i% + 1
    .OutputRec.TotalAmount.Volatilization = CDbl(Trim$(OLine(i%)))
    i% = i% + 1
    .OutputRec.TotalAmount.SolidWaste = CDbl(Trim$(OLine(i%)))
    i% = i% + 1
    .OutputRec.TotalAmount.LiquidWaste = CDbl(Trim$(OLine(i%)))
    i% = i% + 1
    .OutputRec.TotalAmount.Biodegredation = CDbl(Trim$(OLine(i%)))
    i% = i% + 1
    .OutputRec.TotalInfluent = CDbl(Trim$(OLine(i%)))
    i% = i% + 1
    .OutputRec.TotalEffluent = CDbl(Trim$(OLine(i%)))
    i% = i% + 1
    
    If (.Plant.en_InfluentWeir) Then
      .OutputRec.InfluentWeir.EffluentConc = CDbl(Trim$(OLine(i%)))
      i% = i% + 1
      .OutputRec.InfluentWeir.Stripping = CDbl(Trim$(OLine(i%)))
      i% = i% + 1
      .OutputRec.InfluentWeir.pr_Stripping = CDbl(Trim$(OLine(i%)))
      i% = i% + 1
    End If
    
    If (.Plant.en_GritChamber) Then
      .OutputRec.GritChamber.EffluentConc = CDbl(Trim$(OLine(i%)))
      i% = i% + 1
      .OutputRec.GritChamber.Stripping = CDbl(Trim$(OLine(i%)))
      i% = i% + 1
      .OutputRec.GritChamber.pr_Stripping = CDbl(Trim$(OLine(i%)))
      i% = i% + 1
      .OutputRec.GritChamber.Volatilization = CDbl(Trim$(OLine(i%)))
      i% = i% + 1
      .OutputRec.GritChamber.pr_Volatilization = CDbl(Trim$(OLine(i%)))
      i% = i% + 1
    End If
    
    .OutputRec.PrimaryClarifier.EffluentConc = CDbl(Trim$(OLine(i%)))
    i% = i% + 1
    .OutputRec.PrimaryClarifier.Volatilization = CDbl(Trim$(OLine(i%)))
    i% = i% + 1
    .OutputRec.PrimaryClarifier.pr_Volatilization = CDbl(Trim$(OLine(i%)))
    i% = i% + 1
    .OutputRec.PrimaryClarifier.SolidWaste = CDbl(Trim$(OLine(i%)))
    i% = i% + 1
    .OutputRec.PrimaryClarifier.pr_SolidWaste = CDbl(Trim$(OLine(i%)))
    i% = i% + 1
    .OutputRec.PrimaryClarifier.LiquidWaste = CDbl(Trim$(OLine(i%)))
    i% = i% + 1
    .OutputRec.PrimaryClarifier.pr_LiquidWaste = CDbl(Trim$(OLine(i%)))
    i% = i% + 1
    
    If (.Plant.en_PrimaryWeir) Then
      .OutputRec.PrimaryWeir.EffluentConc = CDbl(Trim$(OLine(i%)))
      i% = i% + 1
      .OutputRec.PrimaryWeir.Stripping = CDbl(Trim$(OLine(i%)))
      i% = i% + 1
      .OutputRec.PrimaryWeir.pr_Stripping = CDbl(Trim$(OLine(i%)))
      i% = i% + 1
    End If
    
    .OutputRec.AerationBasin.EffluentConc = CDbl(Trim$(OLine(i%)))
    i% = i% + 1
    .OutputRec.AerationBasin.Stripping = CDbl(Trim$(OLine(i%)))
    i% = i% + 1
    .OutputRec.AerationBasin.pr_Stripping = CDbl(Trim$(OLine(i%)))
    i% = i% + 1
    .OutputRec.AerationBasin.Volatilization = CDbl(Trim$(OLine(i%)))
    i% = i% + 1
    .OutputRec.AerationBasin.pr_Volatilization = CDbl(Trim$(OLine(i%)))
    i% = i% + 1
    .OutputRec.AerationBasin.Biodegredation = CDbl(Trim$(OLine(i%)))
    i% = i% + 1
    .OutputRec.AerationBasin.pr_Biodegredation = CDbl(Trim$(OLine(i%)))
    i% = i% + 1
    
    .OutputRec.SecondaryClarifier.EffluentConc = CDbl(Trim$(OLine(i%)))
    i% = i% + 1
    .OutputRec.SecondaryClarifier.Volatilization = CDbl(Trim$(OLine(i%)))
    i% = i% + 1
    .OutputRec.SecondaryClarifier.pr_Volatilization = CDbl(Trim$(OLine(i%)))
    i% = i% + 1
    .OutputRec.SecondaryClarifier.SolidWaste = CDbl(Trim$(OLine(i%)))
    i% = i% + 1
    .OutputRec.SecondaryClarifier.pr_SolidWaste = CDbl(Trim$(OLine(i%)))
    i% = i% + 1
    .OutputRec.SecondaryClarifier.LiquidWaste = CDbl(Trim$(OLine(i%)))
    i% = i% + 1
    .OutputRec.SecondaryClarifier.pr_LiquidWaste = CDbl(Trim$(OLine(i%)))
    i% = i% + 1
    
    If (.Plant.en_SecondaryWeir) Then
      .OutputRec.SecondaryWeir.EffluentConc = CDbl(Trim$(OLine(i%)))
      i% = i% + 1
      .OutputRec.SecondaryWeir.Stripping = CDbl(Trim$(OLine(i%)))
      i% = i% + 1
      .OutputRec.SecondaryWeir.pr_Stripping = CDbl(Trim$(OLine(i%)))
      i% = i% + 1
    End If
    .KP1_OUT = CDbl(Trim$(OLine(i%)))
    i% = i% + 1
    For j = 1 To 7
      .XVALS_OUT(j) = CDbl(Trim$(OLine(i%)))
      i% = i% + 1
    Next j
  End With
  Call Show_Message( _
      "Model Calculations Complete." & _
      vbCrLf & _
      vbCrLf & _
      ModelIO_Timer_SummaryMsg)
exit_normally_ThisSub:
  ModelFAVOR_ProcessOutput = True
  Exit Function
exit_err_ThisSub:
  ModelFAVOR_ProcessOutput = False
  Exit Function
err_ThisSub:
  Call Show_Trapped_Error("ModelFAVOR_ProcessOutput")
  GoTo exit_err_ThisSub
End Function
Function ModelFAVOR_WriteMainFile() As Boolean
Dim fn_IN_Main1 As String
Dim fn_IN_Main2 As String
Dim f As Integer
Dim i As Integer
  On Error GoTo err_ThisSub
  With NowProj
    fn_IN_Main1 = MAIN_EXE_PATH & "\" & ModelFAVOR_IN_Main1
    fn_IN_Main2 = MAIN_EXE_PATH & "\" & ModelFAVOR_IN_Main2
    f = FreeFile
    Open fn_IN_Main1 For Output As #f
    '
    ' ====================  Plant INFLUENT
    '
    Print #f,
    Print #f, "====================  Plant INFLUENT"
    Print #f,
    
    Print #f, ".Plant Flow Rate (Q, L/day)"
    Print #f, FormatSimulationNumber(.Plant.Flow)
    
    Print #f, "Solids Influent Concentration (X0, mg/L)"
    Print #f, FormatSimulationNumber(.Plant.SolidsConc)
    
    
    
  ' ====================  PHYSICO-CHEMICAL PROPERTIES
    Print #f,
    Print #f, "====================  PHYSICO-CHEMICAL PROPERTIES"
    Print #f,
    
    Print #f, "Barometric Pressure (PB, kPa)"
    Print #f, FormatSimulationNumber(.Plant.ChemicalData.env_Pressure)
    
    Print #f, "Temperature (T, C)"
    Print #f, FormatSimulationNumber(.Plant.ChemicalData.env_Temperature)
    
    Print #f, "Wind Velocity 10 meters above .Plant (WNDVRI, m/s)"
    Print #f, FormatSimulationNumber(.Plant.ChemicalData.env_WindVelocity)
    '
    ' THE FORTRAN IS READING IN AN INTEGER FOR SOME REASON?
    '
    Print #f, "Name of Contaminant (NAME) tetra"
    Print #f, FormatSimulationNumber(1)
    
    Print #f, "Contaminant Influent Concentration (CO1, ug/L)"
    Print #f, FormatSimulationNumber(.Plant.ChemicalData.InfluentConc)
    
    Print #f, "Biodegradation Rate Constant (KB, L/(mg*day))"
    Print #f, FormatSimulationNumber(.Plant.ChemicalData.BiodegredationRate)
    
    Print #f, "Log Octanol Water Coefficient For Contaminant (LOGKOW)"
    Print #f, FormatSimulationNumber(.Plant.ChemicalData.LogKow)
    
    Print #f, "Henry's Constant (H)"
    Print #f, FormatSimulationNumber(.Plant.ChemicalData.VOC_HenrysConstant)
    
    Print #f, "Molecular Weight (MW, g/mol)"
    Print #f, FormatSimulationNumber(.Plant.ChemicalData.VOC_MolecularWeight)
    
    Print #f, "Diffusivity of Contaminant IN H2O (VOCDIF, cm2/sec)"
    Print #f, FormatSimulationNumber(.Plant.ChemicalData.VOC_DiffusivityInH2O)
    
    Print #f, "Gas Phase Contaminant Diffusivity (VOCDFG, cm2/sec)"
    Print #f, FormatSimulationNumber(.Plant.ChemicalData.VOC_DiffusivityInGas)
    
    Print #f, "Oxygen Saturation Concentration at Effective Depth (CSAT, mg/L)"
    Print #f, FormatSimulationNumber(.Plant.ChemicalData.O2_SaturationConc)
    
    Print #f, "Henry's Constant for Oxygen (HOXY)"
    Print #f, FormatSimulationNumber(.Plant.ChemicalData.O2_HenrysConstant)
    
    Print #f, "Diffusivity of Oxygen (OXYDIF, cm2/sec)"
    Print #f, FormatSimulationNumber(.Plant.ChemicalData.O2_Diffusivity)
    
    Print #f, "Density of Water (H2ODEN, kg/m3)"
    Print #f, FormatSimulationNumber(.Plant.ChemicalData.H2O_Density)
    
    Print #f, "Viscosity of Water (H2OVIS, kg/(m*s))"
    Print #f, FormatSimulationNumber(.Plant.ChemicalData.H2O_Viscosity)
    
    Print #f, "Vapor Pressure (PV, kPa)"
    Print #f, FormatSimulationNumber(.Plant.ChemicalData.H2O_VaporPressure)
    
    Print #f, "Process Water Correction (ALPHA)"
    Print #f, FormatSimulationNumber(.Plant.ChemicalData.H2O_Alpha)
    
    Print #f, "Density of Air (AIRDEN, kg/m3)"
    Print #f, FormatSimulationNumber(.Plant.ChemicalData.AIR_Density)
    
    Print #f, "Viscosity of Air (AIRVIS, kg/(m*s))"
    Print #f, FormatSimulationNumber(.Plant.ChemicalData.AIR_Viscosity)
    
     
    
    
    ' ====================  PRIMARY CLARIFIERS
    Print #f,
    Print #f, "====================  PRIMARY CLARIFIERS"
    Print #f,
    
    Print #f, "Number of Primary Clarifiers being Modeled (NPC)"
    Print #f, FormatSimulationNumber(.Plant.PrimaryClarifier.Count)
    
    Print #f, "Covered Primary Carifier Option (CBPC, 0=off -1=on)"
    Print #f, FormatSimulationNumber(.Plant.PrimaryClarifier.IsCovered)
    
    Print #f, "Primary Clarifier Ventilation Air Flow Rate for each Clarifier (QVPC, L/min)"
    Print #f, FormatSimulationNumber(.Plant.PrimaryClarifier.VentilationRate)
    
    Print #f, "Primary Clarifier Basin Depth for each (PCD, m)"
    Print #f, FormatSimulationNumber(.Plant.PrimaryClarifier.Depth)
    
    Print #f, "Primary Clarifier Volume for each (PV1, L)"
    Print #f, FormatSimulationNumber(.Plant.PrimaryClarifier.Volume)
        
    Print #f, "Primary Wastage Flow Rate from each Clarifier  (QW1, L/day)"
    Print #f, FormatSimulationNumber(.Plant.PrimaryClarifier.WastageFlow)
    
    ''''Print #f, "Percent Solids Removal in Primary Clarifier (E)  (This is new as of 1999-May-06)"
    ''''Print #f, FormatSimulationNumber(.Plant.PrimaryClarifier.PercentageRemoval)
    Print #f, "Fractional Solids Removal in Primary Clarifier (E)  (Ranges 0-1)  (This is new as of 1999-May-06)"
    Print #f, FormatSimulationNumber(.Plant.PrimaryClarifier.PercentageRemoval / 100#)
    
    Print #f, "Primary Sorption Mechanism (PSM, 1=Dobbs 2=Matter Muller)"
    Print #f, FormatSimulationNumber(.Plant.PrimaryClarifier.SorptionRemovalMethod + 1)
    
    Print #f, "Primary Volitalization Mechanism (PVM)"
    Print #f, FormatSimulationNumber(.Plant.PrimaryClarifier.VolatilizationRemovalMechanism + 1)
    
    
    
    
    ' ====================  INFLUENT WEIR
    Print #f,
    Print #f, "====================  INFLUENT WEIR"
    Print #f,
    
    Print #f, "Option for Influent Weir Drop (W1, 0=off -1=on)"
    Print #f, FormatSimulationNumber(.Plant.en_InfluentWeir)
    
    Print #f, "Influent Weir Mechanism (WM1, 1=NAPPE 2=POOL)"
    Print #f, FormatSimulationNumber(.Plant.InfluentWeir.ModelingMechanism + 1)
    
    Print #f, "Width of Weir Channel (WW1, m)"
    Print #f, FormatSimulationNumber(.Plant.InfluentWeir.Width)
    
    Print #f, "Distance Between Water Levels Above and Below Weir (Z1, m)"
    Print #f, FormatSimulationNumber(.Plant.InfluentWeir.WaterLevelDiff)
    
    Print #f, "Gas Flow Rate Leaving the Tailwater per Unit Weir Length (QG1, m3/(m*h))"
    Print #f, FormatSimulationNumber(.Plant.InfluentWeir.GasFlow)
    
     
     
  ' ====================  PRIMARY CLARIFIER WEIR
    Print #f,
    Print #f, "====================  PRIMARY CLARIFIER WEIR"
    Print #f,
    
    Print #f, "0ption for Primary Effluent Weir Drop (W2, 0=off -1=on)"
    Print #f, FormatSimulationNumber(.Plant.en_PrimaryWeir)
    
    Print #f, "Primary Effluent Weir Mechanism (WM2, 1=NAPPE 2=POOL)"
    Print #f, FormatSimulationNumber(.Plant.PrimaryWeir.ModelingMechanism + 1)
    
    Print #f, "Length of the Weir Channel (WW2, m)"
    Print #f, FormatSimulationNumber(.Plant.PrimaryWeir.Width)
    
    Print #f, "Distance Between the Water Levels Above and Below Weir (Z2, m)"
    Print #f, FormatSimulationNumber(.Plant.PrimaryWeir.WaterLevelDiff)
    
    Print #f, "Gas Flow Rate Leaving the Tailwater per Unit Weir Length (QG2, m3/(m*h))"
    Print #f, FormatSimulationNumber(.Plant.PrimaryWeir.GasFlow)
      
      
      
  ' ====================  SECONDARY CLARIFIER WEIR
    Print #f,
    Print #f, "====================  SECONDARY CLARIFIER WEIR"
    Print #f,
  
    Print #f, "Option for Effluent Drop (W3, 0=off -1=on)"
    Print #f, FormatSimulationNumber(.Plant.en_SecondaryWeir)
    
    Print #f, "Effluent Weir Mechanism (WM3, 1=NAPPE 2=POOL)"
    Print #f, FormatSimulationNumber(.Plant.SecondaryWeir.ModelingMechanism + 1)
     
    Print #f, "Length of the Weir Channel (WW3, m)"
    Print #f, FormatSimulationNumber(.Plant.SecondaryWeir.Width)
     
    Print #f, "Distance Between Water Levels Above and Below Weir (Z3, m)"
    Print #f, FormatSimulationNumber(.Plant.SecondaryWeir.WaterLevelDiff)
    
    Print #f, "Gas Flow Rate Leaving Tailwater per Unit Weir Length (QG3, m3/(m*h))"
    Print #f, FormatSimulationNumber(.Plant.SecondaryWeir.GasFlow)
  
    
    
  ' ====================  AERATED GRIT CHAMBER
    Print #f,
    Print #f, "====================  AERATED GRIT CHAMBER"
    Print #f,
   
    Print #f, "Option for Modeling of Aerated Grit Chamber (GC, 0=off -1=on)"
    Print #f, FormatSimulationNumber(.Plant.en_GritChamber)
    
    Print #f, "Number of Aerated Grit Chambers being modeled (NAGC)"
    Print #f, FormatSimulationNumber(.Plant.GritChamber.Count)
    
    Print #f, "Covered Aerated Grit Chamber Option (CBAGC, 0=off -1=on)"
    Print #f, FormatSimulationNumber(.Plant.GritChamber.IsCovered)
    
    Print #f, "Aerated Grit Total Ventilation Air Flow Rate for each Chamber (QVAGC, )"
    Print #f, FormatSimulationNumber(.Plant.GritChamber.VentilationRate)
    
    Print #f, "Aerated Grit Chamber Depth for each (AGCD, m)"
    Print #f, FormatSimulationNumber(.Plant.GritChamber.Depth)
    
    Print #f, "Aerated Grit Chamber Volume for each (AGCV, L)"
    Print #f, FormatSimulationNumber(.Plant.GritChamber.Volume)
    
    Print #f, "Gas Flow Rate for each Aerated Grit Chamber (QGGC, L/min)"
    Print #f, FormatSimulationNumber(.Plant.GritChamber.GasFlow)
    
    Print #f, "SOTR For Bubble Aeration in Aerated Grit Chamber (AGBSOT, kg/hr)"
    Print #f, FormatSimulationNumber(.Plant.GritChamber.SOTR)
  
    
    
    
  ' ====================  AERATION BASIN
    Print #f,
    Print #f, "====================  AERATION BASIN"
    Print #f,
   
    Print #f, "Number of Aeration Tanks being Modeled (NAB)"
    Print #f, FormatSimulationNumber(.Plant.AerationBasin.Count)
    
    Print #f, "Covered Aeration Basin Option (CBAB, 0=off -1=on)"
    Print #f, FormatSimulationNumber(.Plant.AerationBasin.IsCovered)
    
    Print #f, "Ventilation Air Flow Rate for each Aeration Basin (QV, )"
    Print #f, FormatSimulationNumber(.Plant.AerationBasin.VentilationRate)
    
    Print #f, "Aeration Basin Depth for each (ABD, m)"
    Print #f, FormatSimulationNumber(.Plant.AerationBasin.Depth)
    
    Print #f, "Sludge Wastage From Each Secondary Clarifer, SQW (L/day)"
    Print #f, FormatSimulationNumber(.Plant.AerationBasin.WastageFlow)
    
    Print #f, "Recycle Flow Rate from each Secondary Clarifier (QR, L/day)"
    Print #f, FormatSimulationNumber(.Plant.AerationBasin.RecycleFlow)
    
    Print #f, "Secondary Aeration Mechanism (SAM, 1=SURFACE 3=DIFFUSED BUBBLE)"
    Print #f, IIf(.Plant.AerationBasin.ModelingMechanism = 0, 1, 3)
    
    Print #f, "SOTR in Aeration Basin (ABBSOT or SUFSOR, kg/hr)"
    Print #f, FormatSimulationNumber(.Plant.AerationBasin.SOTR)
    
    
    
    
  ' ====================  SECONDARY CLARIFIERS
    Print #f,
    Print #f, "====================  SECONDARY CLARIFIERS"
    Print #f,
    
    Print #f, "Number of Secondary Clarifiers being Modeled (NSC)"
    Print #f, FormatSimulationNumber(.Plant.SecondaryClarifier.Count)
    
    Print #f, "Covered Secondary Clarifier Option (CBSC, 0=off -1=on)"
    Print #f, FormatSimulationNumber(.Plant.SecondaryClarifier.IsCovered)
    
    Print #f, "Secondary Clarifier Ventilation Air Flow Rate for each Clarifier (QVSC, L/min)"
    Print #f, FormatSimulationNumber(.Plant.SecondaryClarifier.VentilationRate)
    
    Print #f, "Secondary Clarifier Basin Depth for each (SCBD, m)"
    Print #f, FormatSimulationNumber(.Plant.SecondaryClarifier.Depth)
    
    Print #f, "Secondary Clarifier Basin Volume for each (SCBV, L)"
    Print #f, FormatSimulationNumber(.Plant.SecondaryClarifier.Volume)
    
    Print #f, "Solids Concentration in Secondary Effluent (XSC, mg/L)  (This is new as of 1999-May-06)"
    Print #f, FormatSimulationNumber(.Plant.SecondaryClarifier.EffluentSolidsConc)
    
    Close f
  
  
  
    ' ====================  AERATION BASIN CSTR'S
  
    ' TRANSFER THE TOTALS TO THE CSTR VARIABLES
    If (.Plant.AerationBasin.CSTR.Count = 1) Then
      .Plant.AerationBasin.CSTR.BioMass(0) = .Plant.AerationBasin.BioMass
      .Plant.AerationBasin.CSTR.Volume(0) = .Plant.AerationBasin.Volume
      .Plant.AerationBasin.CSTR.GasFlow(0) = .Plant.AerationBasin.GasFlow
      .Plant.AerationBasin.CSTR.Feed(0) = 1#
    End If
  
    f = FreeFile
    Open fn_IN_Main2 For Output As f
    Print #f, "Step Feed Modeling (SF, 0=off -1=on)"
    Print #f, FormatSimulationNumber(.Plant.AerationBasin.CSTR.UseStepFeed)
    
    Print #f, "Number of Tanks Being Modeled (NTK)"
    Print #f, FormatSimulationNumber(.Plant.AerationBasin.CSTR.Count)
    
    For i = 0 To .Plant.AerationBasin.CSTR.Count - 1
      Print #f, "Biomass Concentration For A Particular Tank (XBM(IC), mg/L)"
      Print #f, FormatSimulationNumber(.Plant.AerationBasin.CSTR.BioMass(i))
      
      Print #f, "Aeration Tank Volume For A Particular Tank (ATV(IC), L)"
      Print #f, FormatSimulationNumber(.Plant.AerationBasin.CSTR.Volume(i))
      
      Print #f, "Gas Flow Rate For A Particular Tank (QG(IC), L/min)"
      Print #f, FormatSimulationNumber(.Plant.AerationBasin.CSTR.GasFlow(i))
      
      Print #f, "Fraction of .Plant Influent Directly Entering CSTR 1 (FFRACT(1))"
      Print #f, FormatSimulationNumber(.Plant.AerationBasin.CSTR.Feed(i))
    Next i
    Close f
  End With
exit_normally_ThisSub:
  ModelFAVOR_WriteMainFile = True
  Exit Function
exit_err_ThisSub:
  ModelFAVOR_WriteMainFile = False
  Exit Function
err_ThisSub:
  Call Show_Trapped_Error("ModelFAVOR_WriteMainFile")
  GoTo exit_err_ThisSub
End Function
Sub ModelFAVOR_WritePathFile()
'Dim f As Integer
'Dim fn_This As String
'Dim qq As String
'  qq = Chr$(34)
'  f = FreeFile
'  fn_This = MAIN_EXE_PATH & "\" & ModelFAVOR_IN_PathFile
'  'fn_This = App.Path & "\" & ModelFAVOR_IN_PathFile
'  Open fn_This For Output As #f
'  Print #f, "1"
'  Print #f, qq & ModelFAVOR_IN_Main & qq
'  Print #f, qq & ModelFAVOR_OUT_SuccessFlag & qq
'  Print #f, qq & ModelFAVOR_OUT_Main & qq
'  Print #f, qq & ModelFAVOR_OUT_CvsT & qq
'  Print #f, qq & ModelFAVOR_OUT_EndCvsT & qq
'  Print #f, qq & ModelFAVOR_OUT_CvsT_Details & qq
'  Close #f
End Sub











'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////     THE FOLLOWING CODE APPLIES TO ALL MODELS, NOT JUST THE ECM.     /////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function ModelIO_DoNumberCheck(N1 As Double, N2 As Double) As Boolean
  If (Abs(N1 + 0.000001) / N2 - 1#) <= 0.001 Then
    'NUMBERS ARE CLOSE TO IDENTICAL.
    ModelIO_DoNumberCheck = True
  Else
    'NUMBERS ARE _NOT_ CLOSE TO IDENTICAL.
    ModelIO_DoNumberCheck = False
  End If
End Function


Sub ModelIO_Timer_Start()
  ModelIO_Timer_TimeStart = Now
End Sub
Sub ModelIO_Timer_End()
Dim Elapsed_Min As Double
Dim TotalTimeStr As String
  ModelIO_Timer_TimeEnd = Now
  Elapsed_Min = _
      DateDiff("s", ModelIO_Timer_TimeStart, _
               ModelIO_Timer_TimeEnd) / 60#
  TotalTimeStr = Format$(Elapsed_Min, "0.00")
  ModelIO_Timer_SummaryMsg = _
      "Calculation Started:    " & ModelIO_Timer_TimeStart & _
      vbCrLf & _
      "Calculation Ended:    " & ModelIO_Timer_TimeEnd & _
      vbCrLf & _
      vbCrLf & _
      "Total Calculation Time = " & TotalTimeStr & " Minutes"
End Sub


Function ModelIO_IsKeepTempFiles() As Boolean
  ModelIO_IsKeepTempFiles = frmMain.mnuMTUItem(40).Checked
  'ModelIO_IsKeepTempFiles = True
  'ModelIO_IsKeepTempFiles = False
End Function


Function FormatSimulationNumber(Value As Variant) As String
Dim pformat$
  Select Case VarType(Value)
    Case vbLong, vbInteger
      pformat$ = "0"
    Case vbDouble, vbSingle
      pformat$ = "0.00000E+00"
    Case vbBoolean
      If (Value) Then
        FormatSimulationNumber = "-1"
      Else
        FormatSimulationNumber = "0"
      End If
      Exit Function
    Case vbString
      FormatSimulationNumber = Value
      Exit Function
  End Select
  FormatSimulationNumber = Format$(Value, pformat$)
End Function

