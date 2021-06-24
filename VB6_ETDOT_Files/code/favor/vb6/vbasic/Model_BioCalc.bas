Attribute VB_Name = "Model_BioCalc"
Option Explicit

'Global Const ModelBIOCALC_IN_PathFile = "BFILM1.IN"
Global Const ModelBIOCALC_IN_Main1 = "indata.dat"
Global Const ModelBIOCALC_IN_Main2 = "{ not available }"
'Global Const ModelBIOCALC_OUT_SuccessFlag = "BFILM1.OUT"
Global Const ModelBIOCALC_OUT_Main1 = "out.dat"
Global Const ModelBIOCALC_OUT_Main2 = "{ not available }"
'Global Const ModelBIOCALC_OUT_CvsT = "BFILM3.OUT"
'Global Const ModelBIOCALC_OUT_EndCvsT = "BFILM4.OUT"
'Global Const ModelBIOCALC_OUT_CvsT_Details = "BFILM5.OUT"

'Const ModelBIOCALC_Version = 1#
Const ModelBIOCALC_ExeName = "biocalc.exe"
'Const ModelBIOCALC_EofTestMarker = 123456#

'Global Const MODELTYPE_PSDM = 0
'Global Const MODELTYPE_CPHSDM = 1
'Global Const MODELTYPE_ECM = 2
'Global Const MODELTYPE_BFILM = 3
Global Const MODELTYPE_BIOCALC = 4

'Const ModelBIOCALC_NMAX = 20
'Private Type ModelBIOCALC_Inputs_Type
'  NX As Integer                                   'DIMENSIONLESS
'  VOID_I As Double                                'DIMENSIONLESS
'  DEN_I As Double                                 'g/cm^3
'  FLRT_I As Double                                'gal/min-ft^2
'  INDEX_IO(1 To ModelBIOCALC_NMAX) As Integer          'DIMENSIONLESS
'  XK_I(1 To ModelBIOCALC_NMAX) As Double              '(umol/g)*(L/umol)^(1/n)
'  XN_I(1 To ModelBIOCALC_NMAX) As Double              'DIMENSIONLESS
'  C0_I(1 To ModelBIOCALC_NMAX) As Double              'ug/L
'  XMW_I(1 To ModelBIOCALC_NMAX) As Double             'g/gmol
'End Type
'Dim ModelBIOCALC_Inputs As ModelBIOCALC_Inputs_Type
'Private Type ModelBIOCALC_Outputs_Type
'  NX As Integer                                   'DIMENSIONLESS
'  C_O(1 To ModelBIOCALC_NMAX, 1 To ModelBIOCALC_NMAX) As Double
'  DGY_O(1 To ModelBIOCALC_NMAX, 1 To ModelBIOCALC_NMAX) As Double
'  FCS_O(1 To ModelBIOCALC_NMAX, 1 To ModelBIOCALC_NMAX) As Double
'  OATS_O(1 To ModelBIOCALC_NMAX) As Double
'  Q_O(1 To ModelBIOCALC_NMAX, 1 To ModelBIOCALC_NMAX) As Double
'  QAVE_O(1 To ModelBIOCALC_NMAX, 1 To ModelBIOCALC_NMAX) As Double
'  SSTC_O(1 To ModelBIOCALC_NMAX) As Double
'  VW_O(1 To ModelBIOCALC_NMAX) As Double
'  ZZZ_O(1 To ModelBIOCALC_NMAX) As Double
'  C0_O(1 To ModelBIOCALC_NMAX) As Double              'ug/L
'End Type
'Dim ModelBIOCALC_Outputs As ModelBIOCALC_Outputs_Type

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




Const Model_BioCalc_declarations_end = True


Function ModelBIOCALC_Go(This_Plant As TYPE_PlantDiagram) As Boolean
  On Error GoTo err_ThisSub
  ''''Call ModelBIOCALC_WritePathFile
  If (ModelBIOCALC_WriteMainFile() = False) Then Exit Function
  Call ModelBIOCALC_CallEXE
  If (ModelBIOCALC_ProcessOutput(This_Plant) = False) Then Exit Function
  If (ModelIO_IsKeepTempFiles() = False) Then
    Call ModelBIOCALC_RemoveLinkFiles
  End If
exit_normally_ThisSub:
  ModelBIOCALC_Go = True
  Exit Function
exit_err_ThisSub:
  ModelBIOCALC_Go = False
  Exit Function
err_ThisSub:
  Call Show_Trapped_Error("ModelBIOCALC_Go")
  GoTo exit_err_ThisSub
End Function


Sub ModelBIOCALC_RemoveLinkFiles()
  Call KillFile_If_Exists(MAIN_EXE_PATH & "\" & ModelBIOCALC_IN_Main1)
  Call KillFile_If_Exists(MAIN_EXE_PATH & "\" & ModelBIOCALC_IN_Main2)
  Call KillFile_If_Exists(MAIN_EXE_PATH & "\" & ModelBIOCALC_OUT_Main1)
  Call KillFile_If_Exists(MAIN_EXE_PATH & "\" & ModelBIOCALC_OUT_Main2)
End Sub
Sub ModelBIOCALC_CallEXE()
Dim cmdline As String
Dim Test As String
  Call ChangeDir_Exes
  cmdline = ModelBIOCALC_ExeName
  Call ModelIO_Timer_Start
  Call FortranLink_ExecAndWaitForProcess(cmdline)
  Call ModelIO_Timer_End
  Call ChangeDir_Main
End Sub
Function ModelBIOCALC_ProcessOutput(This_Plant As TYPE_PlantDiagram) As Boolean
Dim FileNum As Integer
Dim i As Integer
Dim LineCount%
Dim emsg$
Dim OLine(120) As String
Dim fn_OUT_Main1 As String
Dim N As Integer
Dim sThisLine As String
Dim dblConcBiomass(1 To 50) As Double
  On Error GoTo err_ThisSub
  With This_Plant
    fn_OUT_Main1 = MAIN_EXE_PATH & "\" & ModelBIOCALC_OUT_Main1
    If (Not FileExists(fn_OUT_Main1)) Then
      Call Show_Error("Could not open output file `" & _
          fn_OUT_Main1 & "`.")
      GoTo exit_err_ThisSub
    End If

    emsg$ = "Problem Reading Output File"
    FileNum = FreeFile
    Open fn_OUT_Main1 For Input As #FileNum
    LineCount% = 0
    N = .AerationBasin.CSTR.Count
    For i = 1 To N
      Line Input #FileNum, sThisLine
      ' DO NOTHING WITH THIS LINE.
      Line Input #FileNum, sThisLine
      dblConcBiomass(i) = CDbl(Val(sThisLine))
      Line Input #FileNum, sThisLine
      ' DO NOTHING WITH THIS LINE.
    Next i
    Close #FileNum
    '
    ' IMPORT THE NEW VALUES.
    '
    .AerationBasin.CSTR.UniformBioMass = False
    For i = 0 To N - 1
      .AerationBasin.CSTR.BioMass(i) = dblConcBiomass(i + 1)
    Next i%
  End With
    
    
    
''''    Do While (Not EOF(FileNum))
''''      Line Input #FileNum, OLine(LineCount%)
''''      If (Len(OLine(LineCount%)) > 10) Then
''''        LineCount% = LineCount% + 1
''''      End If
''''    Loop
''''    Close FileNum
''''
''''    emsg$ = "Problem Converting Strings"
''''
''''    i = 0
''''    .OutputRec.TotalAmount.pr_Stripping = CDbl(Trim$(OLine(i)))
''''    i = i + 1
''''    .OutputRec.TotalAmount.pr_Volatilization = CDbl(Trim$(OLine(i)))
''''    i = i + 1
''''    .OutputRec.TotalAmount.pr_SolidWaste = CDbl(Trim$(OLine(i)))
''''    i = i + 1
''''    .OutputRec.TotalAmount.pr_LiquidWaste = CDbl(Trim$(OLine(i)))
''''    i = i + 1
''''    .OutputRec.TotalAmount.pr_Biodegredation = CDbl(Trim$(OLine(i)))
''''    i = i + 1
''''    .OutputRec.pr_TotalRemoved = CDbl(Trim$(OLine(i)))
''''    i = i + 1
''''
''''    .OutputRec.TotalAmount.Stripping = CDbl(Trim$(OLine(i)))
''''    i = i + 1
''''    .OutputRec.TotalAmount.Volatilization = CDbl(Trim$(OLine(i)))
''''    i = i + 1
''''    .OutputRec.TotalAmount.SolidWaste = CDbl(Trim$(OLine(i)))
''''    i = i + 1
''''    .OutputRec.TotalAmount.LiquidWaste = CDbl(Trim$(OLine(i)))
''''    i = i + 1
''''    .OutputRec.TotalAmount.Biodegredation = CDbl(Trim$(OLine(i)))
''''    i = i + 1
''''    .OutputRec.TotalInfluent = CDbl(Trim$(OLine(i)))
''''    i = i + 1
''''    .OutputRec.TotalEffluent = CDbl(Trim$(OLine(i)))
''''    i = i + 1
''''
''''    If (.Plant.en_InfluentWeir) Then
''''      .OutputRec.InfluentWeir.EffluentConc = CDbl(Trim$(OLine(i)))
''''      i = i + 1
''''      .OutputRec.InfluentWeir.Stripping = CDbl(Trim$(OLine(i)))
''''      i = i + 1
''''      .OutputRec.InfluentWeir.pr_Stripping = CDbl(Trim$(OLine(i)))
''''      i = i + 1
''''    End If
''''
''''    If (.Plant.en_GritChamber) Then
''''      .OutputRec.GritChamber.EffluentConc = CDbl(Trim$(OLine(i)))
''''      i = i + 1
''''      .OutputRec.GritChamber.Stripping = CDbl(Trim$(OLine(i)))
''''      i = i + 1
''''      .OutputRec.GritChamber.pr_Stripping = CDbl(Trim$(OLine(i)))
''''      i = i + 1
''''      .OutputRec.GritChamber.Volatilization = CDbl(Trim$(OLine(i)))
''''      i = i + 1
''''      .OutputRec.GritChamber.pr_Volatilization = CDbl(Trim$(OLine(i)))
''''      i = i + 1
''''    End If
''''
''''    .OutputRec.PrimaryClarifier.EffluentConc = CDbl(Trim$(OLine(i)))
''''    i = i + 1
''''    .OutputRec.PrimaryClarifier.Volatilization = CDbl(Trim$(OLine(i)))
''''    i = i + 1
''''    .OutputRec.PrimaryClarifier.pr_Volatilization = CDbl(Trim$(OLine(i)))
''''    i = i + 1
''''    .OutputRec.PrimaryClarifier.SolidWaste = CDbl(Trim$(OLine(i)))
''''    i = i + 1
''''    .OutputRec.PrimaryClarifier.pr_SolidWaste = CDbl(Trim$(OLine(i)))
''''    i = i + 1
''''    .OutputRec.PrimaryClarifier.LiquidWaste = CDbl(Trim$(OLine(i)))
''''    i = i + 1
''''    .OutputRec.PrimaryClarifier.pr_LiquidWaste = CDbl(Trim$(OLine(i)))
''''    i = i + 1
''''
''''    If (.Plant.en_PrimaryWeir) Then
''''      .OutputRec.PrimaryWeir.EffluentConc = CDbl(Trim$(OLine(i)))
''''      i = i + 1
''''      .OutputRec.PrimaryWeir.Stripping = CDbl(Trim$(OLine(i)))
''''      i = i + 1
''''      .OutputRec.PrimaryWeir.pr_Stripping = CDbl(Trim$(OLine(i)))
''''      i = i + 1
''''    End If
''''
''''    .OutputRec.AerationBasin.EffluentConc = CDbl(Trim$(OLine(i)))
''''    i = i + 1
''''    .OutputRec.AerationBasin.Stripping = CDbl(Trim$(OLine(i)))
''''    i = i + 1
''''    .OutputRec.AerationBasin.pr_Stripping = CDbl(Trim$(OLine(i)))
''''    i = i + 1
''''    .OutputRec.AerationBasin.Volatilization = CDbl(Trim$(OLine(i)))
''''    i = i + 1
''''    .OutputRec.AerationBasin.pr_Volatilization = CDbl(Trim$(OLine(i)))
''''    i = i + 1
''''    .OutputRec.AerationBasin.Biodegredation = CDbl(Trim$(OLine(i)))
''''    i = i + 1
''''    .OutputRec.AerationBasin.pr_Biodegredation = CDbl(Trim$(OLine(i)))
''''    i = i + 1
''''
''''    .OutputRec.SecondaryClarifier.EffluentConc = CDbl(Trim$(OLine(i)))
''''    i = i + 1
''''    .OutputRec.SecondaryClarifier.Volatilization = CDbl(Trim$(OLine(i)))
''''    i = i + 1
''''    .OutputRec.SecondaryClarifier.pr_Volatilization = CDbl(Trim$(OLine(i)))
''''    i = i + 1
''''    .OutputRec.SecondaryClarifier.SolidWaste = CDbl(Trim$(OLine(i)))
''''    i = i + 1
''''    .OutputRec.SecondaryClarifier.pr_SolidWaste = CDbl(Trim$(OLine(i)))
''''    i = i + 1
''''    .OutputRec.SecondaryClarifier.LiquidWaste = CDbl(Trim$(OLine(i)))
''''    i = i + 1
''''    .OutputRec.SecondaryClarifier.pr_LiquidWaste = CDbl(Trim$(OLine(i)))
''''    i = i + 1
''''
''''    If (.Plant.en_SecondaryWeir) Then
''''      .OutputRec.SecondaryWeir.EffluentConc = CDbl(Trim$(OLine(i)))
''''      i = i + 1
''''      .OutputRec.SecondaryWeir.Stripping = CDbl(Trim$(OLine(i)))
''''      i = i + 1
''''      .OutputRec.SecondaryWeir.pr_Stripping = CDbl(Trim$(OLine(i)))
''''      i = i + 1
''''    End If
''''  End With
  Call Show_Message( _
      "Model Calculations Complete." & _
      vbCrLf & _
      vbCrLf & _
      ModelIO_Timer_SummaryMsg)
exit_normally_ThisSub:
  ModelBIOCALC_ProcessOutput = True
  Exit Function
exit_err_ThisSub:
  ModelBIOCALC_ProcessOutput = False
  Exit Function
err_ThisSub:
  Call Show_Trapped_Error("ModelBIOCALC_ProcessOutput")
  GoTo exit_err_ThisSub
End Function
Function ModelBIOCALC_WriteMainFile() As Boolean
Dim fn_IN_Main1 As String
Dim f As Integer
Dim i As Integer
Dim N As Integer
Dim iThis As Integer
  On Error GoTo err_ThisSub
  With NowProj
    fn_IN_Main1 = MAIN_EXE_PATH & "\" & ModelBIOCALC_IN_Main1
    f = FreeFile
    Open fn_IN_Main1 For Output As #f
    N = .Plant.AerationBasin.CSTR.Count
    Call WriteFortranInput(f, N, "N, number of CSTRs, unitless")
    iThis = IIf(.Plant.AerationBasin.CSTR.UseStepFeed = True, 1, 0)
    Call WriteFortranInput(f, iThis, "NSF, step feed option (0=no step feed, 1=step feed)")
    For i = 1 To N
      If (N = 1) Then
        Call WriteFortranInput(f, 1#, "FFRACT(i=" & Trim$(Str$(i)) & "), i=1...n, feed fraction #i, unitless")
      Else
        Call WriteFortranInput(f, .Plant.AerationBasin.CSTR.Feed(i - 1), "FFRACT(i=" & Trim$(Str$(i)) & "), i=1...n, feed fraction #i, unitless")
      End If
    Next i
    Call WriteFortranInput(f, .Plant.PrimaryClarifier.Count, "NPC, number of primary clarifiers, unitless")
    Call WriteFortranInput(f, .Plant.AerationBasin.Count, "NAB, number of aeration basins, unitless")
    Call WriteFortranInput(f, .Plant.SecondaryClarifier.Count, "NSC, number of secondary clarifiers, unitless")
    Call WriteFortranInput(f, .Plant.Flow, "Q0, influent flow rate, L/day")
    Call WriteFortranInput(f, .Plant.PrimaryClarifier.WastageFlow, "QW1, wastage flow rate from each primary clarifier, L/day")
    Call WriteFortranInput(f, .Plant.AerationBasin.WastageFlow, "QW, wastage flow rate from each secondary clarifier, L/day")
    Call WriteFortranInput(f, .Plant.AerationBasin.RecycleFlow, "QR, recycle flow rate, L/day")
    Call WriteFortranInput(f, .Plant.SecondaryClarifier.EffluentSolidsConc, "XSC, effluent solids concentration, mg/L")
    Call WriteFortranInput(f, .Plant.PrimaryClarifier.PercentageRemoval / 100#, "RE, solids removal efficiency in primary clarifier, fractional units (range of 0-1)")
    For i = 1 To N
      If (N = 1) Then
        Call WriteFortranInput(f, .Plant.AerationBasin.Volume, "V(i=" & Trim$(Str$(i)) & "), i=1...n, volume #i, liters")
      Else
        Call WriteFortranInput(f, .Plant.AerationBasin.CSTR.Volume(i - 1), "V(i=" & Trim$(Str$(i)) & "), i=1...n, volume #i, liters")
      End If
    Next i
    With .Plant.AerationBasin.BioTreat
      Call WriteFortranInput(f, .MaxGrowthRate, "MUM, maximum growth rate constant, day^(-1)")
      Call WriteFortranInput(f, .HalfVelocityConst, "KS, half velocity rate, mg/(L BOD5)")
      Call WriteFortranInput(f, .BacterialDecay, "KD, decay coefficient, day^(-1)")
      Call WriteFortranInput(f, .YieldCoeff, "Y, yield coefficient, (mg VSS)/(mg BOD5)")
      Call WriteFortranInput(f, .BOD5Conc, "S0, substrate concentration, mg/L")
    End With
    Call WriteFortranInput(f, .Plant.SolidsConc, "X0, influent solid concentration, mg/L")
    Call WriteFortranInput(f, 1000#, "ITMAX, maximum iteration count, unitless")
    Call WriteFortranInput(f, 0.00001, "ERRREL, convergence criteria, unitless")
    Close #f
  End With
exit_normally_ThisSub:
  ModelBIOCALC_WriteMainFile = True
  Exit Function
exit_err_ThisSub:
  ModelBIOCALC_WriteMainFile = False
  Exit Function
err_ThisSub:
  Call Show_Trapped_Error("ModelBIOCALC_WriteMainFile")
  GoTo exit_err_ThisSub
End Function
Sub ModelBIOCALC_WritePathFile()
'Dim f As Integer
'Dim fn_This As String
'Dim qq As String
'  qq = Chr$(34)
'  f = FreeFile
'  fn_This = MAIN_EXE_PATH & "\" & ModelBIOCALC_IN_PathFile
'  'fn_This = App.Path & "\" & ModelBIOCALC_IN_PathFile
'  Open fn_This For Output As #f
'  Print #f, "1"
'  Print #f, qq & ModelBIOCALC_IN_Main & qq
'  Print #f, qq & ModelBIOCALC_OUT_SuccessFlag & qq
'  Print #f, qq & ModelBIOCALC_OUT_Main & qq
'  Print #f, qq & ModelBIOCALC_OUT_CvsT & qq
'  Print #f, qq & ModelBIOCALC_OUT_EndCvsT & qq
'  Print #f, qq & ModelBIOCALC_OUT_CvsT_Details & qq
'  Close #f
End Sub








