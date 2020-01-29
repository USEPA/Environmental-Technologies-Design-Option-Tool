Attribute VB_Name = "Model_CalcBiomass___OLD___"
Option Explicit






Const Model_CalcBiomass_declarations_end = True


'CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC
'C C
'C  PROGRAM FOR CALCULATING SUBSTRATE AND BIOMASS CONCENTRATIONS     C
'C  IN A SERIES OF CSTRS, WITH OR WITHOUT STEP FEED                  C
'C C
'C  BASED ON EQUATIONS 5-14 AND 5-15, PAGE 352, DAVIS & CORNWELL,    C
'C  INTRODUCTION TO ENVIRONMENTAL ENGINEERING, MCGRAW-HILL, 1991.    C
'C C
'C  UNITS ARE CONSISTENT WITHIN CODE                                 C
'C C
'C  PARTIAL LIST OF VARIABLE DEFINITIONS                             C
'C --------------------------------------------------------------C
'C  ATV(N) = VOLUMES OF CSTRS (L)                                    C
'C  FFRACT(N) = STEP FEED FRACTIONS FED TO CSTR                      C
'C  ITERMX = MAXIMUM NUMBER OF ITERATIONS                            C
'C  KD = BACTERIAL DECAY RATE (1/day)                                C
'C  KS = HALF VELOCITY CONSTANT (mg BOD5/L)                          C
'C  MUM = MAXIMUM GROWTH RATE CONSTANT (1/day)                       C
'C  NN = NUMBER OF CSTRS                                             C
'C  NSF = STEP FEED OPTION: NSF=0: NO STEP FEED; NSF=1: STEP FEED    C
'C  Q0 = PLANT FLOWRATE (L/day)                                      C
'C  QR = RECYCLE FLOWRATE (L/day)                                    C
'C  QT = Q0+QR=TOTAL FLOW RATE (L/day)                               C
'C  RESMX = ITERATION CONVERGENCE CRITERIA                           C
'C  S0 = INFLUENT SUBSTRATE CONCENTRATION (mg BOD5/L)                C
'C  S(N) = SUBSTRATE CONCN. IN CSTR N (mg BOD5/L)                    C
'C  VT = TOTAL AERATION BASIN VOLUME (L)                             C
'C  XR = BIOMASS CONCN IN RECYCLE STREAM (mg VSS/L)                  C
'C  X(N) = BIOMASS CONCN IN CSTR N (mg VSS/L)                        C
'C  Y = YIELD COEFFCIENT (mg VSS/mg BOD5)                            C
'C C
'CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC
Function CalculateBioMass( _
    Temp_Plant As TYPE_PlantDiagram) _
    As Boolean
Dim X(20) As Double
Dim s(20) As Double
Dim XOLD(20) As Double
Dim SOLD(20) As Double
Dim QN1(20) As Double
Dim QFN(20) As Double
Dim QN(20) As Double
Dim FFRACT(20) As Double
Dim ATV(20) As Double
Dim NN%, NPC%, NAB%, NSC%, NSF%, N%
Dim ITERMX%, ITER%, i%
Dim FSUM#
Dim KX#, KD#, KS#
Dim MUM#
Dim Q0#, QW1#, QR#, QW#, QE1#, QD#, QRD#, QTD#
Dim RESMX#, RES#
Dim ST#, STX#, S0#
Dim TRES#
Dim VT#
Dim XT#, XTX#, XR#
Dim Y#
  On Error GoTo ERROR_CalculateBioMass
  With Temp_Plant
    'C
    'C.....PLANT FLOWRATE
    Q0# = .Flow
    'C
    'C
    'C.....NUMBER OF PRIMARY CLARIFIERS IN PARALELL
    NPC% = .PrimaryClarifier.Count
    'C
    'C.....PRIMARY WASTAGE FLOWRATE (FROM EACH PRIMARY CLARIFIER)
    QW1# = .PrimaryClarifier.WastageFlow
    'C
    'C
    'C.....NUMBER OF AERATION BASINS IN PARALELL
    NAB% = .AerationBasin.Count
    'C
    'C.....RECYCLE FLOWRATE
    QR# = .AerationBasin.RecycleFlow
    'C
    'C.....WASTAGE FLOWRATE
    QW# = .AerationBasin.WastageFlow
    'C
    'C.....TOTAL AERATION BASIN VOLUME
    VT# = .AerationBasin.Volume
    'C
    'C
    'C.....NUMBER OF CSTRS
    NN% = .AerationBasin.CSTR.Count
    'C
    'C.....STEP FEED OPTION: NSF=0: NO STEP FEED; NSF=1: STEP FEED
    NSF% = IIf(.AerationBasin.CSTR.UseStepFeed, 1, 0)
    For i% = 1 To .AerationBasin.CSTR.Count
      'C
      'C.....FEED FRACT
      FFRACT(i%) = .AerationBasin.CSTR.Feed(i% - 1)
      'C
      'C.....TANK VOLUMES
      ATV(i%) = .AerationBasin.CSTR.Volume(i% - 1)
    Next i%
    'C
    'C
    'C.....MAXIMUM GROWTH RATE CONSTANT
    MUM# = .AerationBasin.BioTreat.MaxGrowthRate
    'C
    'C.....HALF VELOCITY CONSTANT
    KS# = .AerationBasin.BioTreat.HalfVelocityConst
    'C
    'C.....BACTERIAL DECAY RATE
    KD# = .AerationBasin.BioTreat.BacterialDecay
    'C
    'C.....YIELD COEFFCIENT
    Y# = .AerationBasin.BioTreat.YieldCoeff
    'C
    'C.....INFLUENT SUBSTRATE CONCENTRATION
    S0# = .AerationBasin.BioTreat.BOD5Conc
    'C
    'C
    'C.....NUMBER OF SECONDARY CLARIFIERS IN PARALLEL
    NSC% = .SecondaryClarifier.Count
    'C
    'C
    'C.....MAXIMUM NUMBER OF ITERATIONS
    ITERMX% = 10000
    'C
    'C.....ITERATION CONVERGENCE CRITERIA
    RESMX# = 0.000001
    'C
    'C
    'C.....INITIAL GUESSES
    XR# = 2# * Y * S0
    s(1) = S0
    If (NN% > 2) Then
      For i% = 2 To NN% - 1
        s(i%) = 0.5 * S0#
      Next i%
    End If
    If (NN% > 1) Then s(NN%) = 0.2 * S0#
    'C
    'C
    'C.....CALCULATED PARAMETERS
    QE1# = (Q0# / CDbl(NPC%)) - QW1#
    QD# = QE1# * CDbl(NPC%) / CDbl(NAB%)
    QRD# = QR# * CDbl(NSC%) / CDbl(NAB%)
    QTD# = QD# + QRD#
    If (NN% = 1) Then NSF% = 0
    'C
    'C
    'C.....INITIALIZE
    ITER% = 0
    For i% = 1 To NN%
        X(i%) = 0#
        s(i%) = 0#
    Next i%
    'C
    'C
    'C.....BEGIN ITERATION LOOP
    Do
      ITER% = ITER% + 1
      For i% = 1 To NN%
        XOLD(i%) = X(i%)
        SOLD(i%) = s(i%)
      Next i%
      ST# = ((QD# * S0#) + (QRD# * s(NN%))) / QTD#
      If (ITER% > 1) Then XR# = X(NN%) * QTD# / (QW# + QRD#)
      XT# = QRD# * XR# / QTD#
      FSUM# = 0#
      'C
      'C
      i% = 1
      If (NSF% = 1) Then QFN(1) = QTD# * FFRACT(i%)
      If (NSF% = 0) Then QFN(1) = QTD#
      If (NSF% = 1) Then QN(1) = QTD# * FFRACT(i%)
      If (NSF% = 0) Then QN(1) = QTD#
      XTX# = XT#
      STX# = ST#
      KX# = (MUM# * s(1) / (KS# + s(1))) - KD#
      X(1) = (QFN(1) * XTX#) / (QN(1) - (ATV(i%) * KX#))
      KS# = MUM# * X(1) / (Y# * (KS# + s(1)))
      s(1) = (QFN(1) * STX#) / (QN(1) + (ATV(i%) * KS#))
      FSUM# = FFRACT(i%)
      'C
      'C
      If (NN% > 2) Then
        For i% = 2 To NN% - 1
          QN1(i%) = FSUM# * QTD#
          If (NSF% = 1) Then QFN(i%) = QTD# * FFRACT(i%)
          If (NSF% = 0) Then QFN(i%) = QTD# - QN1(i%)
          If (NSF% = 1) Then QN(i%) = QN1(i%) + QFN(i%)
          If (NSF% = 0) Then QN(i%) = QTD#
          If (NSF% = 1) Then XTX# = XT#
          If (NSF% = 0) Then XTX# = X(i% - 1)
          If (NSF% = 1) Then STX# = ST#
          If (NSF% = 0) Then STX# = s(i% - 1)
          KX# = (MUM# * s(i%) / (KS# + s(i%))) - KD#
          X(i%) = ((QN1(i%) * X(i% - 1)) + (QFN(i%) * XTX#)) / (QN(i%) - (ATV(i%) * KX#))
          KS# = MUM# * X(i%) / (Y# * (KS# + s(i%)))
          s(i%) = ((QN1(i%) * s(i% - 1)) + (QFN(i%) * STX#)) / (QN(i%) + (ATV(i%) * KS#))
          FSUM# = FSUM# + FFRACT(i%)
        Next i%
      End If
      'C
      'C
      'C.....N=NN
      If (NN% > 1) Then
        i% = NN%
        QN1(i%) = FSUM# * QTD#
        If (NSF% = 1) Then QFN(i%) = QTD# * FFRACT(i%)
        If (NSF% = 0) Then QFN(i%) = QTD# - QN1(i%)
        If (NSF% = 1) Then QN(i%) = QN1(i%) + QFN(i%)
        If (NSF% = 0) Then QN(i%) = QTD#
        If (NSF% = 1) Then XTX# = XT#
        If (NSF% = 0) Then XTX# = X(i% - 1)
        If (NSF% = 1) Then STX# = ST#
        If (NSF% = 0) Then STX# = s(i% - 1)
        KX# = (MUM# * s(i%) / (KS# + s(i%))) - KD#
        X(i%) = ((QN1(i%) * X(i% - 1)) + (QFN(i%) * XTX#)) / (QN(i%) - (ATV(i%) * KX#))
        KS# = MUM# * X(i%) / (Y# * (KS# + s(i%)))
        s(i%) = ((QN1(i%) * s(i% - 1)) + (QFN(i%) * STX#)) / (QN(i%) + (ATV(i%) * KS#))
      End If
      'C
      'C
      'C.....CALCULATE RESIDUAL AND CHECK
      TRES# = 0#
      For i% = 1 To NN%
        RES# = _
            Abs((SOLD(i%) - s(i%)) / s(i%)) + _
            Abs((XOLD(i%) - X(i%)) / X(i%))
        TRES# = TRES# + RES#
      Next i%
    Loop While ((TRES# > RESMX#) And (ITER% < ITERMX%))
    'C
    'C
    If ((TRES# > RESMX#) And (ITER% = ITERMX%)) Then
      MsgBox "Exceeded ITERMX without Converging."
      CalculateBioMass = False
      Exit Function
    End If
    CalculateBioMass = True
    'C
    'C
    'C.....WRITE RESULTS
    .AerationBasin.CSTR.UniformBioMass = False
    For i% = 0 To NN% - 1
      .AerationBasin.CSTR.BioMass(i%) = X(i% + 1)
    Next i%
  End With
  Exit Function
exit_ERROR_CalculateBioMass:
  CalculateBioMass = False
  Exit Function
ERROR_CalculateBioMass:
  MsgBox "Ran into a Problem with the BioMass Calculations." _
       + vbCrLf + vbCrLf _
       + "Error #:" + CStr(Err.Number) + vbCrLf _
       + "Desciption: " + Err.Description
  Resume exit_ERROR_CalculateBioMass
End Function




