Attribute VB_Name = "ModelPSDMInRoom"
Option Explicit

Global Const PSDMR_MODE_INROOM = 1
Global Const PSDMR_MODE_ALONE = 2
Dim intPSDMR_Mode As Integer

Const ModelPSDMInRoom_IN_PathFile = "PROOM1.IN"
Const ModelPSDMInRoom_IN_Main = "PROOM2.IN"
Const ModelPSDMInRoom_OUT_SuccessFlag = "PROOM1.OUT"
Const ModelPSDMInRoom_OUT_Main = "PROOM2.OUT"
Const ModelPSDMInRoom_OUT_CRvsT = "PROOM3.OUT"
Const ModelPSDMInRoom_OUT_CBvsT = "PROOM4.OUT"
''''Const ModelPSDMInRoom_OUT_CvsT = "PROOM3.OUT"

Const ModelPSDMInRoom_Version = 1#
'Const ModelPSDMInRoom_ExeName = "PROOM10C.EXE"
'Const ModelPSDMInRoom_ExeName = "PROOM11.EXE"
Const ModelPSDMInRoom_ExeName = "PROOM12.EXE"
''''Const ModelPSDMInRoom_EofTestMarker = 123456#
Const ModelPSDMInRoom_EofTestMarker = 12345.678

Const ModelPSDMInRoom_MXCOMP = 6
Const ModelPSDMInRoom_MAXPTS = 400
Const ModelPSDMInRoom_MAXDE = 750
Private Type ModelPSDMInRoom_Inputs_Type
  NUMB As Integer
  CHEMICALS(1 To ModelPSDMInRoom_MXCOMP, 1 To 16) As Double
  INITIAL_ROOM_CONC(1 To ModelPSDMInRoom_MXCOMP) As Double
  ADS_PROP(1 To 4) As Double
  C_PROP(1 To 3) As Double
  TT(1 To 3) As Double
  MXX As Integer
  NXX As Integer
  TotalAxialElementCount As Integer
  N_PW As Long
  NINI As Integer
  TINI(1 To ModelPSDMInRoom_MAXPTS) As Double
  IS_IN_ROOM As Integer     '1=PSDMR in Room, 0=PSDMR Alone
  ROOM_VOL As Double        'cm^3
  ROOM_FLOWRATE As Double   'cm^3/s
  ROOM_C0(1 To ModelPSDMInRoom_MXCOMP) As Double                'ug/L
  ROOM_EMIT(1 To ModelPSDMInRoom_MXCOMP) As Double              'ug/s
  RXN_RATE_CONSTANT(1 To ModelPSDMInRoom_MXCOMP) As Double
      '(i): first-order rate constant for destruction of chemical i, 1/s
  RXN_PRODUCT(1 To ModelPSDMInRoom_MXCOMP) As Integer
      '(i): index of chemical that is produced through destruction of chemical i (or 0 if none), unitless
  RXN_RATIO(1 To ModelPSDMInRoom_MXCOMP) As Double
      '(i): number of moles of chemical RXN_PRODUCT(i) produced by destruction of 1 mole of chemical i
  FN_MASSBAL_OUT As String
  FN_CR_OUT As String
  FN_CB_OUT As String
End Type
Dim ModelPSDMInRoom_Inputs As ModelPSDMInRoom_Inputs_Type
'Private Type ModelPSDMInRoom_Inputs2_Type
'  CINI(1 To ModelPSDMInRoom_MXCOMP, 1 To ModelPSDMInRoom_MAXPTS) As Double
'End Type
Dim ModelPSDMInRoom_Inputs_CINI(1 To ModelPSDMInRoom_MXCOMP, 1 To ModelPSDMInRoom_MAXPTS) As Double

Private Type ModelPSDMInRoom_Outputs_Type
  VARS1(1 To 15) As Double
  VARS2(1 To ModelPSDMInRoom_MXCOMP, 1 To 19) As Double
  NITP As Integer
  T(1 To ModelPSDMInRoom_MAXPTS) As Double
  NFLAG As Integer
End Type
Dim ModelPSDMInRoom_Outputs As ModelPSDMInRoom_Outputs_Type
'Private Type ModelPSDMInRoom_Outputs2_Type
'  CPVB(1 To ModelPSDMInRoom_MXCOMP, 1 To ModelPSDMInRoom_MAXPTS) As Double
'End Type
Dim ModelPSDMInRoom_Outputs_CPVB(1 To ModelPSDMInRoom_MXCOMP, 1 To ModelPSDMInRoom_MAXPTS) As Double





Const ModelPSDMInRoom_declarations_end = True


Sub ModelPSDMInRoom_Go( _
    IN_intPSDMR_Mode As Integer)
Dim Failed As Boolean
  intPSDMR_Mode = IN_intPSDMR_Mode
  Call ModelPSDMInRoom_RemoveLinkFiles
  Call ModelPSDMInRoom_WritePathFile
  Call ModelPSDMInRoom_WriteMainFile(Failed)
  If (Failed) Then Exit Sub
  Call ModelPSDMInRoom_CallEXE
  Call ModelPSDMInRoom_ProcessOutput
  If (ModelIO_IsKeepTempFiles() = False) Then
    Call ModelPSDMInRoom_RemoveLinkFiles
  End If
End Sub


Sub ModelPSDMInRoom_RemoveLinkFiles()
  Call KillFile_If_Exists(Exe_Path & "\" & ModelPSDMInRoom_IN_PathFile)
  Call KillFile_If_Exists(Exe_Path & "\" & ModelPSDMInRoom_IN_Main)
  Call KillFile_If_Exists(Exe_Path & "\" & ModelPSDMInRoom_OUT_SuccessFlag)
  Call KillFile_If_Exists(Exe_Path & "\" & ModelPSDMInRoom_OUT_Main)
  Call KillFile_If_Exists(Exe_Path & "\" & ModelPSDMInRoom_OUT_CRvsT)
  Call KillFile_If_Exists(Exe_Path & "\" & ModelPSDMInRoom_OUT_CBvsT)
End Sub
Sub ModelPSDMInRoom_CallEXE()
Dim CmdLine As String
  Call ChangeDir_Exes
  CmdLine = ModelPSDMInRoom_ExeName
  Call ModelIO_Timer_Start
  Call FortranLink_ExecAndWaitForProcess(CmdLine)
  Call ModelIO_Timer_End
  Call ChangeDir_Main
End Sub
Sub ModelPSDMInRoom_ProcessOutput()
Dim f As Integer
Dim fn_This As String
Dim NFLAG As Integer
Dim DummyStr1 As String
Dim temp As String
Dim i As Integer
Dim J As Integer
Dim k As Integer
Dim EOFTESTMARKER As Double
Dim Flag05 As Boolean
Dim Flag50 As Boolean
Dim Flag95 As Boolean
Dim MI As ModelPSDMInRoom_Inputs_Type
Dim MO As ModelPSDMInRoom_Outputs_Type
  MI = ModelPSDMInRoom_Inputs
  'READ SUCCESS FLAG OUTPUT FILE.
  f = FreeFile
  fn_This = Exe_Path & "\" & ModelPSDMInRoom_OUT_SuccessFlag
  If (Not FileExists(fn_This)) Then
    Call Show_Error("Unable to find output file: Calculations failed.")
    Exit Sub
  End If
  Open fn_This For Input As #f
  Line Input #f, DummyStr1
  Input #f, NFLAG
  Close #f
  If (NFLAG <> 0) Then
    Select Case NFLAG
      Case 15
        temp = "WARNING...  T + H = T ON NEXT STEP"
      Case 105
        temp = "KFLAG = -1 FROM INTEGRATOR"
      Case 115
        temp = "H HAS BEEN REDUCED TO AND STEP WILL BE RETRIED"
      Case 155
        temp = "PROBLEM APPEARS UNSOLVABLE WITH GIVEN INPUT"
      Case 205
        temp = "THE REQUESTED ERROR IS SMALLER THAN CAN BE HANDLED"
      Case 255
        temp = "INTEGRATION HALTED BY DRIVER EPS TOO SMALL TO BE ATTAINED FOR THE MACHINE PRECISION"
      Case 305
        temp = "CORRECTOR CONVERGENCE COULD NOT BE ACHIEVED"
      Case 405
        temp = "ILLEGAL INPUT... EPS < 0"
      Case 415
        temp = "ILLEGAL INPUT... N <= 0"
      Case 425
        temp = "ILLEGAL INPUT... (T0-TOUT)*H >= 0"
      Case 435
        temp = "ILLEGAL INPUT... INDEX"
      Case 445
        temp = "INTERPOLATION WAS DONE AS ON NORMAL RETURN; DESIRED PARAMETER CHANGES WERE NOT MADE."
      Case Else
        temp = "Unknown Error"
    End Select
    temp = "Error #" & Trim$(Str$(NFLAG)) & ": " & temp
    Call Show_Error("The PSDM failed to converge." & vbCrLf & temp)
    Exit Sub
  Else
    Call Show_Message( _
        "PSDM Model Calculations Complete." & _
        vbCrLf & _
        vbCrLf & _
        ModelIO_Timer_SummaryMsg)
  End If
  ''READ MAIN OUTPUT FILE.
  'fn_This = Exe_Path & "\" & ModelPSDMInRoom_OUT_Main
  'Open fn_This For Input As #f
  'Line Input #f, DummyStr1
  'For i = 1 To 15
  '  Input #f, MO.VARS1(i)
  'Next i
  'Line Input #f, DummyStr1
  'For i = 1 To MI.NUMB
  '  For j = 1 To 19
  '    Input #f, MO.VARS2(i, j)
  '  Next j
  'Next i
  'Line Input #f, DummyStr1
  'Input #f, MO.NFLAG
  'Line Input #f, DummyStr1
  'Input #f, EOFTESTMARKER
  'If (False = ModelIO_DoNumberCheck(EOFTESTMARKER, ModelPSDMInRoom_EofTestMarker)) Then
  '  Call Show_Error("The model calculations failed: invalid file format (EOF marker).")
  '  Exit Sub
  'End If
  'Close #f
  '
  '//////////////// READ C-vs-t OUTPUT FILE. ////////////////////////////////////////////////////////////////////////////////////////////////
  '
    Select Case intPSDMR_Mode
      '
      '//////////////////////////////////////////////////////////////////////////////////////////
      '////////////   PSDMR-IN-ROOM
      '//////////////////////////////////////////////////////////////////////////////////////////
      Case PSDMR_MODE_INROOM:
        fn_This = Exe_Path & "\" & ModelPSDMInRoom_OUT_CRvsT
      '
      '//////////////////////////////////////////////////////////////////////////////////////////
      '////////////   PSDMR ALONE
      '//////////////////////////////////////////////////////////////////////////////////////////
      Case PSDMR_MODE_ALONE:
        fn_This = Exe_Path & "\" & ModelPSDMInRoom_OUT_CBvsT
    End Select
    ''''fn_This = Exe_Path & "\" & ModelPSDMInRoom_OUT_CRvsT
    f = FreeFile
    Open fn_This For Input As #f
Dim Found As Boolean
Dim ThisRow As Integer
Dim NumArgsExpected As Integer
Dim NumArgsGot As Integer
Dim ThisLine As String
Dim Dummy As String
Dim NumRows As Integer
    Found = False
    ThisRow = 1
    NumArgsExpected = 1 + Number_Component_PFPSDM
    Do While (1 = 1)
      If (EOF(f)) Then
        'UNABLE TO FIND NFLAG!  ASSUME A PROBLEM OCCURRED.
        Exit Do
      End If
      Line Input #1, ThisLine
      ThisLine = Trim$(ThisLine)
      If (UCase$(Trim$(ThisLine)) = UCase$(Trim$("END_OF_DATA"))) Then
        Found = True
        Exit Do
      End If
      ThisLine = Parser2_RemoveDuplicateSeparators(" ", ThisLine)
      NumArgsGot = Parser2_GetNumArgs(" ", ThisLine)
      
      '///////////////////////////////////////////////////////////////////////
      '///if we bypass the error message the output file for PSDMR will work
      '///for more than two components.(Sinan, 08/29/2006)
      '//////////////////////////////////////////////////////////////////////
      If (NumArgsGot <> NumArgsExpected) Then
        'UNEXPECTED NUMBER OF ARGUMENTS!  EXIT.
       Close #f
        Call Show_Error("The model output file `" & fn_This & _
            "` was corrupted (unexpected number of arguments on line #" & _
            Trim$(Str$(ThisRow)) & ".  Calculations failed.")
        Exit Sub
      End If
      '//////////////////////////////////////////////////////////////////////
      
      
      Call Parser2_GetArg(" ", ThisLine, 1, Dummy)
      MO.T(ThisRow) = CDbl(Val(Dummy))
      For i = 1 To Number_Component_PFPSDM
        Call Parser2_GetArg(" ", ThisLine, 1 + i, Dummy)
        ModelPSDMInRoom_Outputs_CPVB(i, ThisRow) = CDbl(Val(Dummy))
      Next i
      ThisRow = ThisRow + 1
    Loop
    If (Not Found) Then
      'ERROR -- UNABLE TO FIND NFLAG!
      Close #f
      Call Show_Error("The model output file `" & fn_This & _
          "` was corrupted (unable to find the end-of-data designator).  " & _
          "Calculations failed.")
      Exit Sub
    End If
    Close #f
    NumRows = ThisRow - 1
    MO.NITP = NumRows
    ''''GoTo SkipPastOldCode1

    'CP(i,j) -- i=component, j=ThisRow
    'T(i,1) -- i=ThisRow
    '
    'T(ThisRow,1), CP(1,ThisRow), ... CP(NCOMP,ThisRow)


'f = FreeFile
'Open "c:\psdm1.txt" For Output As #f
'For J = 1 To NumRows
'Print #f, MO.T(J), ModelPSDMInRoom_Outputs_CPVB(1, J)
'Next J
'Close #f
  
'  fn_This = Exe_Path & "\" & ModelPSDMInRoom_OUT_CRvsT
'  Open fn_This For Input As #f
'  Line Input #f, DummyStr1
'  Input #f, MO.NITP
'  Line Input #f, DummyStr1
'  For i = 1 To MO.NITP
'    Input #f, MO.T(i)
'  Next i
'  Line Input #f, DummyStr1
'  For i = 1 To MI.NUMB
'    For j = 1 To MO.NITP
'      'Input #f, MO.CPVB(i, j)
'      Input #f, ModelPSDMInRoom_Outputs_CPVB(i, j)
'    Next j
'  Next i
'  Line Input #f, DummyStr1
'  Input #f, EOFTESTMARKER
'  If (False = ModelIO_DoNumberCheck(EOFTESTMARKER, ModelPSDMInRoom_EofTestMarker)) Then
'    Call Show_Error("The model calculations failed: invalid file format (EOF marker).")
'    Exit Sub
'  End If
'  Close #f
  ModelPSDMInRoom_Outputs = MO
  '
  '//////////////// TRANSFER OUTPUT DATA TO MORE PERMANENT MEMORY. ////////////////////////////////////////////////////////////////////////////////
  '
  Results.is_psdm_in_room_model = True
  Results.int_Which_PSDMR_Model = intPSDMR_Mode
  Results.npoints = MO.NITP
  Results.NComponent = MI.NUMB
  Results.Bed = Bed
  Results.Carbon = Carbon
  For i = 1 To 15
    PSDM_Inputs.VARS1(i) = MO.VARS1(i)
  Next i
  PSDM_Inputs.VARS1(8) = SF() * 264.17205 * 60 / 10.76391         'Convert m/s to gal/min-ft^2.
  PSDM_Inputs.VARS1(11) = Re()
  PSDM_Inputs.VARS1(12) = Bed.WaterDensity
  PSDM_Inputs.VARS1(13) = Bed.WaterViscosity
  For i = 1 To Number_Component_PFPSDM
    For J = 1 To 18
      PSDM_Inputs.VARS2(i, J) = MO.VARS2(i, J)
    Next J
    PSDM_Inputs.VARS2(i, 6) = Diffl(i)
    PSDM_Inputs.VARS2(i, 18) = SC(i)
    J = Component_Index_PFPSDM(i)
    PSDM_Inputs.VARS2(i, 19) = Component(J).SPDFR
  Next i
  '
  '//////////////// HANDLE MISCELLANEOUS PSDMR STUFF. ///////////////////////////////////////////////////////////////////
  '
  ' TRANSFER Cr,ss VALUES TO Results STRUCTURE.
  ' IS ANY Cr,ss VALUE VERY CLOSE TO ZERO?
  '
  Dim AnyCrCloseToZero As Integer
  AnyCrCloseToZero = False
  For i = 1 To Number_Component_PFPSDM
    k = Component_Index_PFPSDM(i)
    Results.psdmroom_Crss(i) = RoomParams.ROOM_SS_VALUE(k)
    If (Abs(Results.psdmroom_Crss(i)) < 1E-20) Then
      AnyCrCloseToZero = True
    End If
  Next i
  Results.AnyCrCloseToZero = AnyCrCloseToZero
  Select Case intPSDMR_Mode
    '
    '//////////////////////////////////////////////////////////////////////////////////////////
    '////////////   PSDMR-IN-ROOM
    '//////////////////////////////////////////////////////////////////////////////////////////
    Case PSDMR_MODE_INROOM:
      '
      ' FOR THE SPECIAL CASE OF THE PSDMR-IN-ROOM, CONVERT ALL
      ' Cr TO Cr/Cr,ss VALUES IF NONE OF THE VALUES
      ' OF Cr,ss ARE CLOSE TO ZERO.
      '
      If (AnyCrCloseToZero = False) Then
        For i = 1 To Number_Component_PFPSDM
          For J = 1 To MO.NITP
            ModelPSDMInRoom_Outputs_CPVB(i, J) = _
                ModelPSDMInRoom_Outputs_CPVB(i, J) / _
                Results.psdmroom_Crss(i)
          Next J
        Next i
      End If
    '
    '//////////////////////////////////////////////////////////////////////////////////////////
    '////////////   PSDMR ALONE
    '//////////////////////////////////////////////////////////////////////////////////////////
    Case PSDMR_MODE_ALONE:
      Results.AnyCrCloseToZero = True     'TELL PLOTTER UNITS ARE UG/L!
      '
      ' CONVERT FROM UG/L TO UG/L (I.E. NO CHANGE!).
      '
      For i = 1 To Number_Component_PFPSDM
        For J = 1 To MO.NITP
          ModelPSDMInRoom_Outputs_CPVB(i, J) = _
              ModelPSDMInRoom_Outputs_CPVB(i, J) * 1#
        Next J
      Next i
  End Select
  '
  '//////////////// DETERMINE 5%, 50%, AND 95% SATURATION TIMES. ////////////////////////////////////////////////////////////////
  '
  Flag05 = True
  Flag50 = True
  Flag95 = True
  ReDim BrokeThrough(1 To Number_Component_PFPSDM) As Integer
  Dim IsFoulingCase As Integer
  'ReDim NumPoints_Before_BrokeThrough(Number_Component_PFPSDM) As Integer
  For i = 1 To Number_Component_PFPSDM
    BrokeThrough(i) = False
    'NumPoints_Before_BrokeThrough(i) = -1
    Results.NumPoints_Before_ThroughPut_100(i) = MO.NITP
  Next i
  IsFoulingCase = False
  For i = 1 To Number_Component_PFPSDM
    J = Component_Index_PFPSDM(i)
    If (Component(J).K_Reduction) Then
      IsFoulingCase = True
    End If
  Next i
Dim DoNotPrematurelyEndFoulingPlot As Boolean
  DoNotPrematurelyEndFoulingPlot = True
  For i = 1 To Number_Component_PFPSDM
    Results.Component(i) = Component(Component_Index_PFPSDM(i))
    For J = 1 To MO.NITP
     If (DoNotPrematurelyEndFoulingPlot = False) And _
        ((((IsFoulingCase) And (ModelPSDMInRoom_Outputs_CPVB(i, J) > 0.9995)) Or (BrokeThrough(i)))) Then
       '---- Stop the plot as soon as C/C0>=0.9995; this is accomplished
       '.... by setting .CP = -10000#, which tells the plotting routine to
       '.... stop plotting.
       Results.CP(i, J) = -10000#
       If (Not BrokeThrough(i)) Then
         Results.NumPoints_Before_ThroughPut_100(i) = J - 1
       End If
       BrokeThrough(i) = True
       ''---- Assume C/C0=1.0 as soon as C/C0>=0.9995
       'Results.CP(i, j) = 1#
       'If (Not BrokeThrough(i)) Then
       '  Results.NumPoints_Before_ThroughPut_100(i) = j - 1
       'End If
       'BrokeThrough(i) = True
       ''NumPoints_Before_BrokeThrough(i) = j - 1
     Else
       Results.CP(i, J) = ModelPSDMInRoom_Outputs_CPVB(i, J)
     End If
     If J > 2 Then
       If (ModelPSDMInRoom_Outputs_CPVB(i, J) >= 0.05) And (ModelPSDMInRoom_Outputs_CPVB(i, J - 1) < 0.05) And Flag05 Then
          Results.ThroughPut_05(i).T = _
              (MO.T(J) - MO.T(J - 1)) / _
              (ModelPSDMInRoom_Outputs_CPVB(i, J) - ModelPSDMInRoom_Outputs_CPVB(i, J - 1)) * (0.05 - ModelPSDMInRoom_Outputs_CPVB(i, J - 1)) + MO.T(J - 1)
          Results.ThroughPut_05(i).C = _
              ((ModelPSDMInRoom_Outputs_CPVB(i, J) - ModelPSDMInRoom_Outputs_CPVB(i, J - 1)) / (MO.T(J) - MO.T(J - 1)) * _
              (Results.ThroughPut_05(i).T - MO.T(J - 1)) + ModelPSDMInRoom_Outputs_CPVB(i, J - 1)) * _
              Component(Component_Index_PFPSDM(i)).InitialConcentration
          Flag05 = False
       End If
       If (ModelPSDMInRoom_Outputs_CPVB(i, J) >= 0.5) And (ModelPSDMInRoom_Outputs_CPVB(i, J - 1) < 0.5) And Flag50 Then
          Results.ThroughPut_50(i).T = _
              (MO.T(J) - MO.T(J - 1)) / _
              (ModelPSDMInRoom_Outputs_CPVB(i, J) - ModelPSDMInRoom_Outputs_CPVB(i, J - 1)) * (0.5 - ModelPSDMInRoom_Outputs_CPVB(i, J - 1)) + MO.T(J - 1)
          Results.ThroughPut_50(i).C = _
              ((ModelPSDMInRoom_Outputs_CPVB(i, J) - ModelPSDMInRoom_Outputs_CPVB(i, J - 1)) / (MO.T(J) - MO.T(J - 1)) * _
              (Results.ThroughPut_50(i).T - MO.T(J - 1)) + ModelPSDMInRoom_Outputs_CPVB(i, J - 1)) * _
              Component(Component_Index_PFPSDM(i)).InitialConcentration
          Flag50 = False
          If Flag05 Then
            Results.ThroughPut_05(i).T = -1#
            Results.ThroughPut_05(i).C = -1#
            Flag05 = False
          End If
       End If
       If (ModelPSDMInRoom_Outputs_CPVB(i, J) >= 0.95) And (ModelPSDMInRoom_Outputs_CPVB(i, J - 1) < 0.95) And Flag95 Then
          Results.ThroughPut_95(i).T = _
              (MO.T(J) - MO.T(J - 1)) / _
              (ModelPSDMInRoom_Outputs_CPVB(i, J) - ModelPSDMInRoom_Outputs_CPVB(i, J - 1)) * (0.95 - ModelPSDMInRoom_Outputs_CPVB(i, J - 1)) + MO.T(J - 1)
          Results.ThroughPut_95(i).C = _
              ((ModelPSDMInRoom_Outputs_CPVB(i, J) - ModelPSDMInRoom_Outputs_CPVB(i, J - 1)) / (MO.T(J) - MO.T(J - 1)) * _
              (Results.ThroughPut_95(i).T - MO.T(J - 1)) + ModelPSDMInRoom_Outputs_CPVB(i, J - 1)) * _
              Component(Component_Index_PFPSDM(i)).InitialConcentration
          Flag95 = False
          If Flag50 Then
            Results.ThroughPut_50(i).T = -1#
            Results.ThroughPut_50(i).C = -1#
            Flag50 = False
          End If
          If Flag05 Then
            Results.ThroughPut_05(i).T = -1#
            Results.ThroughPut_05(i).C = -1#
            Flag05 = False
          End If
       End If
     End If
    Next J
    If Flag95 Then
       Results.ThroughPut_95(i).T = -1#
       Results.ThroughPut_95(i).C = -1#
       Flag95 = False
    End If
    If Flag50 Then
       Results.ThroughPut_50(i).T = -1#
       Results.ThroughPut_50(i).C = -1#
       Flag50 = False
    End If
    If Flag05 Then
       Results.ThroughPut_05(i).T = -1#
       Results.ThroughPut_05(i).C = -1#
       Flag05 = False
    End If
    Flag05 = True  'Set these flags to True such that
    Flag50 = True  ' Results.ThroughPut_??(I).T and Results.ThroughPut_??(I).C
    Flag95 = True  ' are calculated for the next compound
  Next i
  For i = 1 To Number_Points_Max
    Results.T(i) = MO.T(i)
  Next i
  '
  '//////////////// ENABLE RESULTS MENU COMMANDS. ////////////////////////////////////////////////////////////////////////////////
  '
  frmMain.mnuResultsItem(0).Enabled = True
  'frmMain.mnuResultsItem(10).Enabled = True: MsgBox "Note, remove this line!"
  'If (NData_Points > 0) Then
  '  frmMain.mnuResultsItem(3).Enabled = True
  'End If
End Sub
Sub ModelPSDMInRoom_WriteMainFile(Failed As Boolean)
Dim MI As ModelPSDMInRoom_Inputs_Type
Dim i As Integer
Dim i_ As Integer
Dim J As Integer
Dim Number_Equations As Integer
Dim WorkSpace_Size As Long
Dim msg As String
Dim f As Integer
Dim fn_This As String
Dim ThisLine As String
Dim intThis As Integer
Dim strTemp1 As String
Dim Do_Fouling_For_This_Component As Boolean
  Failed = False
  'CALCULATE WORKSPACE SIZE.
  Number_Equations = Number_Component_PFPSDM * (MC * (NC + 1) - 1)
  If Number_Equations > Max_Equations_DGEAR Then
    msg = "Maximum number of equations PSDM can solve = " & Str$(Max_Equations_DGEAR) & vbCrLf
    msg = msg & "Current number of equations specified for PSDM to solve = " & Str$(Number_Equations) & vbCrLf & vbCrLf
    msg = msg & "(No. of Equations PSDM Must Solve) = NCOMP*(MC*(NC+1)-1)" & vbCrLf & vbCrLf
    msg = msg & "Please ensure that this number does not exceed the maximum." & vbCrLf & vbCrLf
    msg = msg & "Note:  " & vbCrLf
    msg = msg & "    NCOMP = Number of Components = " & Str$(Number_Component_PFPSDM) & vbCrLf
    msg = msg & "    MC = Number of Axial Collocation Points = " & Str$(MC) & vbCrLf
    msg = msg & "    NC = Number of Radial Collocation Points = " & Str$(NC) & vbCrLf
    Call Show_Error(msg)
    Failed = True
    Exit Sub
  End If
  WorkSpace_Size = Number_Equations ^ 2 + 2 * Number_Equations
  'PREPARE INPUTS.
  MI.NUMB = Number_Component_PFPSDM
  For i = 1 To MI.NUMB
    J = Component_Index_PFPSDM(i)
    MI.CHEMICALS(i, 1) = Component(J).MW
    'CONVERT Co FROM mg/L TO ug/L.
    MI.CHEMICALS(i, 2) = Component(J).InitialConcentration * 1000#
    MI.CHEMICALS(i, 3) = Component(J).MolarVolume
    'CONVERT K FROM (mg/g)*(L/mg)^(1/n) to (umol/g)*(L/umol)^(1/n).
    MI.CHEMICALS(i, 4) = Component(J).Use_K * (1000# / Component(J).MW) ^ (1# - Component(J).Use_OneOverN)
    MI.CHEMICALS(i, 5) = Component(J).Use_OneOverN
    MI.CHEMICALS(i, 6) = Component(J).kf
    MI.CHEMICALS(i, 7) = Component(J).Ds
    MI.CHEMICALS(i, 8) = Component(J).Dp
    '
    ' HANDLE FOULING -- NOTE, THIS IS HANDLED DIFFERENTLY
    ' THAN IN THE STANDARD PSDM.
    '
    Do_Fouling_For_This_Component = False
    If (Component(J).K_Reduction) And _
        (Bed.Water_Correlation.Coeff(1) <> 1# And Bed.Water_Correlation.Coeff(2) <> 0# And _
        Bed.Water_Correlation.Coeff(3) <> 0# And Bed.Water_Correlation.Coeff(4) <> 0#) Then
      Do_Fouling_For_This_Component = True
    End If
    If (Do_Fouling_For_This_Component = True) Then
      If (Bed.Phase = 0) Then
        ' LIQUID PHASE : DO NOTHING.
      End If
      If (Bed.Phase = 1) Then
        ' GAS PHASE : SHOW MESSAGE.
        Call Show_Message("Gas-phase fouling (K reduction) correlation is active " & _
            "for chemical #" & Trim$(Str$(J)) & ": `" & Trim$(Component(J).Name) & "`")
      End If
    End If
    If (Do_Fouling_For_This_Component = True) Then
      MI.CHEMICALS(i, 9) = Bed.Water_Correlation.Coeff(1) * Component(J).Correlation.Coeff(1) + _
          Component(J).Correlation.Coeff(2)
      MI.CHEMICALS(i, 10) = Bed.Water_Correlation.Coeff(2) * Component(J).Correlation.Coeff(1)
      MI.CHEMICALS(i, 11) = Bed.Water_Correlation.Coeff(3) * Component(J).Correlation.Coeff(1)
      MI.CHEMICALS(i, 12) = Bed.Water_Correlation.Coeff(4) * Component(J).Correlation.Coeff(1)
    Else
      MI.CHEMICALS(i, 9) = 1#
      MI.CHEMICALS(i, 10) = 0#
      MI.CHEMICALS(i, 11) = 0#
      MI.CHEMICALS(i, 12) = 0#
    End If
    '
    ' END OF FOULING HANDLING CODE.
    '
    MI.CHEMICALS(i, 13) = Component(J).Tortuosity
    If ((Component(J).Constant_Tortuosity) And (Component(J).Use_Tortuosity_Correlation)) Then
      MI.CHEMICALS(i, 14) = 2#
      MI.CHEMICALS(i, 15) = 0#
    Else
      If (Component(J).Use_Tortuosity_Correlation) Then
        MI.CHEMICALS(i, 14) = 0.334
        MI.CHEMICALS(i, 15) = 0.00000661
      Else
        MI.CHEMICALS(i, 14) = 2#
        MI.CHEMICALS(i, 15) = 0#
      End If
    End If
    MI.CHEMICALS(i, 16) = 100000#       'in minutes
    '
    '    ON NEXT LINE, CONVERT mg/L TO ug/L
    MI.INITIAL_ROOM_CONC(i) = RoomParams.INITIAL_ROOM_CONC(i) * 1000#
  Next i
  'NOTE: ADJUSTMENT OF LENGTH AND DIAMETER IS NOW PERFORMED
  'INSIDE THE FORTRAN MODULE.
  ''''MI.ADS_PROP(1) = Bed.Length / CDbl(Bed.NumberOfBeds)
  MI.ADS_PROP(1) = Bed.length
  MI.ADS_PROP(2) = Bed.Diameter
  ''''MI.ADS_PROP(3) = Bed.Weight / CDbl(Bed.NumberOfBeds)
  MI.ADS_PROP(3) = Bed.Weight
  MI.ADS_PROP(4) = Bed.Flowrate
  'If (only_make_input_file) Then
  '  ADS_PROP(1) = ADS_PROP(1) * CDbl(Bed.NumberOfBeds)
  '  ADS_PROP(3) = ADS_PROP(3) * CDbl(Bed.NumberOfBeds)
  'End If
  MI.C_PROP(1) = Carbon.Porosity
  MI.C_PROP(2) = Carbon.Density
  MI.C_PROP(3) = Carbon.ParticleRadius * 100#      'To convert in cm
  MI.TT(1) = TimeP.End
  'Test value of Tinit
  If (TimeP.Init <= 0#) Then
    MI.TT(2) = 0.0001
  Else
    MI.TT(2) = TimeP.Init
  End If
  MI.TT(3) = TimeP.Step
  MI.MXX = MC
  MI.NXX = NC
  MI.TotalAxialElementCount = Bed.NumberOfBeds
  ''''MI.N_PW = WorkSpace_Size
  MI.NINI = Number_Influent_Points
  For J = 1 To MI.NINI
    MI.TINI(J) = T_Influent(J)
    For i = 1 To MI.NUMB
      'CONVERT FROM mg/L TO ug/L.
      ModelPSDMInRoom_Inputs_CINI(i, J) = C_Influent(Component_Index_PFPSDM(i), J) * 1000#
    Next i
  Next J
  '
  ' DETERMINE WHICH TYPE OF PSDMR MODEL.
  Select Case intPSDMR_Mode
    Case PSDMR_MODE_INROOM:
      MI.IS_IN_ROOM = 1
    Case PSDMR_MODE_ALONE:
      MI.IS_IN_ROOM = 0
  End Select
  '
  ' ON NEXT LINE, CONVERT m^3 TO cm^3.
  MI.ROOM_VOL = RoomParams.ROOM_VOL * 1000000#
  '
  ' ON NEXT LINE, CONVERT m^3/s TO cm^3/s.
  MI.ROOM_FLOWRATE = RoomParams.ROOM_FLOWRATE * 1000000#
  For i = 1 To MI.NUMB
    '
    ' ON NEXT LINE, CONVERT mg/L TO ug/L.
    MI.ROOM_C0(i) = RoomParams.ROOM_C0(i) * 1000#
    MI.ROOM_EMIT(i) = RoomParams.ROOM_EMIT(i)
    MI.RXN_RATE_CONSTANT(i) = RoomParams.RXN_RATE_CONSTANT(i)
    MI.RXN_PRODUCT(i) = RoomParams.RXN_PRODUCT(i)
    MI.RXN_RATIO(i) = RoomParams.RXN_RATIO(i)
  Next i
  MI.FN_MASSBAL_OUT = "proom_massbal.out"
  MI.FN_CR_OUT = ModelPSDMInRoom_OUT_CRvsT
  MI.FN_CB_OUT = ModelPSDMInRoom_OUT_CBvsT
  ''''MI.FN_CB_OUT = "proom_cb.out"
  '
  ' ////////////////// WRITE INPUT FILE. ////////////////////////////////////////////////////////////////////////////////////////////////////////////
  '
  f = FreeFile
  fn_This = Exe_Path & "\" & ModelPSDMInRoom_IN_Main
  Open fn_This For Output As #f
  Call WriteFortranInput(f, 0, "NOTE1: ")
  Call WriteFortranInput(f, 0, "NOTE2: ")
  Call WriteFortranInput(f, 0, "NOTE3: ")
  Call WriteFortranInput(f, 0, "NOTE4: ")
  Print #f, String$(75, "-")
  Call WriteFortranInput(f, MI.ADS_PROP(1), "ADS_PROP(1), length, m")
  Call WriteFortranInput(f, MI.ADS_PROP(2), "ADS_PROP(2), diameter, m")
  Call WriteFortranInput(f, MI.ADS_PROP(3), "ADS_PROP(3), weight of adsorbent, kg")
  Call WriteFortranInput(f, MI.ADS_PROP(4), "ADS_PROP(4), influent flow rate, m^3/s")
  Call WriteFortranInput(f, MI.C_PROP(1), "C_PROP(1), void fraction of particle, -")
  Call WriteFortranInput(f, MI.C_PROP(2), "C_PROP(2), apparent density, g/cm^3")
  Call WriteFortranInput(f, MI.C_PROP(3), "C_PROP(3), particle radius, cm")
  Call WriteFortranInput(f, MI.MXX, "MXX: number of axial collocation points")
  Call WriteFortranInput(f, MI.NXX, "NXX: number of radial collocation points")
  Call WriteFortranInput(f, MI.NUMB, "NUMB: number of chemicals")
  Call WriteFortranInput(f, MI.NINI, "NINI: number of influent points")
  Call WriteFortranInput(f, MI.TotalAxialElementCount, "NUMBED: current axial element number in series to simulate")
  Call WriteFortranInput(f, 1, "BEDSIMTYPE: 0=simulate only this axial element, 1=simulate NUMBED number of axial elements all at once")
  Call WriteFortranInput(f, 0, "ISDBUG: debug setting, 0=no debugging")
  Call WriteFortranInput(f, MI.TT(1), "TT(1), time to end simulation, min")
  Call WriteFortranInput(f, MI.TT(2), "TT(2), time to begin simulation, min")
  Call WriteFortranInput(f, MI.TT(3), "TT(3), time step, min")
  Print #f, String$(75, "-")
  Call WriteFortranInput(f, MI.IS_IN_ROOM, "IS_IN_ROOM: equal to 1 if the filter is within a room (treated as a CSTR); extra parameters below if equal to 1.")
  ''''Call WriteFortranInput(f, 1, "IS_IN_ROOM: equal to 1 if the filter is within a room (treated as a CSTR); extra parameters below if equal to 1.")
  Call WriteFortranInput(f, MI.ROOM_VOL, "ROOM_VOL, volume of room, cm^3")
  Call WriteFortranInput(f, MI.ROOM_FLOWRATE, "ROOM_FLOWRATE, volumetric flow through room, cm^3/s")
  For i = 1 To MI.NUMB
    Call WriteFortranInput(f, MI.ROOM_C0(i), "ROOM_C0(i), component #i concentration influent to room, ug/L")
    Call WriteFortranInput(f, MI.ROOM_EMIT(i), "ROOM_EMIT(i), component #i emission rate in room, ug/s")
    Call WriteFortranInput(f, MI.RXN_RATE_CONSTANT(i), "RXN_RATE_CONSTANT(i), first-order rate constant for destruction of chemical i, 1/s")
    Call WriteFortranInput(f, MI.RXN_PRODUCT(i), "RXN_PRODUCT(i), index of chemical that is produced through destruction of chemical i (or 0 if none), unitless")
    Call WriteFortranInput(f, MI.RXN_RATIO(i), "RXN_RATIO(i), number of moles of chemical RXN_PRODUCT(i) produced by destruction of 1 mole of chemical i")
  Next i
  Call WriteFortranInput(f, MI.FN_MASSBAL_OUT, "FN_MASSBAL_OUT, filename of mass balance output file")
  Call WriteFortranInput(f, MI.FN_CR_OUT, "FN_CR_OUT, filename of room concentration vs time output file")
  Call WriteFortranInput(f, MI.FN_CB_OUT, "FN_CB_OUT, filename of bed effluent concentration vs time output file")
  Print #f, String$(75, "-")
  For i = 1 To MI.NUMB
    Call WriteFortranInput(f, "NO_DATA", "NAME(" & Trim$(Str$(i)) & ",1), not actually input")
    Call WriteFortranInput(f, MI.CHEMICALS(i, 1), "CHEMICALS(" & Trim$(Str$(i)) & ",1), molecular weight, g/gmol")
    Call WriteFortranInput(f, MI.CHEMICALS(i, 2), "CHEMICALS(" & Trim$(Str$(i)) & ",2), influent concentration, ug/L")
    Call WriteFortranInput(f, MI.CHEMICALS(i, 3), "CHEMICALS(" & Trim$(Str$(i)) & ",3), molar volume, cm^3/gmol")
    Call WriteFortranInput(f, MI.CHEMICALS(i, 4), "CHEMICALS(" & Trim$(Str$(i)) & ",4), Freundlich K, (umol/g)*(L/umol)^(1/n)")
    Call WriteFortranInput(f, MI.CHEMICALS(i, 5), "CHEMICALS(" & Trim$(Str$(i)) & ",5), Freundlich 1/n, dimless")
    Call WriteFortranInput(f, MI.CHEMICALS(i, 6), "CHEMICALS(" & Trim$(Str$(i)) & ",6), film transfer coefficient (kf), cm/s ")
    Call WriteFortranInput(f, MI.CHEMICALS(i, 7), "CHEMICALS(" & Trim$(Str$(i)) & ",7), surface diffusion coefficient (Ds), cm^2/s")
    Call WriteFortranInput(f, MI.CHEMICALS(i, 8), "CHEMICALS(" & Trim$(Str$(i)) & ",8), pore diffusion coefficient (Dp), cm^2/s")
    Call WriteFortranInput(f, MI.CHEMICALS(i, 9), "CHEMICALS(" & Trim$(Str$(i)) & ",9) = RK1(" & Trim$(Str$(i)) & "), fouling correlation coef. #1, dimless")
    Call WriteFortranInput(f, MI.CHEMICALS(i, 10), "CHEMICALS(" & Trim$(Str$(i)) & ",10) = RK2(" & Trim$(Str$(i)) & "), fouling correlation coef. #2, 1/min")
    Call WriteFortranInput(f, MI.CHEMICALS(i, 11), "CHEMICALS(" & Trim$(Str$(i)) & ",11) = RK3(" & Trim$(Str$(i)) & "), fouling correlation coef. #3, dimless")
    Call WriteFortranInput(f, MI.CHEMICALS(i, 12), "CHEMICALS(" & Trim$(Str$(i)) & ",12) = RK4(" & Trim$(Str$(i)) & "), fouling correlation coef. #4, 1/min")
    Call WriteFortranInput(f, MI.CHEMICALS(i, 13), "CHEMICALS(" & Trim$(Str$(i)) & ",13) = TORTU(" & Trim$(Str$(i)) & "), tortuosity (never used?), dimless")
    Call WriteFortranInput(f, MI.CHEMICALS(i, 14), "CHEMICALS(" & Trim$(Str$(i)) & ",14) = TOR(" & Trim$(Str$(i)) & "), tortuosity coef., dimless")
    Call WriteFortranInput(f, MI.CHEMICALS(i, 15), "CHEMICALS(" & Trim$(Str$(i)) & ",15) = PART(" & Trim$(Str$(i)) & "), part. coef., dimless")
    Call WriteFortranInput(f, MI.CHEMICALS(i, 16), "CHEMICALS(" & Trim$(Str$(i)) & ",16) = TTORTU(" & Trim$(Str$(i)) & "), time parameter, min")
    Call WriteFortranInput(f, MI.INITIAL_ROOM_CONC(i), "INITIAL_ROOM_CONC(" & Trim$(Str$(i)) & "), initial concentration in room, ug/L")
  Next i
  '
  ' TIME-VARIABLE INFLUENT CONCENTRATIONS.
  '
  If (MI.NINI <> 0) Then
    Print #f, String$(75, "-")
    Print #f, "NOTE1A: THE FOLLOWING LINES CONTAIN INFLUENT CONCENTRATION VS TIME DATA."
    Print #f, "NOTE1B: THE ORDER FOR EACH LINE IS: TIME (MINUTES), CONC#1 (UG/L), ..., CONC#n (UG/L)."
    For i = 1 To MI.NINI
      ThisLine = Trim$(Str$(MI.TINI(i)))
      For J = 1 To MI.NUMB
        ThisLine = ThisLine & "    "
        ThisLine = ThisLine & Trim$(Str$(ModelPSDMInRoom_Inputs_CINI(J, i)))
      Next J
      Print #f, ThisLine
    Next i
  End If
  '
  ' TIME-VARIABLE Co.
  '
  With RoomParams
    For i_ = 1 To MI.NUMB
      Print #f, String$(75, "-")
      i = Component_Index_PFPSDM(i_)
      intThis = IIf(.bool_ROOM_COINI_ISTIMEVAR(i) = True, 1, 0)
      Call WriteFortranInput(f, intThis, "bool_ROOM_COINI_ISTIMEVAR(" & Trim$(Str$(i)) & "), whether there are time-variable Co points for chemical #" & Trim$(Str$(i)) & ", dimless")
      If (intThis = 1) Then
        intThis = .int_ROOM_NCOINI(i)
        Call WriteFortranInput(f, intThis, "int_ROOM_NCOINI(" & Trim$(Str$(i)) & "), number of time-variable Co points for chemical #" & Trim$(Str$(i)) & ", dimless")
        Print #f, "dbl_ROOM_TCOINI(" & Trim$(Str$(i)) & ",j), time profile for dbl_ROOM_COINI() array, minutes"
        Print #f, "dbl_ROOM_COINI(" & Trim$(Str$(i)) & ",j), the Co values for these times, ug/L"
        For J = 1 To .int_ROOM_NCOINI(i)
          strTemp1 = " " & Trim$(Str$(.dbl_ROOM_TCOINI(i, J)))
          strTemp1 = strTemp1 & "    "
          strTemp1 = strTemp1 & Trim$(Str$(.dbl_ROOM_COINI(i, J)))
          Print #f, strTemp1
        Next J
      End If
    Next i_
  End With
  '
  ' TIME-VARIABLE w*A.
  '
  With RoomParams
    For i_ = 1 To MI.NUMB
      Print #f, String$(75, "-")
      i = Component_Index_PFPSDM(i_)
      intThis = IIf(.bool_ROOM_EMITINI_ISTIMEVAR(i) = True, 1, 0)
      Call WriteFortranInput(f, intThis, "bool_ROOM_EMITINI_ISTIMEVAR(" & Trim$(Str$(i)) & "), whether there are time-variable w*A points for chemical #" & Trim$(Str$(i)) & ", dimless")
      If (intThis = 1) Then
        intThis = .int_ROOM_NEMITINI(i)
        Call WriteFortranInput(f, intThis, "int_ROOM_NEMITINI(" & Trim$(Str$(i)) & "), number of time-variable w*A points for chemical #" & Trim$(Str$(i)) & ", dimless")
        Print #f, "dbl_ROOM_TEMITINI(" & Trim$(Str$(i)) & ",j), time profile for dbl_ROOM_EMITINI() array, minutes"
        Print #f, "dbl_ROOM_EMITINI(" & Trim$(Str$(i)) & ",j), the w*A values for these times, ug/s"
        For J = 1 To .int_ROOM_NEMITINI(i)
          strTemp1 = " " & Trim$(Str$(.dbl_ROOM_TEMITINI(i, J)))
          strTemp1 = strTemp1 & "    "
          strTemp1 = strTemp1 & Trim$(Str$(.dbl_ROOM_EMITINI(i, J)))
          Print #f, strTemp1
        Next J
      End If
    Next i_
  End With
  '
  ' TIME-VARIABLE K.
  '
Dim dbl_K_Conversion_Factor As Double
  With RoomParams
    For i_ = 1 To MI.NUMB
      Print #f, String$(75, "-")
      i = Component_Index_PFPSDM(i_)
      intThis = IIf(.bool_ROOM_KINI_ISTIMEVAR(i) = True, 1, 0)
      Call WriteFortranInput(f, intThis, "bool_ROOM_KINI_ISTIMEVAR(" & Trim$(Str$(i)) & "), whether there are time-variable K points for chemical #" & Trim$(Str$(i)) & ", dimless")
      If (intThis = 1) Then
        intThis = .int_ROOM_NKINI(i)
        Call WriteFortranInput(f, intThis, "int_ROOM_NKINI(" & Trim$(Str$(i)) & "), number of time-variable K points for chemical #" & Trim$(Str$(i)) & ", dimless")
        Print #f, "dbl_ROOM_TKINI(" & Trim$(Str$(i)) & ",j), time profile for dbl_ROOM_KINI() array, minutes"
        Print #f, "dbl_ROOM_KINI(" & Trim$(Str$(i)) & ",j), the K values for these times, (umol/g)*(L/umol)^(1/n)"
        '
        ' CONVERT K FROM (mg/g)*(L/mg)^(1/n) to (umol/g)*(L/umol)^(1/n).
        dbl_K_Conversion_Factor = 1# * (1000# / Component(i).MW) ^ (1# - Component(i).Use_OneOverN)
        For J = 1 To .int_ROOM_NKINI(i)
          strTemp1 = " " & Trim$(Str$(.dbl_ROOM_TKINI(i, J)))
          strTemp1 = strTemp1 & "    "
          strTemp1 = strTemp1 & Trim$(Str$(.dbl_ROOM_KINI(i, J) * dbl_K_Conversion_Factor))
          Print #f, strTemp1
        Next J
      End If
    Next i_
  End With
  
'  Print #f, "TINI(i), time profile for CINI() array, minutes"
'  For i = 1 To MI.NINI
'    Print #f, MI.TINI(i)
'  Next i
'  Print #f, "CINI(i,j), influent concentration profile, ug/L"
'  For i = 1 To MI.NUMB
'    For j = 1 To MI.NINI
'      Print #f, ModelPSDMInRoom_Inputs_CINI(i, j)
'    Next j
'  Next i
  
'  MI.NINI = Number_Influent_Points
'  For j = 1 To MI.NINI
'    MI.TINI(j) = T_Influent(j)
'    For i = 1 To MI.NUMB
'      'CONVERT FROM mg/L TO ug/L.
'      ModelPSDMInRoom_Inputs_CINI(i, j) = C_Influent(Component_Index_PFPSDM(i), j) * 1000#
'    Next i
'  Next j
  
  Print #f, String$(75, "-")
  Call WriteFortranInput(f, ModelPSDMInRoom_EofTestMarker, "EOFTESTMARKER")
  Close #f
  'STORE FOR LATER USE.
  ModelPSDMInRoom_Inputs = MI
End Sub
Sub ModelPSDMInRoom_WritePathFile()
Dim f As Integer
Dim fn_This As String
Dim qq As String
  qq = Chr$(34)
  f = FreeFile
  fn_This = Exe_Path & "\" & ModelPSDMInRoom_IN_PathFile
  Open fn_This For Output As #f
  Print #f, "1"
  Print #f, qq & ModelPSDMInRoom_IN_Main & qq
  Print #f, qq & ModelPSDMInRoom_OUT_SuccessFlag & qq
  Print #f, qq & ModelPSDMInRoom_OUT_Main & qq
  Print #f, qq & ModelPSDMInRoom_OUT_CRvsT & qq
  Print #f, qq & ModelPSDMInRoom_OUT_CBvsT & qq
  Close #f
End Sub


'Return value:
'  TRUE = Okay to call the PSDM
'  FALSE = Something went wrong, ABORT!  ABORT!
Function Prepare_To_Run_PSDM_In_Room() As Integer
Dim i As Integer
Dim J As Integer
Dim Num_K_Reduction As Integer
Dim Num_A_and_Not_B As Integer
Dim Num_Not_a_and_B As Integer
ReDim Name_A_and_Not_B(1 To Number_Compo_Max) As String
ReDim Name_Not_A_and_B(1 To Number_Compo_Max) As String
Dim Is_A As Integer
Dim Is_B As Integer
Dim msg As String
Dim RetVal As Integer
  '
  ' PERFORM SEVERAL VERIFICATIONS BEFORE RUNNING THE PSDM.
  '
  If (TimeP.Init > TimeP.End) Then
    Call Show_Error("The initial simulation time (" & _
        TimeP.Init / 24# / 60# & " days) is greater than the " & _
        "final simulation time (" & TimeP.End / 24# / 60# & _
        " days).  PSDM cannot be run until this is fixed.")
    Prepare_To_Run_PSDM_In_Room = False
    Exit Function
  End If
  If (TimeP.Step < ((TimeP.End - TimeP.Init) / _
      (Number_Points_Max - 1))) Then
    Call Show_Error("The simulation time step (" & _
        TimeP.Step / 24# / 60# & " days) is too small.  The " & _
        "maximum number of points is 400.  PSDM cannot be run " & _
        "until this is fixed.")
    Prepare_To_Run_PSDM_In_Room = False
    Exit Function
  End If
  Call AllModels_Verify_Selected_Components(MODELTYPE_PSDM)
  If (Number_Component_PFPSDM = 0) Then
    Prepare_To_Run_PSDM_In_Room = False
    Exit Function
  End If
  For i = 1 To Number_Component_PFPSDM
    For J = i + 1 To Number_Component_PFPSDM
      If Trim$(Component(Component_Index_PFPSDM(i)).Name) = _
          Trim$(Component(Component_Index_PFPSDM(J)).Name) Then
        Call Show_Error("Components " & _
            Format$(Component_Index_PFPSDM(i), "0") & " and " & _
            Format$(Component_Index_PFPSDM(J), "0") & _
            " have the same name." & vbCrLf & _
            "Please change one before running the PSDM.")
        Prepare_To_Run_PSDM_In_Room = False
        Exit Function
      End If
    Next J
  Next i
  '
  '---- Make sure # PSDM fouling components is <= 1.
  '
  Num_K_Reduction = 0
  For i = 0 To frmMain.lstComponents.ListCount - 1
    If (frmMain.lstComponents.Selected(i)) Then
      If (Component(i + 1).K_Reduction) Then
        Num_K_Reduction = Num_K_Reduction + 1
      End If
    End If
  Next i
  If (Num_K_Reduction > 1) Then
    Call Show_Error("There are currently " & _
        Trim$(Str$(Num_K_Reduction)) & _
        " components specified for fouling.  Only 1 may be " & _
        "specified for a run of the PSDM.")
    Prepare_To_Run_PSDM_In_Room = False
    Exit Function
  End If
  '
  '---- Show warning if A and not B, or not A and B,
  '.... for any component where:
  '.... A = pore diffusion correlation for tortuosity selected
  '.... B = fouling correlation selected
  '
  Num_A_and_Not_B = 0
  Num_Not_a_and_B = 0
  For i = 0 To frmMain.lstComponents.ListCount - 1
    If (frmMain.lstComponents.Selected(i)) Then
      Is_A = (Component(i + 1).Use_Tortuosity_Correlation)
      Is_B = (Component(i + 1).K_Reduction)
      '---- Check for A and not B case:
      If ((Is_A) And (Not Is_B)) Then
        Num_A_and_Not_B = Num_A_and_Not_B + 1
        Name_A_and_Not_B(Num_A_and_Not_B) = Trim$(Component(i + 1).Name)
      End If
      '---- Check for not A and B case:
      If ((Not Is_A) And (Is_B)) Then
        Num_Not_a_and_B = Num_Not_a_and_B + 1
        Name_Not_A_and_B(Num_Not_a_and_B) = Trim$(Component(i + 1).Name)
      End If
    End If
  Next i
  If ((Num_A_and_Not_B > 0) Or (Num_Not_a_and_B > 0)) Then
    msg = "Warning:" & vbCrLf
    If (Num_A_and_Not_B > 0) Then
      msg = msg & vbCrLf
      msg = msg & "The following components have the pore diffusion "
      msg = msg & "correlation for tortuosity selected, but no "
      msg = msg & "fouling correlation selected:"
      msg = msg & vbCrLf
      For i = 1 To Num_A_and_Not_B
        msg = msg & "    " & Name_A_and_Not_B(i)
        msg = msg & vbCrLf
      Next i
    End If
    If (Num_Not_a_and_B > 0) Then
      msg = msg & vbCrLf
      msg = msg & "The following components have the pore diffusion "
      msg = msg & "correlation for tortuosity NOT selected, but a "
      msg = msg & "fouling correlation is selected:"
      msg = msg & vbCrLf
      For i = 1 To Num_Not_a_and_B
        msg = msg & "    " & Name_Not_A_and_B(i)
        msg = msg & vbCrLf
      Next i
    End If
    msg = msg & vbCrLf
    msg = msg & "This configuration is not the recommended way to run "
    msg = msg & "the PSDM.  It is recommended that you either (a) "
    msg = msg & "turn both correlations on or (b) "
    msg = msg & "turn both correlations off.  Do you wish to proceed "
    msg = msg & "with this currently-specified PSDM run anyway?"
    RetVal = MsgBox(msg, vbQuestion + vbYesNo, AppName_For_Display_Long)
    If (RetVal = vbNo) Then
      Prepare_To_Run_PSDM_In_Room = False
      Exit Function
    End If
  End If
  '
  ' CANCEL SIM IF NAE>1.
  '
  If (Bed.NumberOfBeds > 1) Then
    Call Show_Error("You currently have specified " & _
        Trim$(Str$(Bed.NumberOfBeds)) & _
        " axial elements.  The PSDMR model currently supports only " & _
        "one axial element.  Recommendation: Reset the number of axial " & _
        "elements to one (1) and retry this calculation.")
    Prepare_To_Run_PSDM_In_Room = False
    Exit Function
  End If
  Prepare_To_Run_PSDM_In_Room = True
End Function


Sub Parser2_GetArg(sepchar As String, inline As String, ArgNum As Integer, retStr As String)
Dim i As Integer
Dim J As Integer
  retStr = ""
  J = 1
  For i = 1 To Len(inline)
    If (Mid$(inline, i, 1) = sepchar) Then
      J = J + 1
      If (J > ArgNum) Then Exit For
    Else
      If (J = ArgNum) Then
        retStr = retStr + Mid$(inline, i, 1)
      End If
    End If
  Next i
End Sub
Function Parser2_GetNumArgs(sepchar As String, inline As String) As Integer
Dim NumArgs As Integer
Dim i As Integer
  NumArgs = 1     'between chr #1 and first separator char.
  For i = 1 To Len(inline)
    If (Mid$(inline, i, 1) = sepchar) Then
      NumArgs = NumArgs + 1
    End If
  Next i
  Parser2_GetNumArgs = NumArgs
End Function
Function Parser2_RemoveDuplicateSeparators(sepchar As String, inline As String) As String
Dim retStr As String
Dim i As Integer
Dim ok_append As Integer
Dim thisc As String
  retStr = ""
  For i = 1 To Len(inline)
    ok_append = True
    thisc = Mid$(inline, i, 1)
    If (i > 1) Then
      If (thisc = sepchar) Then
        If (Right$(retStr, 1) = sepchar) Then
          ok_append = False
        End If
      End If
    End If
    If (ok_append) Then
      retStr = retStr & thisc
    End If
  Next i
  Parser2_RemoveDuplicateSeparators = retStr
End Function



