Attribute VB_Name = "ModelCPHSDM"
Option Explicit

Const ModelCPHSDM_IN_PathFile = "CPHSDM1.IN"
Const ModelCPHSDM_IN_Main = "CPHSDM2.IN"
Const ModelCPHSDM_OUT_SuccessFlag = "CPHSDM1.OUT"
Const ModelCPHSDM_OUT_Main = "CPHSDM2.OUT"

Const ModelCPHSDM_Version = 1#
Const ModelCPHSDM_ExeName = "CPHSDM2.EXE"
Const ModelCPHSDM_EofTestMarker = 123456#

'Const ModelCPHSDM_NMAX = 1
Private Type ModelCPHSDM_Inputs_Type
  Bed(1 To 6) As Double
  Compo(1 To 4) As Double
  Kine(1 To 2) As Double
End Type
Dim ModelCPHSDM_Inputs As ModelCPHSDM_Inputs_Type
Private Type ModelCPHSDM_Outputs_Type
  TACT(1 To 210) As Double
  CC(1 To 210) As Double
  PARAM(1 To 7) As Double
  ER_FLAG As Integer
End Type
Dim ModelCPHSDM_Outputs As ModelCPHSDM_Outputs_Type




Const ModelCPHSDM_declarations_end = True


Sub ModelCPHSDM_Go()
  Call ModelCPHSDM_WritePathFile
  Call ModelCPHSDM_WriteMainFile
  Call ModelCPHSDM_CallEXE
  Call ModelCPHSDM_ProcessOutput
  If (ModelIO_IsKeepTempFiles() = False) Then
    Call ModelCPHSDM_RemoveLinkFiles
  End If
End Sub


Sub ModelCPHSDM_RemoveLinkFiles()
  Call KillFile_If_Exists(Exe_Path & "\" & ModelCPHSDM_IN_PathFile)
  Call KillFile_If_Exists(Exe_Path & "\" & ModelCPHSDM_IN_Main)
  Call KillFile_If_Exists(Exe_Path & "\" & ModelCPHSDM_OUT_SuccessFlag)
  Call KillFile_If_Exists(Exe_Path & "\" & ModelCPHSDM_OUT_Main)
End Sub
Sub ModelCPHSDM_CallEXE()
Dim CmdLine As String
  Call ChangeDir_Exes
  CmdLine = ModelCPHSDM_ExeName
  Call ModelIO_Timer_Start
  Call FortranLink_ExecAndWaitForProcess(CmdLine)
  Call ModelIO_Timer_End
  Call ChangeDir_Main
End Sub
Sub ModelCPHSDM_ProcessOutput()
Dim f As Integer
Dim fn_This As String
Dim ER_FLAG As Integer
Dim DummyStr1 As String
Dim temp As String
Dim i As Integer
Dim J As Integer
Dim MO As ModelCPHSDM_Outputs_Type
Dim EOFTESTMARKER As Double
Dim Flag05 As Boolean
Dim Flag50 As Boolean
Dim Flag95 As Boolean
  'READ SUCCESS FLAG OUTPUT FILE.
  f = FreeFile
  fn_This = Exe_Path & "\" & ModelCPHSDM_OUT_SuccessFlag
  If (Not FileExists(fn_This)) Then
    Call Show_Error("Unable to find output file: Calculations failed.")
    Exit Sub
  End If
  Open fn_This For Input As #f
  Line Input #f, DummyStr1
  Input #f, ER_FLAG
  Close #f
  If (ER_FLAG <> 0) Then
    Select Case ER_FLAG
      Case 40
        temp = "The value of 1/n is out of range for St minimum."
      Case 41
        temp = "The value of 1/n is out of range for St minimum."
      Case 42
        temp = "The value of the Biot number is out of range."
      Case 44
        temp = "The value of 1/n is out of range."
      Case Else
        temp = "Unknown Error."
    End Select
    Call Show_Error("The CPHSDM failed to converge." & vbCrLf & temp)
    Exit Sub
  Else
    Call Show_Message( _
        "CPHSDM Model Calculations Complete." & _
        vbCrLf & _
        vbCrLf & _
        ModelIO_Timer_SummaryMsg)
  End If
  'READ MAIN OUTPUT FILE.
  fn_This = Exe_Path & "\" & ModelCPHSDM_OUT_Main
  Open fn_This For Input As #f
  Line Input #f, DummyStr1
  For i = 1 To 210
    Input #f, MO.TACT(i)
  Next i
  Line Input #f, DummyStr1
  For i = 1 To 210
    Input #f, MO.CC(i)
  Next i
  Line Input #f, DummyStr1
  For i = 1 To 7
    Input #f, MO.PARAM(i)
  Next i
  Line Input #f, DummyStr1
  Input #f, EOFTESTMARKER
  If (False = ModelIO_DoNumberCheck(EOFTESTMARKER, ModelCPHSDM_EofTestMarker)) Then
    Call Show_Error("The model calculations failed: invalid file format (EOF marker).")
    Exit Sub
  End If
  Close #f
  ModelCPHSDM_Outputs = MO
  'TRANSFER OUTPUT DATA TO MORE PERMANENT MEMORY.
  For i = 1 To CPM_Max_Points
    CPM_Results.T(i) = MO.TACT(i)           'TACT(I) is in days
    CPM_Results.C_Over_C0(i) = MO.CC(i)     'CC(I) is dimensionless
  Next i
  For i = 1 To 7
    CPM_Results.Par(i) = MO.PARAM(i)
  Next i
  ' Description of CPM_Results.Par
  ' 1 -> Minimum Stanton number
  ' 2 -> Minimum EBCT (min)
  ' 3 -> Minimum Length (cm)
  ' 4 -> Throughput Ratio at 95%
  ' 5 -> Throughput Ratio at 5%
  ' 6 -> EBCT of MTZ (min)
  ' 7 -> Length of MTZ (cm)
  CPM_Results.Bed = Bed
  CPM_Results.Carbon = Carbon
  CPM_Results.Component = Component(Component_Index_CPM)
  ''''CPM_Results.Constant_Tortuosity = Constant_Tortuosity
  ''''CPM_Results.Use_Tortuosity_Correlation = Use_Tortuosity_Correlation
  Flag05 = True
  Flag50 = True
  Flag95 = True
  For J = 1 To CPM_Max_Points
    If (J > 2) Then
      If (MO.CC(J) >= 0.05) And (MO.CC(J - 1) < 0.05) And Flag05 Then
        CPM_Results.ThroughPut_05.T = (MO.TACT(J) - MO.TACT(J - 1)) / _
            (MO.CC(J) - MO.CC(J - 1)) * (0.05 - MO.CC(J - 1)) + MO.TACT(J - 1)
        CPM_Results.ThroughPut_05.C = ((MO.CC(J) - MO.CC(J - 1)) / _
            (MO.TACT(J) - MO.TACT(J - 1)) * _
            (CPM_Results.ThroughPut_05.T - MO.TACT(J - 1)) + MO.CC(J - 1)) * _
            CPM_Results.Component.InitialConcentration
        Flag05 = False
      End If
      If (MO.CC(J) >= 0.5) And (MO.CC(J - 1) < 0.5) And Flag50 Then
        CPM_Results.ThroughPut_50.T = (MO.TACT(J) - MO.TACT(J - 1)) / _
            (MO.CC(J) - MO.CC(J - 1)) * (0.5 - MO.CC(J - 1)) + MO.TACT(J - 1)
        CPM_Results.ThroughPut_50.C = ((MO.CC(J) - MO.CC(J - 1)) / _
            (MO.TACT(J) - MO.TACT(J - 1)) * _
            (CPM_Results.ThroughPut_50.T - MO.TACT(J - 1)) + MO.CC(J - 1)) * _
            CPM_Results.Component.InitialConcentration
        Flag50 = False
      End If
      If (MO.CC(J) >= 0.95) And (MO.CC(J - 1) < 0.95) And Flag95 Then
        CPM_Results.ThroughPut_95.T = (MO.TACT(J) - MO.TACT(J - 1)) / _
            (MO.CC(J) - MO.CC(J - 1)) * (0.95 - MO.CC(J - 1)) + MO.TACT(J - 1)
        CPM_Results.ThroughPut_95.C = ((MO.CC(J) - MO.CC(J - 1)) / _
            (MO.TACT(J) - MO.TACT(J - 1)) * _
            (CPM_Results.ThroughPut_95.T - MO.TACT(J - 1)) + MO.CC(J - 1)) * _
            CPM_Results.Component.InitialConcentration
        Flag95 = False
      End If
    End If
  Next J
  'ENABLE RESULTS MENU COMMANDS.
  frmMain.mnuResultsItem(1).Enabled = True
  If (NData_Points > 0) Then
    frmMain.mnuResultsItem(4).Enabled = True
  End If
End Sub
Sub ModelCPHSDM_WriteMainFile()
Dim f As Integer
Dim fn_This As String
Dim MI As ModelCPHSDM_Inputs_Type
Dim i As Integer
Dim J As Integer
Dim A1 As Double
Dim A2 As Double
Dim A3 As Double
Dim A4 As Double
  'PREPARE INPUTS.
  J = Component_Index_CPM
  '
  '------ INPUT SET #1: BED PROPERTIES. ------
  '
  'PARTICLE DIAMETER (cm).
  MI.Bed(1) = Carbon.ParticleRadius * 200#
  'BED DENSITY (g/cm^3).
  MI.Bed(2) = Bed.Weight * 4# / Bed.Length / Bed.Diameter ^ 2# / PI / 1000#
  'APPARENT PARTICLE DENSITY (g/cm^3).
  MI.Bed(3) = Carbon.Density
  'EBCT (minutes).
  MI.Bed(4) = Bed.Length * PI * Bed.Diameter * Bed.Diameter / 4# / Bed.Flowrate / 60#
  'SUPERFICIAL VELOCITY (cm/s).
  MI.Bed(5) = Bed.Flowrate * 4# / PI / Bed.Diameter ^ 2# * 100#
  'PARTICLE POROSITY (dimensionless).
  MI.Bed(6) = Carbon.Porosity
  '
  '------ INPUT SET #2: COMPONENT PROPERTIES. ------
  '
  'MOLECULAR WEIGHT (g/gmol).
  MI.Compo(1) = Component(J).MW
  'INFLUENT CONCENTRATION (ug/L).
  MI.Compo(2) = Component(J).InitialConcentration * 1000#
  'FREUNDLICH K (umol/g)*(L/umol)^(1/n).
  MI.Compo(3) = Component(J).Use_K * (1000# / Component(J).MW) ^ (1 - Component(J).Use_OneOverN)
  'FREUNDLICH 1/n (dimensionless).
  MI.Compo(4) = Component(J).Use_OneOverN
  '
  '------ INPUT SET #3: KINETIC PARAMETERS. ------
  '
  'FILM TRANSFER COEFFICIENT (cm/s).
  MI.Kine(1) = Component(J).kf
  'SURFACE DIFFUSION COEFFICIENT (cm^2/s).
  MI.Kine(2) = Component(J).Ds
  '
  '------ CALCULATE K REDUCTION DUE TO FOULING. ------
  '
  A1 = Bed.Water_Correlation.Coeff(1) * Component(J).Correlation.Coeff(1) + Component(J).Correlation.Coeff(2)
  A2 = Bed.Water_Correlation.Coeff(2) * Component(J).Correlation.Coeff(1)
  A3 = Bed.Water_Correlation.Coeff(3) * Component(J).Correlation.Coeff(1)
  A4 = Bed.Water_Correlation.Coeff(4) * Component(J).Correlation.Coeff(1)
  If (Bed.Phase = 0) Then
    If (Component(J).K_Reduction) And (Bed.Water_Correlation.Coeff(1) <> 1# And Bed.Water_Correlation.Coeff(2) <> 0# And Bed.Water_Correlation.Coeff(3) <> 0# And Bed.Water_Correlation.Coeff(4) <> 0#) Then
      Dim DG1 As Double
      Dim DG2 As Double
      Dim KovK0 As Double
      Dim DG As Double
      Dim TAU As Double
      Dim T_Minut As Double
      i = 0
      DG2 = 1#
      KovK0 = 1#
ModelCPHSDM_WriteMainFile_NextIteration:
      If (i < Max_Number_Fouling_Iterations) And (Abs(1# - DG1 / DG2) > 0.01) Then
        DG1 = 1000# * MI.Bed(2) * (MI.Bed(2) / MI.Bed(3)) / _
            (1 - MI.Bed(2) / MI.Bed(3)) / MI.Compo(2) * KovK0 * MI.Compo(3) * MI.Compo(2) ^ MI.Compo(4)
        'TAU = EBST * epsilon
        TAU = MI.Bed(4) * (1 - MI.Bed(2) / MI.Bed(3))
        T_Minut = TAU * (DG1 + 1)
        KovK0 = A1 + A2 * T_Minut + A3 * Exp(A4 * T_Minut)
        DG2 = 1000# * MI.Bed(2) * (MI.Bed(2) / MI.Bed(3)) / _
            (1 - MI.Bed(2) / MI.Bed(3)) / MI.Compo(2) * KovK0 * MI.Compo(3) * MI.Compo(2) ^ MI.Compo(4)
        i = i + 1
        GoTo ModelCPHSDM_WriteMainFile_NextIteration
      End If
      If i < Max_Number_Fouling_Iterations Then
        MI.Compo(3) = MI.Compo(3) * KovK0
      Else
        Call Show_Error( _
            "The iterations to evaluate the capacity reduction " & _
            "due to fouling did not converge." & vbCrLf & _
            "It will be assumed there is no fouling.")
      End If
    End If
  End If
  'WRITE INPUT FILE.
  f = FreeFile
  fn_This = Exe_Path & "\" & ModelCPHSDM_IN_Main
  'fn_This = App.Path & "\" & ModelECM_IN_Main
  Open fn_This For Output As #f
  Call WriteFortranInput(f, ModelCPHSDM_Version, "MODULE_VERSION")
  Call WriteFortranInput(f, MI.Bed(1), "Bed(1), particle diameter, cm")
  Call WriteFortranInput(f, MI.Bed(2), "Bed(2), bed density, g/cm^3")
  Call WriteFortranInput(f, MI.Bed(3), "Bed(3), apparent particle density, g/cm^3")
  Call WriteFortranInput(f, MI.Bed(4), "Bed(4), empty bed contact time (EBCT), minutes")
  Call WriteFortranInput(f, MI.Bed(5), "Bed(5), superficial velocity, cm/s")
  Call WriteFortranInput(f, MI.Bed(6), "Bed(6), particle porosity, dimless")
  Call WriteFortranInput(f, MI.Compo(1), "Compo(1), molecular weight, g/gmol")
  Call WriteFortranInput(f, MI.Compo(2), "Compo(2), influent concentration, ug/L")
  Call WriteFortranInput(f, MI.Compo(3), "Compo(3), Freundlich K, (umol/g)*(L/umol)^(1/n)")
  Call WriteFortranInput(f, MI.Compo(4), "Compo(4), Freundlich 1/n, dimless")
  Call WriteFortranInput(f, MI.Kine(1), "Kine(1), film transfer coefficient, cm/s")
  Call WriteFortranInput(f, MI.Kine(2), "Kine(2), surface diffusion coefficient, cm^2/s")
  Call WriteFortranInput(f, ModelCPHSDM_EofTestMarker, "EOFTESTMARKER")
  Close #f
  'STORE FOR LATER USE.
  ModelCPHSDM_Inputs = MI
End Sub
Sub ModelCPHSDM_WritePathFile()
Dim f As Integer
Dim fn_This As String
Dim qq As String
  qq = Chr$(34)
  f = FreeFile
  fn_This = Exe_Path & "\" & ModelCPHSDM_IN_PathFile
  Open fn_This For Output As #f
  Print #f, qq & ModelCPHSDM_IN_Main & qq
  Print #f, qq & ModelCPHSDM_OUT_SuccessFlag & qq
  Print #f, qq & ModelCPHSDM_OUT_Main & qq
  Close #f
End Sub


