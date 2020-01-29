Option Explicit
Option Base 1

'Declare DLL's for water density and water viscosity
'DLLs for Air & Water Properties
    'H2oDens Returns D as kg/m3 and Input T as K
    Declare Sub H2ODens Lib "H2oDens.DLL" (D As Double, T As Double)

    'H2oVisc Returns V as kg/m/sec and Input T as K
    Declare Sub H2OVisc Lib "H2oVisc.DLL" (V As Double, T As Double)

'DLL for PFPDM - Version without variable influent or beds in series
'Declare Sub PFPDM04 Lib "pdm04dll.dll" (Num As Integer, Chem As Double, Ads As Double, casrb As Double, T As Double, CP As Double, ITP As Integer, TT As Double, N As Integer, m As Integer, Nin As Integer, Tin As Double, CIN As Double, Size As Long, FLAG As Integer)

'DLL for PFPDM - Version with variable influent or beds in series
'Declare Sub PFPDM05 Lib "pdm05dll.dll" (Num As Integer, Chem As Double, Ads As Double, casrb As Double, T As Double, CP As Double, ITP As Integer, TT As Double, N As Integer, m As Integer, Nin As Integer, Tin As Double, Cin As Double, Size As Long, FLAG As Integer)

'DLL for PFPDM - Version with exiting dll when reached 1 % of influent concentration
'Declare Sub PFPDM06 Lib "pdm06dll.dll" (Num As Integer, Chem As Double, Ads As Double, casrb As Double, T As Double, CP As Double, ITP As Integer, TT As Double, N As Integer, m As Integer, Nin As Integer, Tin As Double, Cin As Double, Size As Long, FLAG As Integer)

'DLL for PFPDM - Version for setting CPORE to 0 if it is <= 0 and Calculating QH as QT - (Sum of Other Ions)
'Declare Sub PFPDM07 Lib "pdm07dll.dll" (Num As Integer, Chem As Double, Ads As Double, casrb As Double, T As Double, CP As Double, ITP As Integer, TT As Double, N As Integer, m As Integer, Nin As Integer, Tin As Double, Cin As Double, CT_Average As Double, Size As Long, FLAG As Integer)

'------Begin Modification Hokanson: 11-Aug2000
'DLL for PFPDM - Version for setting CPORE to 0 if it is <= 0 and Calculating QH as QT - (Sum of Other Ions)
'Declare Sub PFPDM08 Lib "pdm08dll.dll" (Num As Integer, Chem As Double, Ads As Double, casrb As Double, T As Double, CP As Double, ITP As Integer, TT As Double, N As Integer, m As Integer, Nin As Integer, Tin As Double, Cin As Double, CT_Average As Double, Size As Long, FLAG As Integer)

'------Begin Modification Hokanson: 12-Aug2000
'Declare Sub PFPDM09 Lib "pdm09dll.dll" (Num As Integer, Chem As Double, Ads As Double, casrb As Double, T As Double, CP As Double, ITP As Integer, TT As Double, N As Integer, m As Integer, Nin As Integer, Tin As Double, Cin As Double, CT_Average As Double, Size As Long, FLAG As Integer, EPS As Double, DH0 As Double)
'------End Modification Hokanson: 11-Aug2000

Declare Sub PFPDM10 Lib "pdm10dll.dll" (Num As Integer, Chem As Double, Ads As Double, casrb As Double, T As Double, CP As Double, ITP As Integer, TT As Double, N As Integer, m As Integer, Nin As Integer, Tin As Double, Cin As Double, CT_Average As Double, Size As Long, FLAG As Integer, EPS As Double, DH0 As Double)
'------End Modification Hokanson: 12-Aug2000

'Fit Data subroutines
Declare Sub OBJFUN Lib "OBJFUN.DLL" (NCOMP As Long, NDATA As Long, NP As Long, TP As Double, CP As Double, TD As Double, CD As Double, Cin As Double, FMIN As Double)

Sub Call_PFPDM ()
ReDim Chemicals(NumSelectedComponents_PFPDM, 7) As Double
ReDim C_Prop(8) As Double, Ads_Prop(4) As Double
ReDim CP(NumSelectedComponents_PFPDM, Number_Points_Max)  As Double, TT(5) As Double
ReDim T(Number_Points_Max, 2) As Double
ReDim TOUT(Number_Points_Max) As Double
ReDim CT_Average(Number_Points_Max) As Double
ReDim Cin(NumSelectedComponents_PFPDM, Number_Max_Influent_Points) As Double
ReDim Tin(Number_Max_Influent_Points) As Double
Dim i As Integer, NFLAG As Integer, First As Integer, Nin As Integer
Dim ITP As Integer, j As Integer, F As Integer, Filenamebis As String
Dim Flag05 As Integer, Flag95 As Integer, Flag50 As Integer
Dim temp As String, Error_Code As Integer, Workspace_Size As Long
Dim Number_Equations As Integer
Dim Error_Message As String, MC As Integer, NC As Integer
Dim BedCounter As Integer
Dim NumberInfluentPointsToBed1 As Integer, K As Integer
Dim SumTimeAverageInfluentConcs As Double   'Number used in dimensionless groups
Dim GotToOne As Integer, msg As String
ReDim BedLeavingTimes(0 To Bed.NumberOfBeds + 1) As String
Dim HadVarInfBefore As Integer, HadOutputFileBefore As Integer
'------Begin Modification Hokanson: 11-Aug2000
Dim EPS As Double, DH0 As Double
'------End Modification Hokanson: 11-Aug2000

    Screen.MousePointer = 11

    MC = NumAxialCollocationPoints
    NC = NumRadialCollocationPoints

'------Begin Modification Hokanson: 11-Aug2000
    EPS = EPS_ErrorCriteriaForDGEARIntegrator
    DH0 = DH0_InitialTimeStepForDGEARIntegrator
'------End Modification Hokanson: 11-Aug2000

    Number_Equations = NumSelectedComponents_PFPDM * (MC * (NC + 1) - 1)
    Workspace_Size = Number_Equations ^ 2 + 2 * Number_Equations
    
    Flag50 = True
    Flag95 = True
    Flag05 = True

    frmLoading.Enabled = True
    frmLoading.Label1 = "Performing calculations..."
    frmLoading!Label2.Visible = True
    frmLoading.Show
    frmLoading.Refresh

    NFLAG = 0
    C_Prop(1) = Resin.ParticlePorosity       '(-)
    C_Prop(2) = Resin.ApparentDensity        'g/cm3
    C_Prop(3) = Resin.ParticleRadius * 100#  'To convert in cm
    C_Prop(4) = Resin.Tortuosity             '(-)
    C_Prop(8) = Resin.TotalCapacity          'meq/g

'    Time-Variable Tortuosity info. not available yet so input
'    C_Prop(5), C_Prop(6), and C_Prop(7) as -1 since they won't
'    be used
     C_Prop(5) = -1#
     C_Prop(6) = -1#
     C_Prop(7) = -1#

'    Time Variable Tortuosity code from AdSS:
'    If (Constant_Tortuosity) And (Use_Tortuosity_Correlation) Then
'       C_Prop(5) = 2#
'       C_Prop(6) = 0#
'    Else
'      If Use_Tortuosity_Correlation Then
'       C_Prop(5) = .334
'       C_Prop(6) = .00000661
'      Else
'       C_Prop(5) = 2#
'       C_Prop(6) = 0#
'      End If
'    End If
'    C_Prop(7) = 100000#  ' in mn

    Ads_Prop(1) = Bed.Length / CDbl(Bed.NumberOfBeds)     'm
    Ads_Prop(2) = Bed.Diameter                            'm
    Ads_Prop(3) = Bed.Weight / CDbl(Bed.NumberOfBeds)     'kg
    Ads_Prop(4) = Bed.Flowrate.Value                      'm3/s
    
    TT(1) = TimeParameters.FinalTime
    'Test value of Tinit
    If TimeParameters.InitialTime <= 0# Then
      TT(2) = .0001
    Else
      TT(2) = TimeParameters.InitialTime
    End If
    TT(3) = TimeParameters.TimeStep

' I is the index in the list of components selected for the PFPDM,
'   J is the one in the list of all components (frmIonExchangeMain!lstIons)
    SumTimeAverageInfluentConcs = 0#
    For i = 1 To NumSelectedComponents_PFPDM
       j = Component_Index_PFPDM(i)
       Chemicals(i, 1) = Ion(j).MolecularWeight          'mg/mmol
       Chemicals(i, 2) = Ion(j).InitialConcentration     'mg/L
       Chemicals(i, 3) = AlphaInput(j)                  'Separation Factor
       Chemicals(i, 4) = Ion(j).Kinetic.IonicTransportCoefficient.Value   'cm/s
       Chemicals(i, 5) = Ion(j).Kinetic.PoreDiffusivity.Value             'cm2/s
       Chemicals(i, 6) = Ion(j).Valence
       Chemicals(i, 7) = Resin.PresaturantPercentage(j) / 100#
       SumTimeAverageInfluentConcs = SumTimeAverageInfluentConcs + Chemicals(i, 2) * Chemicals(i, 6) / Chemicals(i, 1)
    Next i

    For j = 1 To Number_Influent_Points
      Tin(j) = T_Influent(j)
      For i = 1 To NumSelectedComponents_PFPDM
       Cin(i, j) = C_Influent(Component_Index_PFPDM(i), j)
      Next i
    Next j
    HadVarInfBefore = False
    HadOutputFileBefore = False
    NumberInfluentPointsToBed1 = Number_Influent_Points
    frmLoading!lblTime(0) = Format$(Now, "mm/dd/yy  h:nn:ss AM/PM")
    frmLoading!lblTotalBeds = Trim$(Str$(Bed.NumberOfBeds))

For BedCounter = 1 To Bed.NumberOfBeds

    Call WriteInfoToFile(1, Number_Influent_Points, HadVarInfBefore, BedCounter, Tin(), Cin())

On Error GoTo Error_DLL:
    frmLoading!Label2.Caption = "Currently Calculating Results for Bed Number " & Trim$(Str$(BedCounter))
    If BedCounter = 1 Then
       frmLoading!lblTime(1) = frmLoading!lblTime(0)
    Else
       frmLoading!lblTime(1) = Format$(Now, "mm/dd/yy  h:nn:ss AM/PM")
    End If
    BedLeavingTimes(BedCounter - 1) = frmLoading!lblTime(1)
    frmLoading.Refresh

'------Begin Modification Hokanson: 11-Aug2000
'    Call PFPDM08(NumSelectedComponents_PFPDM, Chemicals(1, 1), Ads_Prop(1), C_Prop(1), T(1, 1), CP(1, 1), ITP, TT(1), NC, MC, Number_Influent_Points, Tin(1), Cin(1, 1), CT_Average(1), Workspace_Size, NFLAG)
'------Begin Modification Hokanson: 12-Aug2000
'    Call PFPDM09(NumSelectedComponents_PFPDM, Chemicals(1, 1), Ads_Prop(1), C_Prop(1), T(1, 1), CP(1, 1), ITP, TT(1), NC, MC, Number_Influent_Points, Tin(1), Cin(1, 1), CT_Average(1), Workspace_Size, NFLAG, EPS, DH0)
'------End Modification Hokanson: 11-Aug2000
    Call PFPDM10(NumSelectedComponents_PFPDM, Chemicals(1, 1), Ads_Prop(1), C_Prop(1), T(1, 1), CP(1, 1), ITP, TT(1), NC, MC, Number_Influent_Points, Tin(1), Cin(1, 1), CT_Average(1), Workspace_Size, NFLAG, EPS, DH0)
'------End Modification Hokanson: 12-Aug2000

'---Check whether or not the model converged--------------------------
    Select Case NFLAG
    Case 1603
      MsgBox "The DLL for PFPDM could not be loaded into memory. Not enough memory.", MB_ICONEXCLAMATION, Application_Name
    Case 0
'      MsgBox "PFPDM completed.", 64, Application_Name
    Case Else
      Select Case NFLAG
        Case 15
           Error_Message = "WARNING..  T + H = T ON NEXT STEP"
        Case 105
          Error_Message = "KFLAG = -1 FROM INTEGRATOR"
        Case 115
          Error_Message = "H HAS BEEN REDUCED TO AND STEP WILL BE RETRIED"
        Case 155
          Error_Message = "PROBLEM APPEARS UNSOLVABLE WITH GIVEN INPUT"
        Case 205
          Error_Message = "THE REQUESTED ERROR IS SMALLER THAN CAN BE HANDLED"
        Case 255
          Error_Message = "INTEGRATION HALTED BY DRIVER EPS TOO SMALL TO BE ATTAINED FOR THE MACHINE PRECISION"
        Case 305
          Error_Message = "CORRECTOR CONVERGENCE COULD NOT BE ACHIEVED"
        Case 405
          Error_Message = "ILLEGAL INPUT.. EPS < 0"
        Case 415
          Error_Message = "ILLEGAL INPUT.. N <= 0"
        Case 425
          Error_Message = "ILLEGAL INPUT.. (T0-TOUT)*H >= 0 "
        Case 435
          Error_Message = "ILLEGAL INPUT.. INDEX"
        Case 445
          Error_Message = "INTERPOLATION WAS DONE AS  ON NORMAL RETURN.DESIRED PARAMETER CHANGES WERE NOT MADE."

        Case Else
           Error_Message = "Unknown Error"
      End Select
      MsgBox "PFPSDM Failed to converge." & Chr$(13) & "Error " & Format$(NFLAG, "0") & ":" & Error_Message, MB_ICONEXCLAMATION, Application_Name
      Number_Influent_Points = NumberInfluentPointsToBed1
      Exit Sub
    End Select

       For i = 1 To ITP
           TOUT(i) = T(i, 1)
       Next i

       Call WriteInfoToFile(2, ITP, HadOutputFileBefore, BedCounter, TOUT(), CP())

       'Create array of variable influent points for next call to PFPSDM based on results of previous call
       If BedCounter = Bed.NumberOfBeds Then
          BedLeavingTimes(BedCounter) = Format$(Now, "mm/dd/yy  h:nn:ss AM/PM")
          Exit For
       Else
          For j = 1 To ITP
            Tin(j) = T(j, 1)
            
            CP(NumSelectedComponents_PFPDM, j) = CT_Average(j)
            For K = 1 To (NumSelectedComponents_PFPDM - 1)
                If CP(K, j) < EPS_ERROR_CRITERIA Then CP(K, j) = EPS_ERROR_CRITERIA
                Cin(K, j) = CP(K, j) * SumTimeAverageInfluentConcs * Chemicals(K, 1) / Chemicals(K, 6)   'mg/L
                CP(NumSelectedComponents_PFPDM, j) = CP(NumSelectedComponents_PFPDM, j) - CP(K, j)
            Next K
            Cin(NumSelectedComponents_PFPDM, j) = CP(NumSelectedComponents_PFPDM, j) * SumTimeAverageInfluentConcs * Chemicals(NumSelectedComponents_PFPDM, 1) / Chemicals(NumSelectedComponents_PFPDM, 6)   'mg/L
          Next j

          Number_Influent_Points = ITP

       End If

Next BedCounter
     
    Number_Influent_Points = NumberInfluentPointsToBed1  'Set number of influent data points back to the number that the user specified as being fed to bed number 1

    Tin(1) = 0#
    For i = 1 To NumSelectedComponents_PFPDM
        Cin(i, 1) = 0#
    Next i

       
       frmLoading!Label2.Caption = "Currently Calculating Results for Bed Number 1"
       frmLoading!Label2.Visible = False
       Unload frmLoading
       msg = "PFPDM completed." & Chr$(13) & Chr$(13)
       msg = msg & "Calculation Started at:  " & BedLeavingTimes(0) & Chr$(13)
       msg = msg & "  Calculation Ended at:  " & BedLeavingTimes(Bed.NumberOfBeds) & Chr$(13) & Chr$(13)
       msg = msg & "For a complete summary of the times at which calculations began and ended for each bed, see the file CALCTIME.TXT"
       MsgBox msg, 64, Application_Name

       Open "CALCTIME.TXT" For Output As #1
          Print #1, "Ion Exchange Simulation Software - PFPDM"
          Print #1,
          Print #1,
          Print #1, "Calculation Time Summary for:"
          Print #1, "   " & Trim$(Str$(NumSelectedComponents_PFPDM)) & " component(s)"
          If Bed.NumberOfBeds > 1 Then
             Print #1, "   " & Trim$(Str$(Bed.NumberOfBeds)) & " beds in series"
          Else
             Print #1, "   " & Trim$(Str$(Bed.NumberOfBeds)) & " bed"
          End If
          Print #1, "   " & Trim$(Str$(MC)) & " axial collocation points"
          Print #1, "   " & Trim$(Str$(NC)) & " radial collocation points"
          Print #1, "   " & "Total Simulation Time (minutes) = " & Trim$(Format$(TimeParameters.FinalTime, "0.00#####"))
          Print #1,
          Print #1,
          Print #1, "Initial Time was "; BedLeavingTimes(0)
          Print #1,
          Print #1, "Bed"; Tab(15); "Time Completed Bed"
          Print #1,
          For i = 1 To Bed.NumberOfBeds
              Print #1, i; Tab(15); BedLeavingTimes(i)
          Next i
       Close #1

       Screen.MousePointer = 0
       frmIonExchangeMain.Enabled = True

'----Store the results in the Results variable --------------------

    Results.NPoints = ITP
    
    Results.NComponent = NumSelectedComponents_PFPDM
    Results.Bed = Bed
    Results.Resin = Resin
    Results.Use_Tortuosity_Correlation = Use_Tortuosity_Correlation
    Results.Constant_Tortuosity = Constant_Tortuosity

    For i = 1 To NumSelectedComponents_PFPDM
        Results.Component(i) = Ion(Component_Index_PFPDM(i))
        For j = 1 To ITP
            Results.CP(i, j) = CP(i, j)
'            If j > 2 Then
'               If (CP(i, j) >= .05) And (CP(i, j - 1) < .05) And Flag05 Then
'                  Results.ThroughPut_05(i).T = (T(j, 1) - T(j - 1, 1)) / (CP(i, j) - CP(i, j - 1)) * (.05 - CP(i, j - 1)) + T(j - 1, 1)
'                  Results.ThroughPut_05(i).c = ((CP(i, j) - CP(i, j - 1)) / (T(j, 1) - T(j - 1, 1)) * (Results.ThroughPut_05(i).T - T(j - 1, 1)) + CP(i, j - 1)) * Component(Component_Index_PFPSDM(i)).InitialConcentration
'                  Flag05 = False
'               End If
'               If (CP(i, j) >= .5) And (CP(i, j - 1) < .5) And Flag50 Then
'                  Results.ThroughPut_50(i).T = (T(j, 1) - T(j - 1, 1)) / (CP(i, j) - CP(i, j - 1)) * (.5 - CP(i, j - 1)) + T(j - 1, 1)
'                  Results.ThroughPut_50(i).c = ((CP(i, j) - CP(i, j - 1)) / (T(j, 1) - T(j - 1, 1)) * (Results.ThroughPut_50(i).T - T(j - 1, 1)) + CP(i, j - 1)) * Component(Component_Index_PFPSDM(i)).InitialConcentration
'                  Flag50 = False
'                  If Flag05 Then
'                     Results.ThroughPut_05(i).T = -1#
'                     Results.ThroughPut_05(i).c = -1#
'                     Flag05 = False
'                  End If
'               End If
'               If (CP(i, j) >= .95) And (CP(i, j - 1) < .95) And Flag95 Then
'                  Results.ThroughPut_95(i).T = (T(j, 1) - T(j - 1, 1)) / (CP(i, j) - CP(i, j - 1)) * (.95 - CP(i, j - 1)) + T(j - 1, 1)
'                  Results.ThroughPut_95(i).c = ((CP(i, j) - CP(i, j - 1)) / (T(j, 1) - T(j - 1, 1)) * (Results.ThroughPut_95(i).T - T(j - 1, 1)) + CP(i, j - 1)) * Component(Component_Index_PFPSDM(i)).InitialConcentration
'                  Flag95 = False
'                  If Flag50 Then
'                     Results.ThroughPut_50(i).T = -1#
'                     Results.ThroughPut_50(i).c = -1#
'                     Flag50 = False
'                  End If
'                  If Flag05 Then
'                     Results.ThroughPut_05(i).T = -1#
'                     Results.ThroughPut_05(i).c = -1#
'                     Flag05 = False
'                  End If
'               End If
'            End If
        Next j
'        If Flag95 Then
'           Results.ThroughPut_95(i).T = -1#
'           Results.ThroughPut_95(i).c = -1#
'           Flag95 = False
'        End If
'        If Flag50 Then
'           Results.ThroughPut_50(i).T = -1#
'           Results.ThroughPut_50(i).c = -1#
'           Flag50 = False
'        End If
'        If Flag05 Then
'           Results.ThroughPut_05(i).T = -1#
'           Results.ThroughPut_05(i).c = -1#
'           Flag05 = False
'        End If

'        Flag05 = True  'Set these flags to True such that
'        Flag50 = True  ' Results.ThroughPut_??(I).T and Results.ThroughPut_??(I).C
'        Flag95 = True  ' are calculated for the next compound
    Next i

    For i = 1 To Number_Points_Max
      Results.T(i) = T(i, 1)
    Next i
'------------------------------------------------------------------

    'Enable windows and menus

'    frmPFPSDM.Enabled = True
    frmIonExchangeMain!mnuResults(0).Enabled = True
    frmIonExchangeMain!mnuResults(1).Enabled = True
'    frmPFPSDM!mnuResultsItem(3).Enabled = True

    'write results to file
          Open "pfpdmvb.txt" For Output As #1
             Print #1, "Time (min)"; Tab(12); "BVF";
             For i = 1 To NumSelectedComponents_PFPDM
                 Print #1, Tab(24 + 12 * (i - 1)); "C/CT("; Trim$(Str$(i)); ")";
             Next i
             Print #1,
             Print #1,

             For i = 1 To ITP
                 Print #1, Format$(T(i, 1), "0.000E+00"); Tab(12); Format$(T(i, 2), "0.000E+00");
                 For j = 1 To NumSelectedComponents_PFPDM
                     Print #1, Tab(24 + 12 * (j - 1)); Format$(CP(j, i), "0.0000E+00");
                 Next j
                 Print #1,
             Next i

          Close #1


    Exit Sub

Error_DLL:
  Screen.MousePointer = 0
  Unload frmLoading
  Error_Code = Err
  temp = "Error " & Format$(Error_Code, "0") & ": " & Error$(Error_Code)
  MsgBox "Fatal Error in the DLL. Calculations stopped." & Chr$(13) & temp, MB_ICONEXCLAMATION, Application_Name
  Resume Exit_Call_PFPDM
Exit_Call_PFPDM:
End Sub

Sub GetSelectedComponents (ModelToRun As Integer)
    Dim i As Integer, j As Integer

    Select Case ModelToRun
       Case 0   'PFPDM
          If Cations.Available And Anions.Available Then

          ElseIf Cations.Available Then
             NumSelectedCations = 1
             Cations_Selected(NumSelectedCations) = PresaturantCation
             For i = 1 To frmIonExchangeMain!lstIons(0).ListCount
                 If frmIonExchangeMain!lstIons(0).Selected(i - 1) Then
                    For j = 1 To NumberOfCations
                        If Trim$(Cation(j).Name) = Trim$(frmIonExchangeMain!lstIons(0).List(i - 1)) Then
                           NumSelectedCations = NumSelectedCations + 1
                           Cations_Selected(NumSelectedCations) = j
                           Exit For
                        End If
                    Next j
                 End If
             Next i
             NumSelectedAnions = 0
          ElseIf Anions.Available Then
             NumSelectedAnions = 1
             Anions_Selected(NumSelectedAnions) = PresaturantAnion
             For i = 1 To frmIonExchangeMain!lstIons(1).ListCount
                 If frmIonExchangeMain!lstIons(1).Selected(i - 1) Then
                    For j = 1 To NumberOfAnions
                        If Trim$(Anion(j).Name) = Trim$(frmIonExchangeMain!lstIons(1).List(i - 1)) Then
                           NumSelectedAnions = NumSelectedAnions + 1
                           Anions_Selected(NumSelectedAnions) = j
                           Exit For
                        End If
                    Next j
                 End If
             Next i
             NumSelectedCations = 0

          End If

    End Select

End Sub

Sub WriteInfoToFile (FileTag As Integer, NumbPoints As Integer, HadFileBefore As Integer, BedCounter As Integer, T() As Double, C() As Double)
    'FileTag = 1 --> Write Variable Influent Data to File
    'FileTag = 2 --> Write CP data to file

    Dim i As Integer, j As Integer

    'Write variable influent data to the file VARINF.TXT for FileTag = 1 or
    'write output data to the file CPOUT.TXT for FileTag = 2

    If NumbPoints > 0 Then
       If Not HadFileBefore Then
          HadFileBefore = True
          If FileTag = 1 Then
             Open "VARINF.TXT" For Output As #1
          ElseIf FileTag = 2 Then
             Open "CPOUT.TXT" For Output As #1
          End If
       Else
          If FileTag = 1 Then
             Open "VARINF.TXT" For Append As #1
          ElseIf FileTag = 2 Then
             Open "CPOUT.TXT" For Append As #1
          End If
          Print #1,
          Print #1,
          Print #1,
       End If
       
       If FileTag = 1 Then
          Print #1, "Variable Influent Data to Bed Number "; Trim$(Str$(BedCounter)); " is given below:"
       ElseIf FileTag = 2 Then
          Print #1, "Output Data from Bed Number "; Trim$(Str$(BedCounter)); " is given below:"
       End If

       If FileTag = 1 Then
          Print #1, "Time"; Tab(8); "Cin(1,T)"; Tab(20); "Cin(2, T)"; Tab(32); "Cin(3, T)"; Tab(44); "Cin(4, T)"; Tab(56); "Cin(5, T)"; Tab(68); "Cin(6, T)"
          Print #1, "(min)"; Tab(8); "(mg/L)"; Tab(20); "(mg/L)"; Tab(32); "(mg/L)"; Tab(44); "(mg/L)"; Tab(56); "(mg/L)"; Tab(68); "(mg/L)"
       ElseIf FileTag = 2 Then
          Print #1, "Time"; Tab(8); "CP(1,T)/CT"; Tab(20); "CP(2,T)/CT"; Tab(32); "CP(3,T)/CT"; Tab(44); "CP(4,T)/CT"; Tab(56); "CP(5,T)/CT"; Tab(68); "CP(6,T)/CT"
          Print #1, "(min)"; Tab(8); "(-)"; Tab(20); "(-)"; Tab(32); "(-)"; Tab(44); "(-)"; Tab(56); "(-)"; Tab(68); "(-)"
       End If

       For i = 1 To NumbPoints
           Print #1, Format$(T(i), "0.0");
           For j = 1 To NumSelectedComponents_PFPDM
               Print #1, Tab(12 * (j - 1) + 8); Format$(C(j, i), "0.000E+00");
           Next j
           Print #1,
       Next i
       Close #1
    End If

End Sub

