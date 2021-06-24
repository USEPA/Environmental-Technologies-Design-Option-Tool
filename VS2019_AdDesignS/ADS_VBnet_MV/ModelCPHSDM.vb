Option Strict Off
Option Explicit On
Module ModelCPHSDM
	
	Const ModelCPHSDM_IN_PathFile As String = "CPHSDM1.IN"
	Const ModelCPHSDM_IN_Main As String = "CPHSDM2.IN"
	Const ModelCPHSDM_OUT_SuccessFlag As String = "CPHSDM1.OUT"
	Const ModelCPHSDM_OUT_Main As String = "CPHSDM2.OUT"
	
	Const ModelCPHSDM_Version As Double = 1#
	Const ModelCPHSDM_ExeName As String = "CPHSDM2.EXE"
	Const ModelCPHSDM_EofTestMarker As Double = 123456#
	
	'Const ModelCPHSDM_NMAX = 1
	Private Structure ModelCPHSDM_Inputs_Type
		<VBFixedArray(6)> Dim Bed() As Double
		<VBFixedArray(4)> Dim Compo() As Double
		<VBFixedArray(2)> Dim Kine() As Double
		
		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			'UPGRADE_WARNING: Lower bound of array Bed was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			If Bed Is Nothing Then ReDim Bed(6)
			'UPGRADE_WARNING: Lower bound of array Compo was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			If Compo Is Nothing Then ReDim Compo(4)
			'UPGRADE_WARNING: Lower bound of array Kine was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			If Kine Is Nothing Then ReDim Kine(2)
		End Sub
	End Structure
	'UPGRADE_WARNING: Arrays in structure ModelCPHSDM_Inputs may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Dim ModelCPHSDM_Inputs As ModelCPHSDM_Inputs_Type
	Private Structure ModelCPHSDM_Outputs_Type
		<VBFixedArray(210)> Dim TACT() As Double
		<VBFixedArray(210)> Dim CC() As Double
		<VBFixedArray(7)> Dim PARAM() As Double
		Dim ER_FLAG As Short
		
		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			'UPGRADE_WARNING: Lower bound of array TACT was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			If TACT Is Nothing Then ReDim TACT(210)
			'UPGRADE_WARNING: Lower bound of array CC was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			If CC Is Nothing Then ReDim CC(210)
			'UPGRADE_WARNING: Lower bound of array PARAM was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			If PARAM Is Nothing Then ReDim PARAM(7)
		End Sub
	End Structure
	'UPGRADE_WARNING: Arrays in structure ModelCPHSDM_Outputs may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Dim ModelCPHSDM_Outputs As ModelCPHSDM_Outputs_Type
	
	
	
	
	Const ModelCPHSDM_declarations_end As Boolean = True
	
	
	Sub ModelCPHSDM_Go()
		Call ModelCPHSDM_WritePathFile()
		Call ModelCPHSDM_WriteMainFile()
		Call ModelCPHSDM_CallEXE()
		Call ModelCPHSDM_ProcessOutput()
		If (ModelIO_IsKeepTempFiles() = False) Then
			Call ModelCPHSDM_RemoveLinkFiles()
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
		Call ChangeDir_Exes()
		CmdLine = ModelCPHSDM_ExeName
		Call ModelIO_Timer_Start()
		Call FortranLink_ExecAndWaitForProcess(CmdLine)
		Call ModelIO_Timer_End()
		Call ChangeDir_Main()
	End Sub
	Sub ModelCPHSDM_ProcessOutput()
		Dim f As Short
		Dim fn_This As String
		Dim ER_FLAG As Short
		Dim DummyStr1 As String
		Dim temp As String
		Dim i As Short
		Dim J As Short
		'UPGRADE_WARNING: Arrays in structure MO may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim MO As ModelCPHSDM_Outputs_Type
		Dim EOFTESTMARKER As Double
		Dim Flag05 As Boolean
		Dim Flag50 As Boolean
		Dim Flag95 As Boolean

		MO.Initialize()  'Initialization is needed Shang

		'READ SUCCESS FLAG OUTPUT FILE.
		f = FreeFile
		fn_This = Exe_Path & "\" & ModelCPHSDM_OUT_SuccessFlag
		If (Not FileExists(fn_This)) Then
			Call Show_Error("Unable to find output file: Calculations failed.")
			Exit Sub
		End If
		FileOpen(f, fn_This, OpenMode.Input)
		DummyStr1 = LineInput(f)
		Input(f, ER_FLAG)
		FileClose(f)
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
			Call Show_Message("CPHSDM Model Calculations Complete." & vbCrLf & vbCrLf & ModelIO_Timer_SummaryMsg)
		End If
		'READ MAIN OUTPUT FILE.
		fn_This = Exe_Path & "\" & ModelCPHSDM_OUT_Main
		FileOpen(f, fn_This, OpenMode.Input)
		DummyStr1 = LineInput(f)
		For i = 1 To 210
			Input(f, MO.TACT(i))
		Next i
		DummyStr1 = LineInput(f)
		For i = 1 To 210
			Input(f, MO.CC(i))
		Next i
		DummyStr1 = LineInput(f)
		For i = 1 To 7
			Input(f, MO.PARAM(i))
		Next i
		DummyStr1 = LineInput(f)
		Input(f, EOFTESTMARKER)
		If (False = ModelIO_DoNumberCheck(EOFTESTMARKER, ModelCPHSDM_EofTestMarker)) Then
			Call Show_Error("The model calculations failed: invalid file format (EOF marker).")
			Exit Sub
		End If
		FileClose(f)
		'UPGRADE_WARNING: Couldn't resolve default property of object ModelCPHSDM_Outputs. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ModelCPHSDM_Outputs = MO
		If (CPM_Results.Initialized = False) Then    'Not sure if it is the right place for initialization Shang
			CPM_Results.Initialize()
			CPM_Results.Initialized = True
		End If
		'TRANSFER OUTPUT DATA TO MORE PERMANENT MEMORY.
		For i = 1 To CPM_Max_Points
			CPM_Results.T(i) = MO.TACT(i) 'TACT(I) is in days
			CPM_Results.C_Over_C0(i) = MO.CC(i) 'CC(I) is dimensionless
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
		'UPGRADE_WARNING: Couldn't resolve default property of object CPM_Results.Bed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CPM_Results.Bed = Bed
		'UPGRADE_WARNING: Couldn't resolve default property of object CPM_Results.Carbon. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CPM_Results.Carbon = Carbon
		'UPGRADE_WARNING: Couldn't resolve default property of object CPM_Results.Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CPM_Results.Component = Component(Component_Index_CPM)
		''''CPM_Results.Constant_Tortuosity = Constant_Tortuosity
		''''CPM_Results.Use_Tortuosity_Correlation = Use_Tortuosity_Correlation
		Flag05 = True
		Flag50 = True
		Flag95 = True
		For J = 1 To CPM_Max_Points
			If (J > 2) Then
				If (MO.CC(J) >= 0.05) And (MO.CC(J - 1) < 0.05) And Flag05 Then
					CPM_Results.ThroughPut_05.T = (MO.TACT(J) - MO.TACT(J - 1)) / (MO.CC(J) - MO.CC(J - 1)) * (0.05 - MO.CC(J - 1)) + MO.TACT(J - 1)
					CPM_Results.ThroughPut_05.C = ((MO.CC(J) - MO.CC(J - 1)) / (MO.TACT(J) - MO.TACT(J - 1)) * (CPM_Results.ThroughPut_05.T - MO.TACT(J - 1)) + MO.CC(J - 1)) * CPM_Results.Component.InitialConcentration
					Flag05 = False
				End If
				If (MO.CC(J) >= 0.5) And (MO.CC(J - 1) < 0.5) And Flag50 Then
					CPM_Results.ThroughPut_50.T = (MO.TACT(J) - MO.TACT(J - 1)) / (MO.CC(J) - MO.CC(J - 1)) * (0.5 - MO.CC(J - 1)) + MO.TACT(J - 1)
					CPM_Results.ThroughPut_50.C = ((MO.CC(J) - MO.CC(J - 1)) / (MO.TACT(J) - MO.TACT(J - 1)) * (CPM_Results.ThroughPut_50.T - MO.TACT(J - 1)) + MO.CC(J - 1)) * CPM_Results.Component.InitialConcentration
					Flag50 = False
				End If
				If (MO.CC(J) >= 0.95) And (MO.CC(J - 1) < 0.95) And Flag95 Then
					CPM_Results.ThroughPut_95.T = (MO.TACT(J) - MO.TACT(J - 1)) / (MO.CC(J) - MO.CC(J - 1)) * (0.95 - MO.CC(J - 1)) + MO.TACT(J - 1)
					CPM_Results.ThroughPut_95.C = ((MO.CC(J) - MO.CC(J - 1)) / (MO.TACT(J) - MO.TACT(J - 1)) * (CPM_Results.ThroughPut_95.T - MO.TACT(J - 1)) + MO.CC(J - 1)) * CPM_Results.Component.InitialConcentration
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
		Dim f As Short
		Dim fn_This As String
		'UPGRADE_WARNING: Arrays in structure MI may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim MI As ModelCPHSDM_Inputs_Type
		Dim i As Short
		Dim J As Short
		Dim A1 As Double
		Dim A2 As Double
		Dim A3 As Double
		Dim A4 As Double

		MI.Initialize()   'Initialization is needed Shang

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
		Dim DG1 As Double
		Dim DG2 As Double
		Dim KovK0 As Double
		Dim DG As Double
		Dim TAU As Double
		Dim T_Minut As Double
		If (Bed.Phase = 0) Then
			If (Component(J).K_Reduction) And (Bed.Water_Correlation.Coeff(1) <> 1# And Bed.Water_Correlation.Coeff(2) <> 0# And Bed.Water_Correlation.Coeff(3) <> 0# And Bed.Water_Correlation.Coeff(4) <> 0#) Then
				i = 0
				DG2 = 1#
				KovK0 = 1#
ModelCPHSDM_WriteMainFile_NextIteration: 
				If (i < Max_Number_Fouling_Iterations) And (System.Math.Abs(1# - DG1 / DG2) > 0.01) Then
					DG1 = 1000# * MI.Bed(2) * (MI.Bed(2) / MI.Bed(3)) / (1 - MI.Bed(2) / MI.Bed(3)) / MI.Compo(2) * KovK0 * MI.Compo(3) * MI.Compo(2) ^ MI.Compo(4)
					'TAU = EBST * epsilon
					TAU = MI.Bed(4) * (1 - MI.Bed(2) / MI.Bed(3))
					T_Minut = TAU * (DG1 + 1)
					KovK0 = A1 + A2 * T_Minut + A3 * System.Math.Exp(A4 * T_Minut)
					DG2 = 1000# * MI.Bed(2) * (MI.Bed(2) / MI.Bed(3)) / (1 - MI.Bed(2) / MI.Bed(3)) / MI.Compo(2) * KovK0 * MI.Compo(3) * MI.Compo(2) ^ MI.Compo(4)
					i = i + 1
					GoTo ModelCPHSDM_WriteMainFile_NextIteration
				End If
				If i < Max_Number_Fouling_Iterations Then
					MI.Compo(3) = MI.Compo(3) * KovK0
				Else
					Call Show_Error("The iterations to evaluate the capacity reduction " & "due to fouling did not converge." & vbCrLf & "It will be assumed there is no fouling.")
				End If
			End If
		End If
		'WRITE INPUT FILE.
		f = FreeFile
		fn_This = Exe_Path & "\" & ModelCPHSDM_IN_Main
		'fn_This = App.Path & "\" & ModelECM_IN_Main
		FileOpen(f, fn_This, OpenMode.Output)
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
		FileClose(f)
		'STORE FOR LATER USE.
		'UPGRADE_WARNING: Couldn't resolve default property of object ModelCPHSDM_Inputs. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ModelCPHSDM_Inputs = MI
	End Sub
	Sub ModelCPHSDM_WritePathFile()
		Dim f As Short
		Dim fn_This As String
		Dim qq As String
		qq = Chr(34)
		f = FreeFile
		fn_This = Exe_Path & "\" & ModelCPHSDM_IN_PathFile
		FileOpen(f, fn_This, OpenMode.Output)
		PrintLine(f, qq & ModelCPHSDM_IN_Main & qq)
		PrintLine(f, qq & ModelCPHSDM_OUT_SuccessFlag & qq)
		PrintLine(f, qq & ModelCPHSDM_OUT_Main & qq)
		FileClose(f)
	End Sub
End Module