Option Strict Off
Option Explicit On
Module ModelPSDM
	
	Const ModelPSDM_IN_PathFile As String = "PSDM1.IN"
	Const ModelPSDM_IN_Main As String = "PSDM2.IN"
	Const ModelPSDM_OUT_SuccessFlag As String = "PSDM1.OUT"
	Const ModelPSDM_OUT_Main As String = "PSDM2.OUT"
	Const ModelPSDM_OUT_CvsT As String = "PSDM3.OUT"
	
	Const ModelPSDM_Version As Double = 1#
	Const ModelPSDM_ExeName As String = "PSDM12.EXE"
	Const ModelPSDM_EofTestMarker As Double = 123456#
	
	Const ModelPSDM_MXCOMP As Short = 6
	Const ModelPSDM_MAXPTS As Short = 400
	Const ModelPSDM_MAXDE As Short = 750
	Private Structure ModelPSDM_Inputs_Type
		Dim NUMB As Short
		<VBFixedArray(ModelPSDM_MXCOMP, 16)> Dim CHEMICALS(, ) As Double
		<VBFixedArray(4)> Dim ADS_PROP() As Double
		<VBFixedArray(3)> Dim C_PROP() As Double
		<VBFixedArray(3)> Dim TT() As Double
		Dim MXX As Short
		Dim NXX As Short
		Dim TotalAxialElementCount As Short
		Dim N_PW As Integer
		Dim NINI As Short
		<VBFixedArray(ModelPSDM_MAXPTS)> Dim TINI() As Double
		
		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			'UPGRADE_WARNING: Lower bound of array CHEMICALS was changed from 1,1 to 0,0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim CHEMICALS(ModelPSDM_MXCOMP, 16)
			'UPGRADE_WARNING: Lower bound of array ADS_PROP was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim ADS_PROP(4)
			'UPGRADE_WARNING: Lower bound of array C_PROP was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim C_PROP(3)
			'UPGRADE_WARNING: Lower bound of array TT was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim TT(3)
			'UPGRADE_WARNING: Lower bound of array TINI was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim TINI(ModelPSDM_MAXPTS)
		End Sub
	End Structure
	'UPGRADE_WARNING: Arrays in structure ModelPSDM_Inputs may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Dim ModelPSDM_Inputs As ModelPSDM_Inputs_Type
	'Private Type ModelPSDM_Inputs2_Type
	'  CINI(1 To ModelPSDM_MXCOMP, 1 To ModelPSDM_MAXPTS) As Double
	'End Type
	'UPGRADE_WARNING: Lower bound of array ModelPSDM_Inputs_CINI was changed from 1,1 to 0,0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Dim ModelPSDM_Inputs_CINI(ModelPSDM_MXCOMP, ModelPSDM_MAXPTS) As Double
	
	Private Structure ModelPSDM_Outputs_Type
		<VBFixedArray(15)> Dim VARS1() As Double
		<VBFixedArray(ModelPSDM_MXCOMP, 19)> Dim VARS2(, ) As Double
		Dim NITP As Short
		<VBFixedArray(ModelPSDM_MAXPTS)> Dim T() As Double
		Dim NFLAG As Short
		
		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			'UPGRADE_WARNING: Lower bound of array VARS1 was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim VARS1(15)
			'UPGRADE_WARNING: Lower bound of array VARS2 was changed from 1,1 to 0,0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim VARS2(ModelPSDM_MXCOMP, 19)
			'UPGRADE_WARNING: Lower bound of array T was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim T(ModelPSDM_MAXPTS)
		End Sub
	End Structure
	'UPGRADE_WARNING: Arrays in structure ModelPSDM_Outputs may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Dim ModelPSDM_Outputs As ModelPSDM_Outputs_Type
	'Private Type ModelPSDM_Outputs2_Type
	'  CPVB(1 To ModelPSDM_MXCOMP, 1 To ModelPSDM_MAXPTS) As Double
	'End Type
	'UPGRADE_WARNING: Lower bound of array ModelPSDM_Outputs_CPVB was changed from 1,1 to 0,0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Dim ModelPSDM_Outputs_CPVB(ModelPSDM_MXCOMP, ModelPSDM_MAXPTS) As Double
	
	
	
	
	
	Const ModelPSDM_declarations_end As Boolean = True
	
	
	Sub ModelPSDM_Go()
		Dim Failed As Boolean
		Call ModelPSDM_WritePathFile()
		Call ModelPSDM_WriteMainFile(Failed)
		If (Failed) Then Exit Sub
		Call ModelPSDM_CallEXE()
		Call ModelPSDM_ProcessOutput()
		If (ModelIO_IsKeepTempFiles() = False) Then
			Call ModelPSDM_RemoveLinkFiles()
		End If
	End Sub
	
	
	Sub ModelPSDM_RemoveLinkFiles()
		Call KillFile_If_Exists(Exe_Path & "\" & ModelPSDM_IN_PathFile)
		Call KillFile_If_Exists(Exe_Path & "\" & ModelPSDM_IN_Main)
		Call KillFile_If_Exists(Exe_Path & "\" & ModelPSDM_OUT_SuccessFlag)
		Call KillFile_If_Exists(Exe_Path & "\" & ModelPSDM_OUT_Main)
		Call KillFile_If_Exists(Exe_Path & "\" & ModelPSDM_OUT_CvsT)
	End Sub
	Sub ModelPSDM_CallEXE()
		Dim CmdLine As String
		Call ChangeDir_Exes()
		CmdLine = ModelPSDM_ExeName
		Call ModelIO_Timer_Start()
		Call FortranLink_ExecAndWaitForProcess(CmdLine)
		Call ModelIO_Timer_End()
		Call ChangeDir_Main()
	End Sub
	Sub ModelPSDM_ProcessOutput()
		Dim f As Short
		Dim fn_This As String
		Dim NFLAG As Short
		Dim DummyStr1 As String
		Dim temp As String
		Dim i As Short
		Dim j As Short
		Dim EOFTESTMARKER As Double
		Dim Flag05 As Boolean
		Dim Flag50 As Boolean
		Dim Flag95 As Boolean
		'UPGRADE_WARNING: Arrays in structure MI may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim MI As ModelPSDM_Inputs_Type
		'UPGRADE_WARNING: Arrays in structure MO may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim MO As ModelPSDM_Outputs_Type
		'UPGRADE_WARNING: Couldn't resolve default property of object MI. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'

		MO.Initialize()    'Initialization Shang
		If (Results.Initialized = False) Then
			Results.Initialize()           'Initialization is needed , not sure if this is the right place
			Results.Initialized = True
		End If
		MI = ModelPSDM_Inputs
		'READ SUCCESS FLAG OUTPUT FILE.
		f = FreeFile
		fn_This = Exe_Path & "\" & ModelPSDM_OUT_SuccessFlag
		If (Not FileExists(fn_This)) Then
			Call Show_Error("Unable to find output file: Calculations failed.")
			Exit Sub
		End If
		FileOpen(f, fn_This, OpenMode.Input)
		DummyStr1 = LineInput(f)
		Input(f, NFLAG)
		FileClose(f)
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
			temp = "Error #" & Trim(Str(NFLAG)) & ": " & temp
			Call Show_Error("The PSDM failed to converge." & vbCrLf & temp)
			Exit Sub
		Else
			Call Show_Message("PSDM Model Calculations Complete." & vbCrLf & vbCrLf & ModelIO_Timer_SummaryMsg)
		End If
		'READ MAIN OUTPUT FILE.
		fn_This = Exe_Path & "\" & ModelPSDM_OUT_Main
		FileOpen(f, fn_This, OpenMode.Input)
		DummyStr1 = LineInput(f)
		For i = 1 To 15
			Input(f, MO.VARS1(i))
		Next i
		DummyStr1 = LineInput(f)
		For i = 1 To MI.NUMB
			For j = 1 To 19
				Input(f, MO.VARS2(i, j))
			Next j
		Next i
		DummyStr1 = LineInput(f)
		Input(f, MO.NFLAG)
		DummyStr1 = LineInput(f)
		Input(f, EOFTESTMARKER)
		If (False = ModelIO_DoNumberCheck(EOFTESTMARKER, ModelPSDM_EofTestMarker)) Then
			Call Show_Error("The model calculations failed: invalid file format (EOF marker).")
			Exit Sub
		End If
		FileClose(f)
		'READ C-vs-t OUTPUT FILE.
		fn_This = Exe_Path & "\" & ModelPSDM_OUT_CvsT
		FileOpen(f, fn_This, OpenMode.Input)
		DummyStr1 = LineInput(f)
		Input(f, MO.NITP)
		DummyStr1 = LineInput(f)
		For i = 1 To MO.NITP
			Input(f, MO.T(i))
		Next i
		DummyStr1 = LineInput(f)
		For i = 1 To MI.NUMB
			For j = 1 To MO.NITP
				'Input #f, MO.CPVB(i, j)
				Input(f, ModelPSDM_Outputs_CPVB(i, j))
			Next j
		Next i
		DummyStr1 = LineInput(f)
		Input(f, EOFTESTMARKER)
		If (False = ModelIO_DoNumberCheck(EOFTESTMARKER, ModelPSDM_EofTestMarker)) Then
			Call Show_Error("The model calculations failed: invalid file format (EOF marker).")
			Exit Sub
		End If
		FileClose(f)
		'UPGRADE_WARNING: Couldn't resolve default property of object ModelPSDM_Outputs. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ModelPSDM_Outputs = MO
		'TRANSFER OUTPUT DATA TO MORE PERMANENT MEMORY.
		Results.is_psdm_in_room_model = False
		Results.npoints = MO.NITP
		Results.NComponent = MI.NUMB
		'UPGRADE_WARNING: Couldn't resolve default property of object Results.Bed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Results.Bed = Bed
		'UPGRADE_WARNING: Couldn't resolve default property of object Results.Carbon. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Results.Carbon = Carbon

		PSDM_Inputs.Initialize()  'Initialization Shang

		For i = 1 To 15
			PSDM_Inputs.VARS1(i) = MO.VARS1(i)
		Next i
		PSDM_Inputs.VARS1(8) = SF() * 264.17205 * 60 / 10.76391 'Convert m/s to gal/min-ft^2.
		PSDM_Inputs.VARS1(11) = Re()
		PSDM_Inputs.VARS1(12) = Bed.WaterDensity
		PSDM_Inputs.VARS1(13) = Bed.WaterViscosity
		For i = 1 To Number_Component_PFPSDM
			For j = 1 To 18
				PSDM_Inputs.VARS2(i, j) = MO.VARS2(i, j)
			Next j
			PSDM_Inputs.VARS2(i, 6) = Diffl(i)
			PSDM_Inputs.VARS2(i, 18) = SC(i)
			j = Component_Index_PFPSDM(i)
			PSDM_Inputs.VARS2(i, 19) = Component(j).SPDFR
		Next i
		'DETERMINE 5%, 50%, AND 95% SATURATION TIMES.
		Flag05 = True
		Flag50 = True
		Flag95 = True
		'UPGRADE_WARNING: Lower bound of array BrokeThrough was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		Dim BrokeThrough(Number_Component_PFPSDM) As Short
		Dim IsFoulingCase As Short
		'ReDim NumPoints_Before_BrokeThrough(Number_Component_PFPSDM) As Integer
		For i = 1 To Number_Component_PFPSDM
			BrokeThrough(i) = False
			'NumPoints_Before_BrokeThrough(i) = -1
			Results.NumPoints_Before_ThroughPut_100(i) = MO.NITP
		Next i
		IsFoulingCase = False
		For i = 1 To Number_Component_PFPSDM
			j = Component_Index_PFPSDM(i)
			If (Component(j).K_Reduction) Then
				IsFoulingCase = True
			End If
		Next i
		For i = 1 To Number_Component_PFPSDM
			'UPGRADE_WARNING: Couldn't resolve default property of object Results.Component(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Results.Component(i) = Component(Component_Index_PFPSDM(i))
			For j = 1 To MO.NITP
				If (((IsFoulingCase) And (ModelPSDM_Outputs_CPVB(i, j) > 0.9995)) Or (BrokeThrough(i))) Then
					'---- Stop the plot as soon as C/C0>=0.9995; this is accomplished
					'.... by setting .CP = -10000#, which tells the plotting routine to
					'.... stop plotting.
					Results.CP(i, j) = -10000#
					If (Not BrokeThrough(i)) Then
						Results.NumPoints_Before_ThroughPut_100(i) = j - 1
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
					Results.CP(i, j) = ModelPSDM_Outputs_CPVB(i, j)
				End If
				If j > 2 Then
					If (ModelPSDM_Outputs_CPVB(i, j) >= 0.05) And (ModelPSDM_Outputs_CPVB(i, j - 1) < 0.05) And Flag05 Then
						Results.ThroughPut_05(i).T = (MO.T(j) - MO.T(j - 1)) / (ModelPSDM_Outputs_CPVB(i, j) - ModelPSDM_Outputs_CPVB(i, j - 1)) * (0.05 - ModelPSDM_Outputs_CPVB(i, j - 1)) + MO.T(j - 1)
						Results.ThroughPut_05(i).C = ((ModelPSDM_Outputs_CPVB(i, j) - ModelPSDM_Outputs_CPVB(i, j - 1)) / (MO.T(j) - MO.T(j - 1)) * (Results.ThroughPut_05(i).T - MO.T(j - 1)) + ModelPSDM_Outputs_CPVB(i, j - 1)) * Component(Component_Index_PFPSDM(i)).InitialConcentration
						Flag05 = False
					End If
					If (ModelPSDM_Outputs_CPVB(i, j) >= 0.5) And (ModelPSDM_Outputs_CPVB(i, j - 1) < 0.5) And Flag50 Then
						Results.ThroughPut_50(i).T = (MO.T(j) - MO.T(j - 1)) / (ModelPSDM_Outputs_CPVB(i, j) - ModelPSDM_Outputs_CPVB(i, j - 1)) * (0.5 - ModelPSDM_Outputs_CPVB(i, j - 1)) + MO.T(j - 1)
						Results.ThroughPut_50(i).C = ((ModelPSDM_Outputs_CPVB(i, j) - ModelPSDM_Outputs_CPVB(i, j - 1)) / (MO.T(j) - MO.T(j - 1)) * (Results.ThroughPut_50(i).T - MO.T(j - 1)) + ModelPSDM_Outputs_CPVB(i, j - 1)) * Component(Component_Index_PFPSDM(i)).InitialConcentration
						Flag50 = False
						If Flag05 Then
							Results.ThroughPut_05(i).T = -1#
							Results.ThroughPut_05(i).C = -1#
							Flag05 = False
						End If
					End If
					If (ModelPSDM_Outputs_CPVB(i, j) >= 0.95) And (ModelPSDM_Outputs_CPVB(i, j - 1) < 0.95) And Flag95 Then
						Results.ThroughPut_95(i).T = (MO.T(j) - MO.T(j - 1)) / (ModelPSDM_Outputs_CPVB(i, j) - ModelPSDM_Outputs_CPVB(i, j - 1)) * (0.95 - ModelPSDM_Outputs_CPVB(i, j - 1)) + MO.T(j - 1)
						Results.ThroughPut_95(i).C = ((ModelPSDM_Outputs_CPVB(i, j) - ModelPSDM_Outputs_CPVB(i, j - 1)) / (MO.T(j) - MO.T(j - 1)) * (Results.ThroughPut_95(i).T - MO.T(j - 1)) + ModelPSDM_Outputs_CPVB(i, j - 1)) * Component(Component_Index_PFPSDM(i)).InitialConcentration
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
			Next j
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
			Flag05 = True 'Set these flags to True such that
			Flag50 = True ' Results.ThroughPut_??(I).T and Results.ThroughPut_??(I).C
			Flag95 = True ' are calculated for the next compound
		Next i
		For i = 1 To Number_Points_Max
			Results.T(i) = MO.T(i)
		Next i
		'ENABLE RESULTS MENU COMMANDS.
		frmMain.mnuResultsItem(0).Enabled = True
		If (NData_Points > 0) Then
			frmMain.mnuResultsItem(3).Enabled = True
		End If
	End Sub
	Sub ModelPSDM_WriteMainFile(ByRef Failed As Boolean)
		'UPGRADE_WARNING: Arrays in structure MI may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim MI As ModelPSDM_Inputs_Type
		Dim i As Short
		Dim j As Short
		Dim Number_Equations As Short
		Dim WorkSpace_Size As Integer
		Dim msg As String
		Dim f As Short
		Dim fn_This As String

		MI.Initialize()   'Initialziation Shang

		Failed = False
		'CALCULATE WORKSPACE SIZE.
		Number_Equations = Number_Component_PFPSDM * (MC * (NC + 1) - 1)
		If Number_Equations > Max_Equations_DGEAR Then
			msg = "Maximum number of equations PSDM can solve = " & Str(Max_Equations_DGEAR) & vbCrLf
			msg = msg & "Current number of equations specified for PSDM to solve = " & Str(Number_Equations) & vbCrLf & vbCrLf
			msg = msg & "(No. of Equations PSDM Must Solve) = NCOMP*(MC*(NC+1)-1)" & vbCrLf & vbCrLf
			msg = msg & "Please ensure that this number does not exceed the maximum." & vbCrLf & vbCrLf
			msg = msg & "Note:  " & vbCrLf
			msg = msg & "    NCOMP = Number of Components = " & Str(Number_Component_PFPSDM) & vbCrLf
			msg = msg & "    MC = Number of Axial Collocation Points = " & Str(MC) & vbCrLf
			msg = msg & "    NC = Number of Radial Collocation Points = " & Str(NC) & vbCrLf
			Call Show_Error(msg)
			Failed = True
			Exit Sub
		End If
		WorkSpace_Size = Number_Equations ^ 2 + 2 * Number_Equations
		'PREPARE INPUTS.
		MI.NUMB = Number_Component_PFPSDM
		For i = 1 To MI.NUMB
			j = Component_Index_PFPSDM(i)
			MI.CHEMICALS(i, 1) = Component(j).MW
			'CONVERT Co FROM mg/L TO ug/L.
			MI.CHEMICALS(i, 2) = Component(j).InitialConcentration * 1000#
			MI.CHEMICALS(i, 3) = Component(j).MolarVolume
			'CONVERT K FROM (mg/g)*(L/mg)^(1/n) to (umol/g)*(L/umol)^(1/n).
			MI.CHEMICALS(i, 4) = Component(j).Use_K * (1000# / Component(j).MW) ^ (1# - Component(j).Use_OneOverN)
			MI.CHEMICALS(i, 5) = Component(j).Use_OneOverN
			MI.CHEMICALS(i, 6) = Component(j).kf
			MI.CHEMICALS(i, 7) = Component(j).Ds
			MI.CHEMICALS(i, 8) = Component(j).Dp
			If (Bed.Phase = 0) Then
				If (Component(j).K_Reduction) And (Bed.Water_Correlation.Coeff(1) <> 1# And Bed.Water_Correlation.Coeff(2) <> 0# And Bed.Water_Correlation.Coeff(3) <> 0# And Bed.Water_Correlation.Coeff(4) <> 0#) Then
					MI.CHEMICALS(i, 9) = Bed.Water_Correlation.Coeff(1) * Component(j).Correlation.Coeff(1) + Component(j).Correlation.Coeff(2)
					MI.CHEMICALS(i, 10) = Bed.Water_Correlation.Coeff(2) * Component(j).Correlation.Coeff(1)
					MI.CHEMICALS(i, 11) = Bed.Water_Correlation.Coeff(3) * Component(j).Correlation.Coeff(1)
					MI.CHEMICALS(i, 12) = Bed.Water_Correlation.Coeff(4) * Component(j).Correlation.Coeff(1)
				Else
					MI.CHEMICALS(i, 9) = 1#
					MI.CHEMICALS(i, 10) = 0#
					MI.CHEMICALS(i, 11) = 0#
					MI.CHEMICALS(i, 12) = 0#
				End If
			Else
				MI.CHEMICALS(i, 9) = 1#
				MI.CHEMICALS(i, 10) = 0#
				MI.CHEMICALS(i, 11) = 0#
				MI.CHEMICALS(i, 12) = 0#
			End If
			MI.CHEMICALS(i, 13) = Component(j).Tortuosity
			If ((Component(j).Constant_Tortuosity) And (Component(j).Use_Tortuosity_Correlation)) Then
				MI.CHEMICALS(i, 14) = 2#
				MI.CHEMICALS(i, 15) = 0#
			Else
				If (Component(j).Use_Tortuosity_Correlation) Then
					MI.CHEMICALS(i, 14) = 0.334
					MI.CHEMICALS(i, 15) = 0.00000661
				Else
					MI.CHEMICALS(i, 14) = 2#
					MI.CHEMICALS(i, 15) = 0#
				End If
			End If
			MI.CHEMICALS(i, 16) = 100000# 'in minutes
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
		MI.C_PROP(3) = Carbon.ParticleRadius * 100# 'To convert in cm
		MI.TT(1) = TimeP.End_Renamed
		'Test value of Tinit
		If (TimeP.Init <= 0#) Then
			MI.TT(2) = 0.0001
		Else
			MI.TT(2) = TimeP.Init
		End If
		MI.TT(3) = TimeP.Step_Renamed
		MI.MXX = MC
		MI.NXX = NC
		MI.TotalAxialElementCount = Bed.NumberOfBeds
		MI.N_PW = WorkSpace_Size
		MI.NINI = Number_Influent_Points
		For j = 1 To MI.NINI
			MI.TINI(j) = T_Influent(j)
			For i = 1 To MI.NUMB
				'CONVERT FROM mg/L TO ug/L.
				ModelPSDM_Inputs_CINI(i, j) = C_Influent(Component_Index_PFPSDM(i), j) * 1000#
			Next i
		Next j
		'WRITE INPUT FILE.
		f = FreeFile
		fn_This = Exe_Path & "\" & ModelPSDM_IN_Main
		FileOpen(f, fn_This, OpenMode.Output)
		Call WriteFortranInput(f, ModelPSDM_Version, "MODULE_VERSION")
		Call WriteFortranInput(f, MI.NUMB, "NUMB, number of chemicals in simulation")
		For i = 1 To MI.NUMB
			Call WriteFortranInput(f, MI.CHEMICALS(i, 1), "CHEMICALS(" & Trim(Str(i)) & ",1), molecular weight, g/gmol")
			Call WriteFortranInput(f, MI.CHEMICALS(i, 2), "CHEMICALS(" & Trim(Str(i)) & ",2), influent concentration, ug/L")
			Call WriteFortranInput(f, MI.CHEMICALS(i, 3), "CHEMICALS(" & Trim(Str(i)) & ",3), molar volume, cm^3/gmol")
			Call WriteFortranInput(f, MI.CHEMICALS(i, 4), "CHEMICALS(" & Trim(Str(i)) & ",4), Freundlich K, (umol/g)*(L/umol)^(1/n)")
			Call WriteFortranInput(f, MI.CHEMICALS(i, 5), "CHEMICALS(" & Trim(Str(i)) & ",5), Freundlich 1/n, dimless")
			Call WriteFortranInput(f, MI.CHEMICALS(i, 6), "CHEMICALS(" & Trim(Str(i)) & ",6), film transfer coefficient (kf), cm/s ")
			Call WriteFortranInput(f, MI.CHEMICALS(i, 7), "CHEMICALS(" & Trim(Str(i)) & ",7), surface diffusion coefficient (Ds), cm^2/s")
			Call WriteFortranInput(f, MI.CHEMICALS(i, 8), "CHEMICALS(" & Trim(Str(i)) & ",8), pore diffusion coefficient (Dp), cm^2/s")
			Call WriteFortranInput(f, MI.CHEMICALS(i, 9), "CHEMICALS(" & Trim(Str(i)) & ",9) = RK1(" & Trim(Str(i)) & "), fouling correlation coef. #1, dimless")
			Call WriteFortranInput(f, MI.CHEMICALS(i, 10), "CHEMICALS(" & Trim(Str(i)) & ",10) = RK2(" & Trim(Str(i)) & "), fouling correlation coef. #2, 1/min")
			Call WriteFortranInput(f, MI.CHEMICALS(i, 11), "CHEMICALS(" & Trim(Str(i)) & ",11) = RK3(" & Trim(Str(i)) & "), fouling correlation coef. #3, dimless")
			Call WriteFortranInput(f, MI.CHEMICALS(i, 12), "CHEMICALS(" & Trim(Str(i)) & ",12) = RK4(" & Trim(Str(i)) & "), fouling correlation coef. #4, 1/min")
			Call WriteFortranInput(f, MI.CHEMICALS(i, 13), "CHEMICALS(" & Trim(Str(i)) & ",13) = TORTU(" & Trim(Str(i)) & "), tortuosity (never used?), dimless")
			Call WriteFortranInput(f, MI.CHEMICALS(i, 14), "CHEMICALS(" & Trim(Str(i)) & ",14) = TOR(" & Trim(Str(i)) & "), tortuosity coef., dimless")
			Call WriteFortranInput(f, MI.CHEMICALS(i, 15), "CHEMICALS(" & Trim(Str(i)) & ",15) = PART(" & Trim(Str(i)) & "), part. coef., dimless")
			Call WriteFortranInput(f, MI.CHEMICALS(i, 16), "CHEMICALS(" & Trim(Str(i)) & ",16) = TTORTU(" & Trim(Str(i)) & "), time parameter, min")
		Next i
		Call WriteFortranInput(f, MI.ADS_PROP(1), "ADS_PROP(1), length of bed, m")
		Call WriteFortranInput(f, MI.ADS_PROP(2), "ADS_PROP(2), diameter of bed, m")
		Call WriteFortranInput(f, MI.ADS_PROP(3), "ADS_PROP(3), weight of adsorbent in bed, kg")
		Call WriteFortranInput(f, MI.ADS_PROP(4), "ADS_PROP(4), influent flow rate, m^3/s")
		Call WriteFortranInput(f, MI.C_PROP(1), "C_PROP(1), particle void fraction, dimless")
		Call WriteFortranInput(f, MI.C_PROP(2), "C_PROP(2), particle density, g/cm^3")
		Call WriteFortranInput(f, MI.C_PROP(3), "C_PROP(3), particle radius, cm")
		Call WriteFortranInput(f, MI.TT(1), "TT(1), time to end simulation, minutes")
		Call WriteFortranInput(f, MI.TT(2), "TT(2), time to begin simulation, minutes")
		Call WriteFortranInput(f, MI.TT(3), "TT(3), time step, minutes")
		Call WriteFortranInput(f, MI.MXX, "MXX, number of axial collocation points, dimless")
		Call WriteFortranInput(f, MI.NXX, "NXX, number of radial collocation points, dimless")
		Call WriteFortranInput(f, MI.TotalAxialElementCount, "TotalAxialElementCount, number of axial elements, dimless")
		Call WriteFortranInput(f, MI.N_PW, "N_PW, equation workspace size, bytes")
		Call WriteFortranInput(f, MI.NINI, "NINI, number of influent concentration points, dimless")
		PrintLine(f, "TINI(i), time profile for CINI() array, minutes")
		For i = 1 To MI.NINI
			PrintLine(f, MI.TINI(i))
		Next i
		PrintLine(f, "CINI(i,j), influent concentration profile, ug/L")
		For i = 1 To MI.NUMB
			For j = 1 To MI.NINI
				PrintLine(f, ModelPSDM_Inputs_CINI(i, j))
			Next j
		Next i
		Call WriteFortranInput(f, ModelPSDM_EofTestMarker, "EOFTESTMARKER")
		FileClose(f)
		'STORE FOR LATER USE.
		'UPGRADE_WARNING: Couldn't resolve default property of object ModelPSDM_Inputs. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ModelPSDM_Inputs = MI
	End Sub
	Sub ModelPSDM_WritePathFile()
		Dim f As Short
		Dim fn_This As String
		Dim qq As String
		qq = Chr(34)
		f = FreeFile
		fn_This = Exe_Path & "\" & ModelPSDM_IN_PathFile
		FileOpen(f, fn_This, OpenMode.Output)
		PrintLine(f, qq & ModelPSDM_IN_Main & qq)
		PrintLine(f, qq & ModelPSDM_OUT_SuccessFlag & qq)
		PrintLine(f, qq & ModelPSDM_OUT_Main & qq)
		PrintLine(f, qq & ModelPSDM_OUT_CvsT & qq)
		FileClose(f)
	End Sub
	
	
	'Return value:
	'  TRUE = Okay to call the PSDM
	'  FALSE = Something went wrong, ABORT!  ABORT!
	Function Prepare_To_Run_PSDM() As Short
		Dim i As Short
		Dim j As Short
		Dim Num_K_Reduction As Short
		Dim Num_A_and_Not_B As Short
		Dim Num_Not_a_and_B As Short
		'UPGRADE_WARNING: Lower bound of array Name_A_and_Not_B was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		Dim Name_A_and_Not_B(Number_Compo_Max) As String
		'UPGRADE_WARNING: Lower bound of array Name_Not_A_and_B was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		Dim Name_Not_A_and_B(Number_Compo_Max) As String
		Dim Is_A As Short
		Dim Is_B As Short
		Dim msg As String
		Dim RetVal As Short
		'
		' PERFORM SEVERAL VERIFICATIONS BEFORE RUNNING THE PSDM.
		'
		If (TimeP.Init > TimeP.End_Renamed) Then
			Call Show_Error("The initial simulation time (" & TimeP.Init / 24# / 60# & " days) is greater than the " & "final simulation time (" & TimeP.End_Renamed / 24# / 60# & " days).  PSDM cannot be run until this is fixed.")
			Prepare_To_Run_PSDM = False
			Exit Function
		End If
		If (TimeP.Step_Renamed < ((TimeP.End_Renamed - TimeP.Init) / (Number_Points_Max - 1))) Then
			Call Show_Error("The simulation time step (" & TimeP.Step_Renamed / 24# / 60# & " days) is too small.  The " & "maximum number of points is 400.  PSDM cannot be run " & "until this is fixed.")
			Prepare_To_Run_PSDM = False
			Exit Function
		End If
		Call AllModels_Verify_Selected_Components(MODELTYPE_PSDM)
		If (Number_Component_PFPSDM = 0) Then
			Prepare_To_Run_PSDM = False
			Exit Function
		End If
		For i = 1 To Number_Component_PFPSDM
			For j = i + 1 To Number_Component_PFPSDM
				If Trim(Component(Component_Index_PFPSDM(i)).Name) = Trim(Component(Component_Index_PFPSDM(j)).Name) Then
					Call Show_Error("Components " & VB6.Format(Component_Index_PFPSDM(i), "0") & " and " & VB6.Format(Component_Index_PFPSDM(j), "0") & " have the same name." & vbCrLf & "Please change one before running the PSDM.")
					Prepare_To_Run_PSDM = False
					Exit Function
				End If
			Next j
		Next i
		'
		'---- Make sure # PSDM fouling components is <= 1.
		'
		Num_K_Reduction = 0
		For i = 0 To frmMain.lstComponents.Items.Count - 1
			If (frmMain.lstComponents.GetSelected(i)) Then
				If (Component(i + 1).K_Reduction) Then
					Num_K_Reduction = Num_K_Reduction + 1
				End If
			End If
		Next i
		If (Num_K_Reduction > 1) Then
			Call Show_Error("There are currently " & Trim(Str(Num_K_Reduction)) & " components specified for fouling.  Only 1 may be " & "specified for a run of the PSDM.")
			Prepare_To_Run_PSDM = False
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
		For i = 0 To frmMain.lstComponents.Items.Count - 1
			If (frmMain.lstComponents.GetSelected(i)) Then
				Is_A = (Component(i + 1).Use_Tortuosity_Correlation)
				Is_B = (Component(i + 1).K_Reduction)
				'---- Check for A and not B case:
				If ((Is_A) And (Not Is_B)) Then
					Num_A_and_Not_B = Num_A_and_Not_B + 1
					Name_A_and_Not_B(Num_A_and_Not_B) = Trim(Component(i + 1).Name)
				End If
				'---- Check for not A and B case:
				If ((Not Is_A) And (Is_B)) Then
					Num_Not_a_and_B = Num_Not_a_and_B + 1
					Name_Not_A_and_B(Num_Not_a_and_B) = Trim(Component(i + 1).Name)
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
			RetVal = MsgBox(msg, MsgBoxStyle.Question + MsgBoxStyle.YesNo, AppName_For_Display_Long)
			If (RetVal = MsgBoxResult.No) Then
				Prepare_To_Run_PSDM = False
				Exit Function
			End If
		End If
		Prepare_To_Run_PSDM = True
	End Function
End Module