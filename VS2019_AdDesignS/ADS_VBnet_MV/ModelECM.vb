Option Strict Off
Option Explicit On
Module ModelECM
	
	Const ModelECM_IN_PathFile As String = "ECM1.IN"
	Const ModelECM_IN_Main As String = "ECM2.IN"
	Const ModelECM_OUT_SuccessFlag As String = "ECM1.OUT"
	Const ModelECM_OUT_Main As String = "ECM2.OUT"
	
	Const ModelECM_Version As Double = 1#
	Const ModelECM_ExeName As String = "ECM5.EXE"
	Const ModelECM_EofTestMarker As Double = 123456#
	
	Public Const MODELTYPE_PSDM As Short = 0
	Public Const MODELTYPE_CPHSDM As Short = 1
	Public Const MODELTYPE_ECM As Short = 2
	
	Const ModelECM_NMAX As Short = 20
	Private Structure ModelECM_Inputs_Type
		Dim NX As Short 'DIMENSIONLESS
		Dim VOID_I As Double 'DIMENSIONLESS
		Dim DEN_I As Double 'g/cm^3
		Dim FLRT_I As Double 'gal/min-ft^2
		<VBFixedArray(ModelECM_NMAX)> Dim INDEX_IO() As Short 'DIMENSIONLESS
		<VBFixedArray(ModelECM_NMAX)> Dim XK_I() As Double '(umol/g)*(L/umol)^(1/n)
		<VBFixedArray(ModelECM_NMAX)> Dim XN_I() As Double 'DIMENSIONLESS
		<VBFixedArray(ModelECM_NMAX)> Dim C0_I() As Double 'ug/L
		<VBFixedArray(ModelECM_NMAX)> Dim XMW_I() As Double 'g/gmol
		
		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			'UPGRADE_WARNING: Lower bound of array INDEX_IO was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim INDEX_IO(ModelECM_NMAX)
			'UPGRADE_WARNING: Lower bound of array XK_I was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim XK_I(ModelECM_NMAX)
			'UPGRADE_WARNING: Lower bound of array XN_I was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim XN_I(ModelECM_NMAX)
			'UPGRADE_WARNING: Lower bound of array C0_I was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim C0_I(ModelECM_NMAX)
			'UPGRADE_WARNING: Lower bound of array XMW_I was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim XMW_I(ModelECM_NMAX)
		End Sub
	End Structure
	'UPGRADE_WARNING: Arrays in structure ModelECM_Inputs may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Dim ModelECM_Inputs As ModelECM_Inputs_Type
	Private Structure ModelECM_Outputs_Type
		Dim NX As Short 'DIMENSIONLESS
		<VBFixedArray(ModelECM_NMAX, ModelECM_NMAX)> Dim C_O(, ) As Double
		<VBFixedArray(ModelECM_NMAX, ModelECM_NMAX)> Dim DGY_O(, ) As Double
		<VBFixedArray(ModelECM_NMAX, ModelECM_NMAX)> Dim FCS_O(, ) As Double
		<VBFixedArray(ModelECM_NMAX)> Dim OATS_O() As Double
		<VBFixedArray(ModelECM_NMAX, ModelECM_NMAX)> Dim Q_O(, ) As Double
		<VBFixedArray(ModelECM_NMAX, ModelECM_NMAX)> Dim QAVE_O(, ) As Double
		<VBFixedArray(ModelECM_NMAX)> Dim SSTC_O() As Double
		<VBFixedArray(ModelECM_NMAX)> Dim VW_O() As Double
		<VBFixedArray(ModelECM_NMAX)> Dim ZZZ_O() As Double
		<VBFixedArray(ModelECM_NMAX)> Dim C0_O() As Double 'ug/L
		
		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			'UPGRADE_WARNING: Lower bound of array C_O was changed from 1,1 to 0,0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim C_O(ModelECM_NMAX, ModelECM_NMAX)
			'UPGRADE_WARNING: Lower bound of array DGY_O was changed from 1,1 to 0,0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim DGY_O(ModelECM_NMAX, ModelECM_NMAX)
			'UPGRADE_WARNING: Lower bound of array FCS_O was changed from 1,1 to 0,0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim FCS_O(ModelECM_NMAX, ModelECM_NMAX)
			'UPGRADE_WARNING: Lower bound of array OATS_O was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim OATS_O(ModelECM_NMAX)
			'UPGRADE_WARNING: Lower bound of array Q_O was changed from 1,1 to 0,0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim Q_O(ModelECM_NMAX, ModelECM_NMAX)
			'UPGRADE_WARNING: Lower bound of array QAVE_O was changed from 1,1 to 0,0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim QAVE_O(ModelECM_NMAX, ModelECM_NMAX)
			'UPGRADE_WARNING: Lower bound of array SSTC_O was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim SSTC_O(ModelECM_NMAX)
			'UPGRADE_WARNING: Lower bound of array VW_O was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim VW_O(ModelECM_NMAX)
			'UPGRADE_WARNING: Lower bound of array ZZZ_O was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim ZZZ_O(ModelECM_NMAX)
			'UPGRADE_WARNING: Lower bound of array C0_O was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim C0_O(ModelECM_NMAX)
		End Sub
	End Structure
	'UPGRADE_WARNING: Arrays in structure ModelECM_Outputs may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Dim ModelECM_Outputs As ModelECM_Outputs_Type
	
	'MISC VARIABLES (TIMER).
	Public ModelIO_Timer_TimeStart As String
	Public ModelIO_Timer_TimeEnd As String
	Public ModelIO_Timer_SummaryMsg As String
	
	
	
	
	Const ModelECM_declarations_end As Boolean = True
	
	
	Sub ModelECM_Go()
		Call ModelECM_WritePathFile()
		Call ModelECM_WriteMainFile()
		Call ModelECM_CallEXE()
		Call ModelECM_ProcessOutput()
		If (ModelIO_IsKeepTempFiles() = False) Then
			Call ModelECM_RemoveLinkFiles()
		End If
	End Sub
	
	
	Sub ModelECM_RemoveLinkFiles()
		Call KillFile_If_Exists(Exe_Path & "\" & ModelECM_IN_PathFile)
		Call KillFile_If_Exists(Exe_Path & "\" & ModelECM_IN_Main)
		Call KillFile_If_Exists(Exe_Path & "\" & ModelECM_OUT_SuccessFlag)
		Call KillFile_If_Exists(Exe_Path & "\" & ModelECM_OUT_Main)
	End Sub
	Sub ModelECM_CallEXE()
		Dim CmdLine As String
		Dim Test As String
		Call ChangeDir_Exes()
		'ChDir App.Path
		'frmMain.CommonDialog1.Filename = App.Path & "\*.*"
		'CmdLine = Exe_Path & "\" & ModelECM_ExeName
		'CmdLine = ModelECM_ExeName
		'ChDir App.Path & "\exes"
		'Test = Dir("*.*")
		'Test = Dir: Call Show_Message(Test)
		'Test = Dir: Call Show_Message(Test)
		'Test = Dir: Call Show_Message(Test)
		'Test = Dir: Call Show_Message(Test)
		'Test = Dir: Call Show_Message(Test)
		'Test = Dir: Call Show_Message(Test)
		'Call Show_Message(CurDir$)
		CmdLine = ModelECM_ExeName
		Call ModelIO_Timer_Start()
		Call FortranLink_ExecAndWaitForProcess(CmdLine)
		Call ModelIO_Timer_End()
		Call ChangeDir_Main()
	End Sub
	Sub ModelECM_ProcessOutput()
		Dim f As Short
		Dim fn_This As String
		Dim DummyStr1 As String
		Dim Flag_IMSL As Short
		'UPGRADE_WARNING: Arrays in structure MI may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim MI As ModelECM_Inputs_Type
		'UPGRADE_WARNING: Arrays in structure MO may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim MO As ModelECM_Outputs_Type
		Dim i As Short
		Dim j As Short
		Dim L As Short
		Dim EOFTESTMARKER As Double
		Dim MASSBAL_C0_e_Vf() As Double
		Dim MASSBAL_TERM_SUM() As Double
		Dim MASSBAL_PERCENT_ERR() As Double
		'Call debug_output("e1")
		'MAKE COPY OF INPUT DATA.
		'UPGRADE_WARNING: Couldn't resolve default property of object MI. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		MO.Initialize()   'Shang add
		MI = ModelECM_Inputs
		'READ SUCCESS FLAG OUTPUT FILE.
		f = FreeFile
		fn_This = Exe_Path & "\" & ModelECM_OUT_SuccessFlag
		If (Not FileExists(fn_This)) Then
			Call Show_Error("Unable to find output file: Calculations failed.")
			Exit Sub
		End If
		FileOpen(f, fn_This, OpenMode.Input)
		DummyStr1 = LineInput(f)
		Input(f, Flag_IMSL)
		FileClose(f)
		If (Flag_IMSL <> 0) Then
			Call Show_Error("The model calculations failed.")
			Exit Sub
		End If
		'Call debug_output("e2")
		'READ MAIN OUTPUT FILE.
		fn_This = Exe_Path & "\" & ModelECM_OUT_Main
		FileOpen(f, fn_This, OpenMode.Input)
		DummyStr1 = LineInput(f)
		Input(f, MO.NX)
		DummyStr1 = LineInput(f)
		For i = 1 To MO.NX
			Input(f, MI.INDEX_IO(i))
		Next i
		'Call debug_output("e2a")
		DummyStr1 = LineInput(f)
		For i = 1 To MO.NX
			For j = 1 To MO.NX
				Input(f, MO.C_O(i, j))
			Next j
		Next i
		DummyStr1 = LineInput(f)
		For i = 1 To MO.NX
			For j = 1 To MO.NX
				Input(f, MO.DGY_O(i, j))
			Next j
		Next i
		DummyStr1 = LineInput(f)
		For i = 1 To MO.NX
			For j = 1 To MO.NX
				Input(f, MO.FCS_O(i, j))
			Next j
		Next i
		DummyStr1 = LineInput(f)
		For i = 1 To MO.NX
			Input(f, MO.OATS_O(i))
		Next i
		DummyStr1 = LineInput(f)
		For i = 1 To MO.NX
			For j = 1 To MO.NX
				Input(f, MO.Q_O(i, j))
			Next j
		Next i
		'Call debug_output("e2b")
		DummyStr1 = LineInput(f)
		For i = 1 To MO.NX
			For j = 1 To MO.NX
				Input(f, MO.QAVE_O(i, j))
			Next j
		Next i
		DummyStr1 = LineInput(f)
		For i = 1 To MO.NX
			Input(f, MO.SSTC_O(i))
		Next i
		DummyStr1 = LineInput(f)
		For i = 1 To MO.NX
			Input(f, MO.VW_O(i))
		Next i
		DummyStr1 = LineInput(f)
		For i = 1 To MO.NX
			Input(f, MO.ZZZ_O(i))
		Next i
		'Call debug_output("e2c")
		DummyStr1 = LineInput(f)
		For i = 1 To MO.NX
			Input(f, MO.C0_O(i))
		Next i
		'Call debug_output("e2d")
		DummyStr1 = LineInput(f)
		Input(f, EOFTESTMARKER)
		If (False = ModelIO_DoNumberCheck(EOFTESTMARKER, ModelECM_EofTestMarker)) Then
			Call Show_Error("The model calculations failed: invalid file format (EOF marker).")
			Exit Sub
		End If
		FileClose(f)
		'UPGRADE_WARNING: Couldn't resolve default property of object ModelECM_Outputs. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ModelECM_Outputs = MO
		'UPGRADE_WARNING: Couldn't resolve default property of object ModelECM_Inputs. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ModelECM_Inputs = MI
		'Call debug_output("e3")
		'PERFORM MASS-BALANCE CALCULATION.
		'UPGRADE_WARNING: Lower bound of array MASSBAL_C0_e_Vf was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim MASSBAL_C0_e_Vf(MO.NX)
		'UPGRADE_WARNING: Lower bound of array MASSBAL_TERM_SUM was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim MASSBAL_TERM_SUM(MO.NX)
		'UPGRADE_WARNING: Lower bound of array MASSBAL_PERCENT_ERR was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim MASSBAL_PERCENT_ERR(MO.NX)
		'UPGRADE_WARNING: Couldn't resolve default property of object ModelECM_NMAX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Call ModelECM_MASS_BALANCE(MO.NX, MO.VW_O, MO.C_O, MO.Q_O, MI.VOID_I, MI.DEN_I, MI.FLRT_I, MO.C0_O, MASSBAL_C0_e_Vf, MASSBAL_TERM_SUM, MASSBAL_PERCENT_ERR)
		'------------------------------------------------------------------------------------------------------------------------
		'f = FreeFile
		'Open "c:\vb5ecm.txt" For Output As #f
		'Print #f, "MO.NX"
		'Print #f, MO.NX
		'Print #f, "MO.VW_O(i)"
		'For i = 1 To MO.NX
		'  Print #f, MO.VW_O(i)
		'Next i
		'Print #f, "MO.C_O(i,j)"
		'For i = 1 To MO.NX
		'  For j = 1 To MO.NX
		'    Print #f, MO.C_O(i, j)
		'  Next j
		'Next i
		'Print #f, "MO.Q_O(i,j)"
		'For i = 1 To MO.NX
		'  For j = 1 To MO.NX
		'    Print #f, MO.Q_O(i, j)
		'  Next j
		'Next i
		'Print #f, "MI.VOID_I"
		'Print #f, MI.VOID_I
		'Print #f, "MI.DEN_I"
		'Print #f, MI.DEN_I
		'Print #f, "MI.FLRT_I"
		'Print #f, MI.FLRT_I
		'Print #f, "MO.C0_O(i)"
		'For i = 1 To MI.NX
		'  Print #f, MO.C0_O(i)
		'Next i
		'Close #f
		'------------------------------------------------------------------------------------------------------------------------
		'TRANSFER OUTPUT DATA TO MORE PERMANENT MEMORY.
		For i = 1 To MO.NX
			Output_ECM(i).Index = MI.INDEX_IO(i)
			Output_ECM(i).Initialize()  'Shang add 
		Next i
		For L = 1 To MO.NX
			'J = Output_ECM(L).Index
			j = L
			For i = 1 To MO.NX
				If (i = 1) Then
					Output_ECM(j).C_Over_C0(i) = 1#
				Else
					Output_ECM(j).C_Over_C0(i) = MO.FCS_O(L, i)
				End If
				Output_ECM(j).Solid_Concentration(i) = MO.Q_O(L, i)
				Output_ECM(j).Liquid_Concentration(i) = MO.C_O(L, i)
			Next i
			Output_ECM(j).Bed_Volume_Fed = MO.OATS_O(L)
			Output_ECM(j).Dimensionless_Bed_Length = MO.ZZZ_O(L)
			Output_ECM(j).SS_Treatment_Capacity = MO.SSTC_O(L)
			Output_ECM(j).Wave_Velocity = MO.VW_O(L)
			Output_ECM(j).Carbon_Usage_Rate = MI.DEN_I * 1000# * 1000# / MO.OATS_O(L)
		Next L
		'Call debug_output("e4")

		Output_ECM_MASSBAL.Initialize() 'Shang add 
		For i = 1 To MO.NX
			Output_ECM_MASSBAL.MASSBAL_C0_e_Vf(i) = MASSBAL_C0_e_Vf(i)
			Output_ECM_MASSBAL.MASSBAL_TERM_SUM(i) = MASSBAL_TERM_SUM(i)
			Output_ECM_MASSBAL.MASSBAL_PERCENT_ERR(i) = MASSBAL_PERCENT_ERR(i)
		Next i
		frmMain.mnuResultsItem(2).Enabled = True
		Number_Component_ECM = MO.NX
		Call Show_Message("ECM Model Calculations Complete." & vbCrLf & vbCrLf & ModelIO_Timer_SummaryMsg)
		'Call debug_output("e5")
	End Sub
	Sub ModelECM_WriteMainFile()
		Dim f As Short
		Dim fn_This As String
		'UPGRADE_WARNING: Arrays in structure MI may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim MI As ModelECM_Inputs_Type
		Dim i As Short
		Dim j As Short
		'PREPARE INPUTS.
		MI.Initialize()             'Shang add
		MI.NX = Number_Component_ECM
		MI.VOID_I = 1# - Bed.Weight / (Bed.Diameter / 2#) ^ 2# / PI / Bed.length / Carbon.Density / 1000#
		MI.DEN_I = (1# - MI.VOID_I) * Carbon.Density
		MI.FLRT_I = (Bed.Flowrate / Bed.Diameter ^ 2# * 4# / PI) * (60# * 0.3048 ^ 2 * 1000# / 3.785)
		For i = 1 To MI.NX
			'UPGRADE_WARNING: Couldn't resolve default property of object Component_Index_ECM(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			j = Component_Index_ECM(i)
			MI.INDEX_IO(i) = j
			'CONVERT FROM (mg/g)*(L/mg)^(1/n) TO (umol/g)*(L/umol)^(1/n).
			MI.XK_I(i) = Component(j).Use_K * (1000# / Component(j).MW) ^ (1 - Component(j).Use_OneOverN)
			MI.XN_I(i) = Component(j).Use_OneOverN
			'CONVERT FROM mg/L TO ug/L.
			MI.C0_I(i) = Component(j).InitialConcentration * 1000#
			MI.XMW_I(i) = Component(j).MW
		Next i
		'WRITE INPUT FILE.
		f = FreeFile
		fn_This = Exe_Path & "\" & ModelECM_IN_Main
		'fn_This = App.Path & "\" & ModelECM_IN_Main
		FileOpen(f, fn_This, OpenMode.Output)
		Call WriteFortranInput(f, ModelECM_Version, "MODULE_VERSION")
		Call WriteFortranInput(f, MI.NX, "NX, number of components")
		Call WriteFortranInput(f, MI.VOID_I, "VOID_I, bed void fraction, dim'less")
		Call WriteFortranInput(f, MI.DEN_I, "DEN_I, bed density, g/cm^3")
		Call WriteFortranInput(f, MI.FLRT_I, "FLRT_I, superficial flow rate, gal/min-ft^2")
		For i = 1 To MI.NX
			Call WriteFortranInput(f, MI.INDEX_IO(i), "INDEX_IO(" & Trim(Str(i)) & "), component index")
			Call WriteFortranInput(f, MI.XK_I(i), "XK_I(" & Trim(Str(i)) & "), Freundlich K, (umol/g)*(L/umol)^(1/n)")
			Call WriteFortranInput(f, MI.XN_I(i), "XN_I(" & Trim(Str(i)) & "), Freundlich 1/n, dim'less")
			Call WriteFortranInput(f, MI.C0_I(i), "C0_I(" & Trim(Str(i)) & "), influent concentration, ug/L")
			Call WriteFortranInput(f, MI.XMW_I(i), "XMW_I(" & Trim(Str(i)) & "), molecular weight, g/gmol")
		Next i
		Call WriteFortranInput(f, ModelECM_EofTestMarker, "EOFTESTMARKER")
		FileClose(f)
		'STORE FOR LATER USE.
		'UPGRADE_WARNING: Couldn't resolve default property of object ModelECM_Inputs. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ModelECM_Inputs = MI
	End Sub
	Sub ModelECM_WritePathFile()
		Dim f As Short
		Dim fn_This As String
		Dim qq As String
		qq = Chr(34)
		f = FreeFile
		fn_This = Exe_Path & "\" & ModelECM_IN_PathFile
		'fn_This = App.Path & "\" & ModelECM_IN_PathFile
		FileOpen(f, fn_This, OpenMode.Output)
		PrintLine(f, qq & ModelECM_IN_Main & qq)
		PrintLine(f, qq & ModelECM_OUT_SuccessFlag & qq)
		PrintLine(f, qq & ModelECM_OUT_Main & qq)
		FileClose(f)
	End Sub


	'C********************************************************************
	'C
	'C MASS_BALANCE
	'C
	'C Description:  This routine will do the mass balance on the output
	'C               from the ECM program for each component and tell
	'C               the percent error on the mass balance.
	'C
	'C Input Variables:
	'C    N =        Number of Components
	'C    VW =       Array of Wave Velocities for each zone (1 to N)
	'C (cm / s)
	'C    C =        Array of Liquid Phase Concentrations for Each
	'C               Component in Each Zone : C(Component,Zone) -
	'C               N x N two-dimensional array (ug/L)
	'C    Q =        Array of Gas Phase Concentrations for Each
	'C               Component in Each Zone : q(Component,Zone) -
	'C               N x N two-dimensional array (ug/g)
	'C    EBED =     Void Fraction of Bed (-)
	'C    DEN =      Bulk Density of Adsorbent (g/cm3)
	'C FLRT = Flowrate(gpm / ft2)
	'C    COK =      Array of Liquid Phase Influent Concentrations
	'C               (1 to N) (ug/L)
	'C
	'C Output Variables:
	'C    C0_e_Vf =  Left-hand side of mass balance (ug/cm2/s). Array
	'C               from 1 to N.
	'C    TERM_SUM = Right-hand side of mass balance (ug/cm2/s).
	'C               Array from 1 to N.
	'C    PERCENT_ERR = Percent difference between C0_e_Vf and
	'C                  TERM_SUM (%). Array from 1 to N.
	'C
	'C Variables internal to this Subroutine:
	'C    VF =       Interstitial fluid velocity (L/cm2/s)
	'C
	'C********************************************************************
	Sub ModelECM_MASS_BALANCE(ByRef N As Short, ByRef VW() As Double, ByRef C(,) As Double, ByRef Q(,) As Double, ByRef EBED As Double, ByRef DEN As Double, ByRef FLRT As Double, ByRef COK() As Double, ByRef OUTPUT_C0_e_Vf() As Double, ByRef OUTPUT_TERM_SUM() As Double, ByRef OUTPUT_PERCENT_ERR() As Double)
		Dim VF As Double
		Dim i As Short
		Dim j As Short
		Dim TERM As Double
		VF = FLRT * 1000.0# / 60.0# / (30.48 ^ 2.0#) / 264.17 / EBED
		For i = 1 To N
			OUTPUT_C0_e_Vf(i) = COK(i) * VF * EBED
		Next i
		For i = 1 To N
			OUTPUT_TERM_SUM(i) = 0#
			For j = 1 To N
				If (j = 1) Then
					TERM = VW(j) * (Q(i, j) * DEN + C(i, j) * EBED / 1000.0#)
				Else
					TERM = (VW(j) - VW(j - 1)) * (Q(i, j) * DEN + C(i, j) * EBED / 1000.0#)
				End If
				OUTPUT_TERM_SUM(i) = OUTPUT_TERM_SUM(i) + TERM
			Next j
			OUTPUT_PERCENT_ERR(i) = System.Math.Abs((OUTPUT_C0_e_Vf(i) - OUTPUT_TERM_SUM(i)) / OUTPUT_C0_e_Vf(i)) * 100.0#
		Next i
		'
		'      SUBROUTINE ECM_MASSBAL (N,VW,C,Q,EBED,DEN,FLRT,COK,C0_e_Vf,
		'     &                         TERM_SUM,PERCENT_ERR,VF)
		'
		'      IMPLICIT NONE
		'      INTEGER N,I,J,K
		'      DOUBLE PRECISION VW(N),C(N,N),Q(N,N),EBED,DEN,FLRT,COK(N)
		'      DOUBLE PRECISION TERM,C0_e_Vf(N),TERM_SUM(N),
		'     &                 PERCENT_ERR(N),VF
		'
		'      VF = FLRT * 1000.0D0 / 60.0D0 / (30.48D0**2) / 264.17D0 / EBED
		'
		'      DO 10, I = 1,N
		'         C0_e_Vf(i) = COK(i) * VF * EBED
		'10    CONTINUE
		'
		'C**** Note:  I = Number of Component, J = Number of Zone
		'
		'      DO 20, I = 1,N
		'         TERM_SUM(i) = 0#
		'         DO 30, J = 1,N
		'            IF (J.EQ.1) THEN
		'               TERM = VW(j) * (Q(i, j) * DEN + C(i, j) * EBED / 1000#)
		'            Else
		'               TERM = (VW(J)-VW(J-1)) *
		'     &                (Q(I,J)*DEN+C(I,J)*EBED/1000.0D0)
		'            End If
		'            TERM_SUM(i) = TERM_SUM(i) + TERM
		'30       CONTINUE
		'         PERCENT_ERR(i) = DABS(((C0_e_Vf(i) - TERM_SUM(i)) / C0_e_Vf(i)))
		'     &                      * 100.0D0
		'20    CONTINUE
		'
		'      End
		'
		'C********************************************************************
		'
	End Sub










	'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	'//////////////////     THE FOLLOWING CODE APPLIES TO ALL MODELS, NOT JUST THE ECM.     /////////////////////////////////////
	'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	Sub AllModels_Verify_Selected_Components(ByRef Model As Short)
		Dim i As Short
		Select Case Model
			Case MODELTYPE_PSDM
				Number_Component_PFPSDM = 0
				For i = 1 To Number_Component
					If (frmMain.lstComponents.GetSelected(i - 1)) Then
						Number_Component_PFPSDM = Number_Component_PFPSDM + 1
						If Number_Component_PFPSDM > Number_Compo_Max_PFPSDM Then
							Call Show_Error("You selected too many components for the PSDM!")
							Number_Component_PFPSDM = 0
							Exit Sub
						End If
						Component_Index_PFPSDM(Number_Component_PFPSDM) = i
					End If
				Next i
				If Number_Component_PFPSDM = 0 Then
					Call Show_Error("You did not select any component for the PSDM!")
				End If
			Case MODELTYPE_CPHSDM
				Number_Component_CPM = 0
				For i = 1 To Number_Component
					If (frmMain.lstComponents.GetSelected(i - 1)) Then
						Number_Component_CPM = Number_Component_CPM + 1
						If Number_Component_CPM > Number_Compo_Max_CPM Then
							Call Show_Error("You selected too many components for the CPHSDM!")
							Number_Component_CPM = 0
							Exit Sub
						End If
						Component_Index_CPM = i
					End If
				Next i
				If Number_Component_CPM = 0 Then
					Call Show_Error("You did not select any component for the CPHSDM!")
				End If
			Case MODELTYPE_ECM
				Number_Component_ECM = 0
				For i = 1 To Number_Component
					If (frmMain.lstComponents.GetSelected(i - 1)) Then
						Number_Component_ECM = Number_Component_ECM + 1
						If (Number_Component_ECM > Number_Compo_Max_ECM) Then
							Call Show_Error("You selected too many components for the ECM!")
							Number_Component_ECM = 0
							Exit Sub
						End If
						'UPGRADE_WARNING: Couldn't resolve default property of object Component_Index_ECM(Number_Component_ECM). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Component_Index_ECM(Number_Component_ECM) = i
					End If
				Next i
				If (Number_Component_ECM = 0) Then
					Call Show_Error("You did not select any component for the ECM!")
				End If
		End Select
	End Sub
	
	
	Function ModelIO_DoNumberCheck(ByRef N1 As Double, ByRef N2 As Double) As Boolean
		If (System.Math.Abs(N1 + 0.000001) / N2 - 1#) <= 0.001 Then
			'NUMBERS ARE CLOSE TO IDENTICAL.
			ModelIO_DoNumberCheck = True
		Else
			'NUMBERS ARE _NOT_ CLOSE TO IDENTICAL.
			ModelIO_DoNumberCheck = False
		End If
	End Function
	
	
	Sub ModelIO_Timer_Start()
		ModelIO_Timer_TimeStart = CStr(Now)
	End Sub
	Sub ModelIO_Timer_End()
		Dim Elapsed_Min As Double
		Dim TotalTimeStr As String
		ModelIO_Timer_TimeEnd = CStr(Now)
		'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
		Elapsed_Min = DateDiff(Microsoft.VisualBasic.DateInterval.Second, CDate(ModelIO_Timer_TimeStart), CDate(ModelIO_Timer_TimeEnd)) / 60#
		TotalTimeStr = VB6.Format(Elapsed_Min, "0.00")
		ModelIO_Timer_SummaryMsg = "Calculation Started:    " & ModelIO_Timer_TimeStart & vbCrLf & "Calculation Ended:    " & ModelIO_Timer_TimeEnd & vbCrLf & vbCrLf & "Total Calculation Time = " & TotalTimeStr & " Minutes"
	End Sub
	
	
	Function ModelIO_IsKeepTempFiles() As Boolean
		ModelIO_IsKeepTempFiles = frmMain.mnuMTUItem(40).Checked
	End Function
End Module