Option Strict Off
Option Explicit On
Module ModelIPE
	
	Const ModelIPE_IN_PathFile As String = "IPES1.IN"
	Const ModelIPE_IN_Main As String = "IPES2.IN"
	Const ModelIPE_OUT_SuccessFlag As String = "IPES1.OUT"
	Const ModelIPE_OUT_Main As String = "IPES2.OUT"
	
	Const ModelIPE_Version As Double = 1#
	''''Const ModelIPE_ExeName = "IPES3.EXE"
	Const ModelIPE_ExeName As String = "IPES4.EXE"
	Const ModelIPE_EofTestMarker As Double = 123456#
	
	Public Const MODULECODE_ADLIQ As Short = 1
	Public Const MODULECODE_HOFMAN As Short = 5
	Public Const MODULECODE_SPEQ As Short = 4
	
	Dim SHARED_MODULECODE As Short
	Dim SHARED_NL As Short
	Dim SHARED_OMAG As Double
	
	'
	'///////////// ADLIQ INPUTS / OUTPUTS ////////////////////////////////////////////////////////////////////////////////////////////////
	'
	Private Structure ModelIPE_ADLIQ_Inputs_Type
		Dim BB As Double 'POLANYI PARAMETER.
		Dim W0 As Double 'POLANYI PARAMETER.
		Dim GM As Double 'POLANYI PARAMETER (dimless).
		Dim CBULK As Double 'BULK CONCENTRATION (ug/L).
		Dim ORGDEN As Double 'ORGANIC DENSITY (g/cm^3).
		Dim TT As Double 'TEMPERATURE (degrees Kelvin).
		Dim FWT As Double 'MOLECULAR WEIGHT (g/gmol).
		Dim SOLUB As Double 'AQUEOUS SOLUBILITY (mg/L).
		Dim NL As Short 'NUMBER OF REGRESSION POINTS (dimless).
		Dim OMAG As Double 'ORDER OF MAGNITUDE OF REGRESSION (dimless).
		Dim VOLM_NBP As Double 'MOLAR VOLUME AT NORMAL BOILING POINT (cm^3/gmol).  (new as of 1999-May-14)
	End Structure
	Dim ModelIPE_ADLIQ_Inputs As ModelIPE_ADLIQ_Inputs_Type
	Private Structure ModelIPE_ADLIQ_Outputs_Type
		Dim CSAV As Double 'AVERAGE BULK CONC (ug/L).
		Dim QSAV As Double 'POLANYI ADSORPTION CAPACITY (ug/g).
		Dim XK1 As Double 'FREUNDLICH K (ug/g)*(L/ug)^(1/n).
		Dim XK2 As Double 'FREUNDLICH K (umol/g)*(L/umol)^(1/n).
		Dim XNF As Double 'FREUNDLICH 1/N (dimless).
		Dim CBEG As Double 'CORRELATION LOWER BOUND (ug/L).
		Dim CEND As Double 'CORRELATION UPPER BOUND (ug/L).
		Dim RSQD As Double 'REGRESSION R-SQUARED (dimless).
		Dim RMSE As Double 'ROOT MEAN SQUARE ERROR (dimless?).
		'<VBFixedArray(30)> Dim ErrMat() As Short 'ERROR MATRIX.
		Dim ErrMat() As Short
		Dim ALERR As Short 'HAS ANY ERROR/WARNING OCCURRED?

		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			'UPGRADE_WARNING: Lower bound of array ErrMat was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim ErrMat(30)
		End Sub
	End Structure
	'UPGRADE_WARNING: Arrays in structure ModelIPE_ADLIQ_Outputs may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Dim ModelIPE_ADLIQ_Outputs As ModelIPE_ADLIQ_Outputs_Type


	'
	'///////////// HOFMAN INPUTS / OUTPUTS ////////////////////////////////////////////////////////////////////////////////////////////////
	'
	Private Structure ModelIPE_HOFMAN_Inputs_Type
		'<VBFixedArray(11)> Dim IN_DATA() As Double 'VARIOUS PARAMETERS.
		Dim IN_DATA() As Double
		Dim NL As Short 'NUMBER OF REGRESSION POINTS (dimless).
		Dim VOLM_NBP As Double 'MOLAR VOLUME AT NORMAL BOILING POINT (cm^3/gmol).  (new as of 1999-May-14)

		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			'UPGRADE_WARNING: Lower bound of array IN_DATA was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim IN_DATA(11)
		End Sub
	End Structure
	'UPGRADE_WARNING: Arrays in structure ModelIPE_HOFMAN_Inputs may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Dim ModelIPE_HOFMAN_Inputs As ModelIPE_HOFMAN_Inputs_Type
	Private Structure ModelIPE_HOFMAN_Outputs_Type
		Dim CSAV As Double 'AVERAGE BULK CONC (ug/L).
		Dim QSAV As Double 'POLANYI ADSORPTION CAPACITY (ug/g).
		Dim XK1 As Double 'FREUNDLICH K (ug/g)*(L/ug)^(1/n).
		Dim XK2 As Double 'FREUNDLICH K (umol/g)*(L/umol)^(1/n).
		Dim XNF As Double 'FREUNDLICH 1/N (dimless).
		Dim CBEG As Double 'CORRELATION LOWER BOUND (ug/L).
		Dim CEND As Double 'CORRELATION UPPER BOUND (ug/L).
		Dim RSQD As Double 'REGRESSION R-SQUARED (dimless).
		Dim RMSE As Double 'ROOT MEAN SQUARE ERROR (dimless?).
		'<VBFixedArray(30)> Dim ErrMat() As Short 'ERROR MATRIX.
		Dim ErrMat() As Short
		Dim HOERR As Short 'HAS ANY ERROR/WARNING OCCURRED?

		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			'UPGRADE_WARNING: Lower bound of array ErrMat was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim ErrMat(30)
		End Sub
	End Structure
	'UPGRADE_WARNING: Arrays in structure ModelIPE_HOFMAN_Outputs may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Dim ModelIPE_HOFMAN_Outputs As ModelIPE_HOFMAN_Outputs_Type


	'
	'///////////// SPEQ INPUTS / OUTPUTS ////////////////////////////////////////////////////////////////////////////////////////////////
	'
	Private Structure ModelIPE_SPEQ_Inputs_Type
		'<VBFixedArray(10)> Dim IN_DATA() As Double 'VARIOUS PARAMETERS.
		Dim IN_DATA() As Double
		Dim NL As Short 'NUMBER OF REGRESSION POINTS (dimless).
		Dim XERR As Short '??? FORCING TO ZERO SEEMS ACCEPTABLE.

		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			'UPGRADE_WARNING: Lower bound of array IN_DATA was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim IN_DATA(10)
		End Sub
	End Structure
	'UPGRADE_WARNING: Arrays in structure ModelIPE_SPEQ_Inputs may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Dim ModelIPE_SPEQ_Inputs As ModelIPE_SPEQ_Inputs_Type
	Private Structure ModelIPE_SPEQ_Outputs_Type
		Dim CSAV As Double 'AVERAGE BULK CONC (ug/L).
		Dim QSAV As Double 'POLANYI ADSORPTION CAPACITY (ug/g).
		Dim XK1 As Double 'FREUNDLICH K (ug/g)*(L/ug)^(1/n).
		Dim XK2 As Double 'FREUNDLICH K (umol/g)*(L/umol)^(1/n).
		Dim XNF As Double 'FREUNDLICH 1/N (dimless).
		Dim CBEG As Double 'CORRELATION LOWER BOUND (ug/L).
		Dim CEND As Double 'CORRELATION UPPER BOUND (ug/L).
		'<VBFixedArray(30)> Dim ErrMat() As Short 'ERROR MATRIX.
		Dim ErrMat() As Short
		Dim SQERR As Short 'HAS ANY ERROR/WARNING OCCURRED?

		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			'UPGRADE_WARNING: Lower bound of array ErrMat was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim ErrMat(30)
		End Sub
	End Structure
	'UPGRADE_WARNING: Arrays in structure ModelIPE_SPEQ_Outputs may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Dim ModelIPE_SPEQ_Outputs As ModelIPE_SPEQ_Outputs_Type





	Const ModelIPE_declarations_end As Boolean = True


	Sub ModelIPE_Go(ByRef WhichModule As Short, ByRef INPUT_NL As Short, ByRef INPUT_OMAG As Double, ByRef Raise_Dirty_Flag As Boolean)
		Dim Found As Boolean
		SHARED_NL = INPUT_NL
		SHARED_OMAG = INPUT_OMAG
		SHARED_MODULECODE = WhichModule
		Found = False
		Raise_Dirty_Flag = False
		Select Case WhichModule
			Case MODULECODE_ADLIQ '3-PARAMETER POLANYI (LIQUID).
				Call ModelIPE_ADLIQ_Go(Raise_Dirty_Flag)
			Case MODULECODE_SPEQ 'D-R SPREADING PRESSURE (GAS).
				Call ModelIPE_SPEQ_Go(Raise_Dirty_Flag)
			Case MODULECODE_HOFMAN 'D-R UNIFORM ADSORBATE (LIQUID).
				Call ModelIPE_HOFMAN_Go(Raise_Dirty_Flag)
			Case Else
				Call Show_Error("Invalid IPE module type " & Trim(Str(WhichModule)) & ".  Select a different " & "IPE module type.")
				Exit Sub
		End Select
	End Sub




	'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	'////////////////    ADLIQ MODULE    //////////////////////////////////////////////////////////////////////////////////
	'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	Sub ModelIPE_ADLIQ_ProcessOutput(ByRef Raise_Dirty_Flag As Boolean)
		Dim f As Short
		Dim fn_This As String
		Dim DummyStr1 As String
		Dim DummyVal1 As Short
		Dim i As Short
		Dim Flag_IPE As Short
		Dim MI As ModelIPE_ADLIQ_Inputs_Type
		'UPGRADE_WARNING: Arrays in structure MO may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim MO As ModelIPE_ADLIQ_Outputs_Type
		Dim OkayToUse As Boolean
		Dim EOFTESTMARKER As Double
		'READ SUCCESS FLAG OUTPUT FILE.
		MO.Initialize()   'Shang added
		f = FreeFile()
		fn_This = Exe_Path & "\" & ModelIPE_OUT_SuccessFlag
		If (Not FileExists(fn_This)) Then
			Call Show_Error("Unable to find output file: Calculations failed.")
			Exit Sub
		End If
		FileOpen(f, fn_This, OpenMode.Input)
		DummyStr1 = LineInput(f)
		Input(f, Flag_IPE)
		FileClose(f)
		'READ MAIN OUTPUT FILE.
		fn_This = Exe_Path & "\" & ModelIPE_OUT_Main
		FileOpen(f, fn_This, OpenMode.Input)
		DummyStr1 = LineInput(f) : Input(f, MO.CSAV)
		DummyStr1 = LineInput(f) : Input(f, MO.QSAV)
		DummyStr1 = LineInput(f) : Input(f, MO.XK1)
		DummyStr1 = LineInput(f) : Input(f, MO.XK2)
		DummyStr1 = LineInput(f) : Input(f, MO.XNF)
		DummyStr1 = LineInput(f) : Input(f, MO.CBEG)
		DummyStr1 = LineInput(f) : Input(f, MO.CEND)
		DummyStr1 = LineInput(f) : Input(f, MO.RSQD)
		DummyStr1 = LineInput(f) : Input(f, MO.RMSE)
		DummyStr1 = LineInput(f)
		For i = 1 To 30
			Input(f, MO.ErrMat(i))
		Next i
		DummyStr1 = LineInput(f)
		Input(f, MO.ALERR) 'COPY OF ALERR / Flag_IPE.
		DummyStr1 = LineInput(f)
		Input(f, EOFTESTMARKER)
		If (False = ModelIO_DoNumberCheck(EOFTESTMARKER, ModelIPE_EofTestMarker)) Then
			Call Show_Error("The model calculations failed: invalid file format (EOF marker).")
			Exit Sub
		End If
		FileClose(f)
		'UPGRADE_WARNING: Couldn't resolve default property of object ModelIPE_ADLIQ_Outputs. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ModelIPE_ADLIQ_Outputs = MO
		'UPGRADE_WARNING: Couldn't resolve default property of object MI. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		MI = ModelIPE_ADLIQ_Inputs
		'DISPLAY WARNINGS/ERRORS IF NECESSARY.
		OkayToUse = AllIPEModels_ErrorCheck(MO.ALERR, MO.ErrMat)
		If (Not OkayToUse) Then Exit Sub
		'PROCESS IPE RESULTS.
		IPES_Data.Input_Renamed.BB = MI.BB
		IPES_Data.Input_Renamed.W0 = MI.W0
		IPES_Data.Input_Renamed.GM = MI.GM
		'Conversion from ug/g to mg/g
		IPES_Data.Output.QSAV = MO.QSAV / 1000.0#
		'Conversion from ug/g to mg/g
		IPES_Data.Output.CSAV = MO.CSAV / 1000.0#
		IPES_Data.Output.CBEG = MO.CBEG / 1000.0#
		IPES_Data.Output.CEND = MO.CEND / 1000.0#
		'Conversion from (ug/g)x(l/ug)^(1/n) to (mg/g)x(l/mg)^(1/n)
		IPES_Data.Output.XK1 = MO.XK1 * (1000.0#) ^ (MO.XNF - 1)
		'Conversion from (umol/g)x(l/umol)^(1/n) to (mmol/g)x(l/mmol)^(1/n)
		IPES_Data.Output.XK2 = MO.XK2 * (1000.0#) ^ (MO.XNF - 1)
		IPES_Data.Output.XN = MO.XNF
		IPES_Data.Output.RSQD = MO.RSQD
		IPES_Data.Output.RMSE = MO.RMSE
		'TRANSFER K AND 1/n.
		Component(0).IPESResult_K = IPES_Data.Output.XK1
		Component(0).IPESResult_OneOverN = IPES_Data.Output.XN
		'RAISE DIRTY FLAG.
		Raise_Dirty_Flag = True
		'DISPLAY RESULTS.
		Call frmModelIPEResults.frmModelIPEResults_Run(SHARED_MODULECODE)
	End Sub
	Sub ModelIPE_ADLIQ_WriteMainFile()
		Dim MI As ModelIPE_ADLIQ_Inputs_Type
		Dim f As Short
		Dim fn_This As String
		'
		' PREPARE INPUTS.
		' NOTE: IT IS ASSUMED THAT Component(0) CONTAINS THE
		' CHEMICAL PROPERTIES OF INTEREST.
		'
		MI.BB = Carbon.BB
		MI.W0 = Carbon.W0
		MI.GM = Carbon.PolanyiExponent
		MI.CBULK = Component(0).InitialConcentration * 1000.0#
		MI.ORGDEN = Component(0).Liquid_Density
		MI.TT = Bed.Temperature + 273.15
		MI.FWT = Component(0).MW
		MI.SOLUB = Component(0).Aqueous_Solubility
		MI.NL = SHARED_NL
		MI.OMAG = SHARED_OMAG
		MI.VOLM_NBP = Component(0).MolarVolume
		'
		' WRITE INPUT FILE.
		'
		f = FreeFile()
		fn_This = Exe_Path & "\" & ModelIPE_IN_Main
		FileOpen(f, fn_This, OpenMode.Output)
		Call WriteFortranInput(f, ModelIPE_Version, "MODULE_VERSION")
		Call WriteFortranInput(f, MI.BB, "BB, Polanyi parameter")
		Call WriteFortranInput(f, MI.W0, "W0, Polanyi parameter")
		Call WriteFortranInput(f, MI.GM, "GM, Polanyi exponent")
		Call WriteFortranInput(f, MI.CBULK, "CBULK, bulk concentration, ug/L")
		Call WriteFortranInput(f, MI.ORGDEN, "ORGDEN, organic density, g/cm^3")
		Call WriteFortranInput(f, MI.TT, "TT, temperature, degK")
		Call WriteFortranInput(f, MI.FWT, "FWT, molecular weight, g/gmol")
		Call WriteFortranInput(f, MI.SOLUB, "SOLUB, aqueous solubility, mg/L")
		Call WriteFortranInput(f, MI.NL, "NL, number of regression points, dimless")
		Call WriteFortranInput(f, MI.OMAG, "OMAG, order of magnitude of regression, dimless")
		Call WriteFortranInput(f, MI.VOLM_NBP, "VOLM_NBP, molar volume at the normal boiling point, cm^3/gmol")
		Call WriteFortranInput(f, ModelIPE_EofTestMarker, "EOFTESTMARKER")
		FileClose(f)
		'
		' STORE FOR LATER USE.
		'
		'UPGRADE_WARNING: Couldn't resolve default property of object ModelIPE_ADLIQ_Inputs. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ModelIPE_ADLIQ_Inputs = MI
	End Sub
	Sub ModelIPE_ADLIQ_Go(ByRef Raise_Dirty_Flag As Boolean)
		Call ModelIPE_WritePathFile(MODULECODE_ADLIQ)
		Call ModelIPE_ADLIQ_WriteMainFile()
		Call ModelIPE_CallEXE()
		Call ModelIPE_ADLIQ_ProcessOutput(Raise_Dirty_Flag)
		If (ModelIO_IsKeepTempFiles() = False) Then
			Call ModelIPE_RemoveLinkFiles()
		End If
	End Sub


	'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	'////////////////    HOFMAN MODULE    /////////////////////////////////////////////////////////////////////////////////
	'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	Sub ModelIPE_HOFMAN_ProcessOutput(ByRef Raise_Dirty_Flag As Boolean)
		Dim f As Short
		Dim fn_This As String
		Dim DummyStr1 As String
		Dim DummyVal1 As Short
		Dim i As Short
		Dim Flag_IPE As Short
		'UPGRADE_WARNING: Arrays in structure MI may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim MI As ModelIPE_HOFMAN_Inputs_Type
		'UPGRADE_WARNING: Arrays in structure MO may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim MO As ModelIPE_HOFMAN_Outputs_Type
		Dim OkayToUse As Boolean
		Dim EOFTESTMARKER As Double
		'READ SUCCESS FLAG OUTPUT FILE.
		f = FreeFile()
		fn_This = Exe_Path & "\" & ModelIPE_OUT_SuccessFlag
		If (Not FileExists(fn_This)) Then
			Call Show_Error("Unable to find output file: Calculations failed.")
			Exit Sub
		End If
		FileOpen(f, fn_This, OpenMode.Input)
		DummyStr1 = LineInput(f)
		Input(f, Flag_IPE)
		FileClose(f)
		'READ MAIN OUTPUT FILE.
		fn_This = Exe_Path & "\" & ModelIPE_OUT_Main
		FileOpen(f, fn_This, OpenMode.Input)
		DummyStr1 = LineInput(f) : Input(f, MO.CSAV)
		DummyStr1 = LineInput(f) : Input(f, MO.QSAV)
		DummyStr1 = LineInput(f) : Input(f, MO.XK1)
		DummyStr1 = LineInput(f) : Input(f, MO.XK2)
		DummyStr1 = LineInput(f) : Input(f, MO.XNF)
		DummyStr1 = LineInput(f) : Input(f, MO.CBEG)
		DummyStr1 = LineInput(f) : Input(f, MO.CEND)
		DummyStr1 = LineInput(f) : Input(f, MO.RSQD)
		DummyStr1 = LineInput(f) : Input(f, MO.RMSE)
		DummyStr1 = LineInput(f)
		For i = 1 To 30
			Input(f, MO.ErrMat(i))
		Next i
		DummyStr1 = LineInput(f)
		Input(f, MO.HOERR) 'COPY OF HOERR / Flag_IPE.
		DummyStr1 = LineInput(f)
		Input(f, EOFTESTMARKER)
		If (False = ModelIO_DoNumberCheck(EOFTESTMARKER, ModelIPE_EofTestMarker)) Then
			Call Show_Error("The model calculations failed: invalid file format (EOF marker).")
			Exit Sub
		End If
		FileClose(f)
		'UPGRADE_WARNING: Couldn't resolve default property of object ModelIPE_HOFMAN_Outputs. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ModelIPE_HOFMAN_Outputs = MO
		'UPGRADE_WARNING: Couldn't resolve default property of object MI. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		MI = ModelIPE_HOFMAN_Inputs
		'DISPLAY WARNINGS/ERRORS IF NECESSARY.
		OkayToUse = AllIPEModels_ErrorCheck(MO.HOERR, MO.ErrMat)
		If (Not OkayToUse) Then Exit Sub
		'PROCESS IPE RESULTS.
		IPES_Data.Input_Renamed.BB = MI.IN_DATA(1)
		IPES_Data.Input_Renamed.W0 = MI.IN_DATA(2)
		IPES_Data.Input_Renamed.GM = MI.IN_DATA(9)
		'Conversion from ug/g to mg/g
		IPES_Data.Output.QSAV = MO.QSAV / 1000.0#
		'Conversion from ug/g to mg/g
		IPES_Data.Output.CSAV = MO.CSAV / 1000.0#
		IPES_Data.Output.CBEG = MO.CBEG / 1000.0#
		IPES_Data.Output.CEND = MO.CEND / 1000.0#
		'Conversion from (ug/g)x(l/ug)^(1/n) to (mg/g)x(l/mg)^(1/n)
		IPES_Data.Output.XK1 = MO.XK1 * (1000.0#) ^ (MO.XNF - 1)
		'Conversion from (umol/g)x(l/umol)^(1/n) to (mmol/g)x(l/mmol)^(1/n)
		IPES_Data.Output.XK2 = MO.XK2 * (1000.0#) ^ (MO.XNF - 1)
		IPES_Data.Output.XN = MO.XNF
		IPES_Data.Output.RSQD = MO.RSQD
		IPES_Data.Output.RMSE = MO.RMSE
		'TRANSFER K AND 1/n.
		Component(0).IPESResult_K = IPES_Data.Output.XK1
		Component(0).IPESResult_OneOverN = IPES_Data.Output.XN
		'RAISE DIRTY FLAG.
		Raise_Dirty_Flag = True
		'DISPLAY RESULTS.
		Call frmModelIPEResults.frmModelIPEResults_Run(SHARED_MODULECODE)
	End Sub
	Sub ModelIPE_HOFMAN_WriteMainFile()
		'UPGRADE_WARNING: Arrays in structure MI may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim MI As ModelIPE_HOFMAN_Inputs_Type
		Dim f As Short
		Dim fn_This As String
		'
		' PREPARE INPUTS.
		' NOTE: IT IS ASSUMED THAT Component(0) CONTAINS THE
		' CHEMICAL PROPERTIES OF INTEREST.
		'
		MI.Initialize()

		MI.IN_DATA(1) = Carbon.BB
		MI.IN_DATA(2) = Carbon.W0
		MI.IN_DATA(3) = Bed.Temperature + 273.15
		MI.IN_DATA(4) = Component(0).InitialConcentration * 1000.0#
		MI.IN_DATA(5) = Component(0).Liquid_Density
		MI.IN_DATA(6) = Component(0).MW
		MI.IN_DATA(7) = Component(0).Vapor_Pressure / 101325 * 760 'UNITS: mmHg.
		MI.IN_DATA(8) = Component(0).Aqueous_Solubility
		MI.IN_DATA(9) = Component(0).Refractive_Index
		MI.IN_DATA(10) = Carbon.PolanyiExponent
		MI.IN_DATA(11) = SHARED_OMAG
		MI.NL = SHARED_NL
		MI.VOLM_NBP = Component(0).MolarVolume
		'
		' WRITE INPUT FILE.
		'
		f = FreeFile()
		fn_This = Exe_Path & "\" & ModelIPE_IN_Main
		FileOpen(f, fn_This, OpenMode.Output)
		Call WriteFortranInput(f, ModelIPE_Version, "MODULE_VERSION")
		Call WriteFortranInput(f, MI.IN_DATA(1), "IN_DATA(1), Polanyi parameter BB")
		Call WriteFortranInput(f, MI.IN_DATA(2), "IN_DATA(2), Polanyi parameter W0")
		Call WriteFortranInput(f, MI.IN_DATA(3), "IN_DATA(3), temperature, degK")
		Call WriteFortranInput(f, MI.IN_DATA(4), "IN_DATA(4), bulk concentration, ug/L")
		Call WriteFortranInput(f, MI.IN_DATA(5), "IN_DATA(5), organic density, g/cm^3")
		Call WriteFortranInput(f, MI.IN_DATA(6), "IN_DATA(6), molecular weight, g/gmol")
		Call WriteFortranInput(f, MI.IN_DATA(7), "IN_DATA(7), vapor pressure, mmHg")
		Call WriteFortranInput(f, MI.IN_DATA(8), "IN_DATA(8), aqueous solubility, mg/L")
		Call WriteFortranInput(f, MI.IN_DATA(9), "IN_DATA(9), refractive index, dimless")
		Call WriteFortranInput(f, MI.IN_DATA(10), "IN_DATA(10), Polanyi exponent GM, dimless")
		Call WriteFortranInput(f, MI.IN_DATA(11), "IN_DATA(11), order of magnitude of regression, dimless")
		Call WriteFortranInput(f, MI.NL, "NL, number of regression points, dimless")
		Call WriteFortranInput(f, MI.VOLM_NBP, "VOLM_NBP, molar volume at the normal boiling point, cm^3/gmol")
		Call WriteFortranInput(f, ModelIPE_EofTestMarker, "EOFTESTMARKER")
		FileClose(f)
		'
		' STORE FOR LATER USE.
		'
		'UPGRADE_WARNING: Couldn't resolve default property of object ModelIPE_HOFMAN_Inputs. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ModelIPE_HOFMAN_Inputs = MI
	End Sub
	Sub ModelIPE_HOFMAN_Go(ByRef Raise_Dirty_Flag As Boolean)
		Call ModelIPE_WritePathFile(MODULECODE_HOFMAN)
		Call ModelIPE_HOFMAN_WriteMainFile()
		Call ModelIPE_CallEXE()
		Call ModelIPE_HOFMAN_ProcessOutput(Raise_Dirty_Flag)
		If (ModelIO_IsKeepTempFiles() = False) Then
			Call ModelIPE_RemoveLinkFiles()
		End If
	End Sub


	'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	'////////////////    SPEQ MODULE    ///////////////////////////////////////////////////////////////////////////////////
	'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	Sub ModelIPE_SPEQ_ProcessOutput(ByRef Raise_Dirty_Flag As Boolean)
		Dim f As Short
		Dim fn_This As String
		Dim DummyStr1 As String
		Dim DummyVal1 As Short
		Dim i As Short
		Dim Flag_IPE As Short
		'UPGRADE_WARNING: Arrays in structure MI may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim MI As ModelIPE_SPEQ_Inputs_Type
		'UPGRADE_WARNING: Arrays in structure MO may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim MO As ModelIPE_SPEQ_Outputs_Type
		Dim OkayToUse As Boolean
		Dim EOFTESTMARKER As Double
		'READ SUCCESS FLAG OUTPUT FILE.
		MO.Initialize()   'Shang added
		f = FreeFile()
		fn_This = Exe_Path & "\" & ModelIPE_OUT_SuccessFlag
		If (Not FileExists(fn_This)) Then
			Call Show_Error("Unable to find output file: Calculations failed.")
			Exit Sub
		End If
		FileOpen(f, fn_This, OpenMode.Input)
		DummyStr1 = LineInput(f)
		Input(f, Flag_IPE)
		FileClose(f)
		'READ MAIN OUTPUT FILE.
		fn_This = Exe_Path & "\" & ModelIPE_OUT_Main
		FileOpen(f, fn_This, OpenMode.Input)
		DummyStr1 = LineInput(f) : Input(f, MO.CSAV)
		DummyStr1 = LineInput(f) : Input(f, MO.QSAV)
		DummyStr1 = LineInput(f) : Input(f, MO.XK1)
		DummyStr1 = LineInput(f) : Input(f, MO.XK2)
		DummyStr1 = LineInput(f) : Input(f, MO.XNF)
		DummyStr1 = LineInput(f) : Input(f, MO.CBEG)
		DummyStr1 = LineInput(f) : Input(f, MO.CEND)
		'Line Input #f, DummyStr1: Input #f, MO.RSQD
		'Line Input #f, DummyStr1: Input #f, MO.RMSE
		DummyStr1 = LineInput(f)
		For i = 1 To 30
			Input(f, MO.ErrMat(i))
		Next i
		DummyStr1 = LineInput(f)
		Input(f, MO.SQERR) 'COPY OF SQERR / Flag_IPE.
		DummyStr1 = LineInput(f)
		Input(f, EOFTESTMARKER)
		If (False = ModelIO_DoNumberCheck(EOFTESTMARKER, ModelIPE_EofTestMarker)) Then
			Call Show_Error("The model calculations failed: invalid file format (EOF marker).")
			Exit Sub
		End If
		FileClose(f)
		'UPGRADE_WARNING: Couldn't resolve default property of object ModelIPE_SPEQ_Outputs. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ModelIPE_SPEQ_Outputs = MO
		'UPGRADE_WARNING: Couldn't resolve default property of object MI. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		MI = ModelIPE_SPEQ_Inputs
		'DISPLAY WARNINGS/ERRORS IF NECESSARY.
		OkayToUse = AllIPEModels_ErrorCheck(MO.SQERR, MO.ErrMat)
		If (Not OkayToUse) Then Exit Sub
		'PROCESS IPE RESULTS.
		IPES_Data.Input_Renamed.BB = MI.IN_DATA(1)
		IPES_Data.Input_Renamed.W0 = MI.IN_DATA(2)
		IPES_Data.Input_Renamed.GM = MI.IN_DATA(9)
		'---- Conversion from ug/g to mg/g
		IPES_Data.Output.QSAV = MO.QSAV / 1000.0#
		'---- Conversion from ug/g to mg/g
		IPES_Data.Output.CSAV = MO.CSAV / 1000.0#
		'IPES_Data.Output.CBEG = MO.CBEG / 1000#
		'IPES_Data.Output.CEND = MO.CEND / 1000#
		'---- Conversion from (ug/g)x(l/ug)^(1/n) to (mg/g)x(l/mg)^(1/n)
		IPES_Data.Output.XK1 = MO.XK1 * (1000.0#) ^ (MO.XNF - 1)
		'---- Conversion from (umol/g)x(l/umol)^(1/n) to (mmol/g)x(l/mmol)^(1/n)
		IPES_Data.Output.XK2 = MO.XK2 * (1000.0#) ^ (MO.XNF - 1)
		IPES_Data.Output.XN = MO.XNF
		'IPES_Data.Output.RSQD = MO.RSQD
		'IPES_Data.Output.RMSE = MO.RMSE
		'CURRENTLY, THE SPEQ() ROUTINE DOES NOT PROPERLY OUTPUT
		'THE VALUES FOR RSQD, RMSE, CBED, OR CEND.
		IPES_Data.Output.RSQD = 0#
		IPES_Data.Output.RMSE = 0#
		IPES_Data.Output.CBEG = 0#
		IPES_Data.Output.CEND = 0#
		'TRANSFER K AND 1/n.
		Component(0).IPESResult_K = IPES_Data.Output.XK1
		Component(0).IPESResult_OneOverN = IPES_Data.Output.XN
		'RAISE DIRTY FLAG.
		Raise_Dirty_Flag = True
		'DISPLAY RESULTS.
		Call frmModelIPEResults.frmModelIPEResults_Run(SHARED_MODULECODE)
	End Sub
	Sub ModelIPE_SPEQ_WriteMainFile()
		'UPGRADE_WARNING: Arrays in structure MI may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim MI As ModelIPE_SPEQ_Inputs_Type
		Dim f As Short
		Dim fn_This As String
		'PREPARE INPUTS.
		'NOTE: IT IS ASSUMED THAT Component(0) CONTAINS THE
		'CHEMICAL PROPERTIES OF INTEREST.
		MI.Initialize()

		MI.IN_DATA(1) = Carbon.BB
		MI.IN_DATA(2) = Carbon.W0
		MI.IN_DATA(3) = Bed.Temperature + 273.15
		MI.IN_DATA(4) = Component(0).InitialConcentration * 1000.0#
		MI.IN_DATA(5) = Component(0).Liquid_Density
		MI.IN_DATA(6) = Component(0).MW
		MI.IN_DATA(7) = Component(0).Vapor_Pressure / 101325 * 760 'UNITS: mmHg.
		MI.IN_DATA(8) = Component(0).Refractive_Index
		MI.IN_DATA(9) = Carbon.PolanyiExponent
		MI.IN_DATA(10) = 0.000001

		'MI.BB = Carbon.BB
		'MI.W0 = Carbon.W0
		'MI.GM = Carbon.PolanyiExponent
		'MI.CBULK = Component(0).InitialConcentration * 1000.0#
		'MI.ORGDEN = Component(0).Liquid_Density
		'MI.TT = Bed.Temperature + 273.15
		'MI.FWT = Component(0).MW
		'MI.SOLUB = Component(0).Aqueous_Solubility
		'MI.NL = SHARED_NL
		'MI.OMAG = SHARED_OMAG
		'MI.VOLM_NBP = Component(0).MolarVolume


		'WARNING: If IN_DATA(10) (the tolerance for the SPEQ()
		'subroutine) is set to 1e-7, 1e-8, or lower, the SPEQ()
		'routine will attempt to achieve a ridiculous number of
		'significant figures.  Because 3 significant figures are
		'perhaps above the limit for most K and 1/n calculations
		'of this type, a tolerance of 1e-6 or even 1e-5 or 1e-4
		'should be used instead.
		MI.NL = SHARED_NL
		MI.XERR = 0 'I DON'T KNOW WHAT THIS VARIABLE DOES.
		'MI.XERR = CInt(SHARED_OMAG)
		'WRITE INPUT FILE.
		f = FreeFile
		fn_This = Exe_Path & "\" & ModelIPE_IN_Main
		FileOpen(f, fn_This, OpenMode.Output)
		Call WriteFortranInput(f, ModelIPE_Version, "MODULE_VERSION")
		Call WriteFortranInput(f, MI.IN_DATA(1), "IN_DATA(1), Polanyi parameter BB")
		Call WriteFortranInput(f, MI.IN_DATA(2), "IN_DATA(2), Polanyi parameter W0")
		Call WriteFortranInput(f, MI.IN_DATA(3), "IN_DATA(3), temperature, degK")
		Call WriteFortranInput(f, MI.IN_DATA(4), "IN_DATA(4), bulk concentration, ug/L")
		Call WriteFortranInput(f, MI.IN_DATA(5), "IN_DATA(5), organic density, g/cm^3")
		Call WriteFortranInput(f, MI.IN_DATA(6), "IN_DATA(6), molecular weight, g/gmol")
		Call WriteFortranInput(f, MI.IN_DATA(7), "IN_DATA(7), vapor pressure, mmHg")
		Call WriteFortranInput(f, MI.IN_DATA(8), "IN_DATA(8), refractive index, dimless")
		Call WriteFortranInput(f, MI.IN_DATA(9), "IN_DATA(9), Polanyi exponent GM, dimless")
		Call WriteFortranInput(f, MI.IN_DATA(10), "IN_DATA(10), SPEQ numerical tolerance, e.g. 1e-6, dimless")
		Call WriteFortranInput(f, MI.NL, "NL, number of regression points, dimless")
		Call WriteFortranInput(f, MI.XERR, "XERR, not sure what this does; 0 seems a good value")
		Call WriteFortranInput(f, ModelIPE_EofTestMarker, "EOFTESTMARKER")
		FileClose(f)
		'STORE FOR LATER USE.
		'UPGRADE_WARNING: Couldn't resolve default property of object ModelIPE_SPEQ_Inputs. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ModelIPE_SPEQ_Inputs = MI
	End Sub
	Sub ModelIPE_SPEQ_Go(ByRef Raise_Dirty_Flag As Boolean)
		Call ModelIPE_WritePathFile(MODULECODE_SPEQ)
		Call ModelIPE_SPEQ_WriteMainFile()
		Call ModelIPE_CallEXE()
		Call ModelIPE_SPEQ_ProcessOutput(Raise_Dirty_Flag)
		If (ModelIO_IsKeepTempFiles() = False) Then
			Call ModelIPE_RemoveLinkFiles()
		End If
	End Sub
	
	
	'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	'////////////////    SHARED SUBROUTINES    ////////////////////////////////////////////////////////////////////////////
	'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	Sub ModelIPE_CallEXE()
		Dim CmdLine As String
		Call ChangeDir_Exes()
		CmdLine = ModelIPE_ExeName
		Call FortranLink_ExecAndWaitForProcess(CmdLine)
		Call ChangeDir_Main()
	End Sub
	Sub ModelIPE_WritePathFile(ByRef WhichModule As Short)
		Dim f As Short
		Dim fn_This As String
		Dim qq As String
		qq = Chr(34)
		f = FreeFile
		fn_This = Exe_Path & "\" & ModelIPE_IN_PathFile
		FileOpen(f, fn_This, OpenMode.Output)
		PrintLine(f, Trim(Str(WhichModule)))
		PrintLine(f, qq & ModelIPE_IN_Main & qq)
		PrintLine(f, qq & ModelIPE_OUT_SuccessFlag & qq)
		PrintLine(f, qq & ModelIPE_OUT_Main & qq)
		FileClose(f)
	End Sub
	Sub ModelIPE_RemoveLinkFiles()
		Call KillFile_If_Exists(Exe_Path & "\" & ModelIPE_IN_PathFile)
		Call KillFile_If_Exists(Exe_Path & "\" & ModelIPE_IN_Main)
		Call KillFile_If_Exists(Exe_Path & "\" & ModelIPE_OUT_SuccessFlag)
		Call KillFile_If_Exists(Exe_Path & "\" & ModelIPE_OUT_Main)
	End Sub
	'RETURNS:
	'    TRUE = OKAY TO USE THIS DATA.
	'    FALSE = ERRORS HAVE INVALIDATED THIS DATA.
	Function AllIPEModels_ErrorCheck(ByRef ErrFlag As Short, ByRef ErrMat() As Short) As Boolean
		Dim Ret_OkayToUse As Boolean
		Dim NoMsgs As Boolean
		Dim OnlyWarning As Boolean
		Dim temp As String
		Dim ThisMsg As String
		Dim i As Short
		Ret_OkayToUse = True
		NoMsgs = True
		OnlyWarning = False
		If ((ErrFlag <> 0) Or (ErrMat(1) <> 0)) Then
			NoMsgs = False
		End If
		If (Not NoMsgs) Then
			temp = ""
			For i = 1 To 30
				If (ErrMat(i) = 0) Then Exit For
				If (i <> 1) Then temp = temp & vbCrLf
				Select Case ErrMat(i)
					Case 11
						ThisMsg = "11 -- ERROR: You specified a bulk concentration that is greater than the chemical's saturation concentration."
					Case 12
						ThisMsg = "12 -- WARNING: You specified a bulk concentration " & "and order of magnitude " & "which define a concentration range that goes " & "higher than the chemical's saturation " & "concentration.  The upper limit of concentration was " & "adjusted to 99% of the " & "solubility concentration."
						OnlyWarning = True
					Case 13
						ThisMsg = "13 -- ERROR: You specified inappropriate isotherm " & "regression limits: CBEG > CEND."
					Case 14
						ThisMsg = "14 -- ERROR: There was a mathematical error: " & "the model tried to take the DEXP() of a number < -710."
					Case 15
						ThisMsg = "15 -- ERROR: There was a mathematical error: " & "the model tried to raise ten to a number < -710."
					Case 16
						ThisMsg = "16 -- ERROR: The Polanyi correlation range was exceeded (QCAP < 1.0E-03)."
					Case 17
						ThisMsg = "17 -- WARNING: Some of the highest concentrations " & "in the concentration range specified are " & "in the 'pore-filling' regime where " & "(Pi/Ps) > 0.2, and therefore may correspond " & "to capillary condensation."
						OnlyWarning = True
					Case 18
						ThisMsg = "18 -- ERROR: You specified a bulk concentration " & "that is greater than the chemical's saturation concentration."
					Case 19
						ThisMsg = "19 -- ERROR: The upper isotherm limit exceeds the " & "chemical's saturation concentration."
					Case 20
						ThisMsg = "20 -- ERROR: Error in non-linear equation routine GOLDEN."
					Case 21
						ThisMsg = "21 -- ERROR: The D-R correlation range was exceeded (QCAP < 1.0E-03)."
					Case Else
						ThisMsg = "Unknown error #" & Trim(Str(ErrMat(i)))
				End Select
				temp = temp & ThisMsg
			Next i
			If (OnlyWarning) Then
				temp = "The following warning(s) occurred:" & vbCrLf & vbCrLf & temp
			Else
				temp = "The following warning(s) and/or error(s) occurred:" & vbCrLf & vbCrLf & temp
			End If
			Call Show_Error(temp)
			'MsgBox "The following errors occured:" & Chr$(13) & Format$(Flag, "0") & Chr$(13) & Format$(Flag2(1), "0"), 64, AppName_For_Display_long
		End If
		If ((OnlyWarning) Or (NoMsgs)) Then
			Ret_OkayToUse = True
		Else
			Ret_OkayToUse = False
		End If
		AllIPEModels_ErrorCheck = Ret_OkayToUse
	End Function
End Module