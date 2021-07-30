Option Strict Off
Option Explicit On

Module MiscUI
	
	Public Declare Function ShellExecute Lib "shell32.dll"  Alias "ShellExecuteA"(ByVal hwnd As Integer, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Integer) As Integer
	
	Public Const CBOYAXISTYPE_C_CO As Short = 1
	Public Const CBOYAXISTYPE_UG_L As Short = 2
	Public Const CBOYAXISTYPE_MG_L As Short = 3
	Public Const CBOYAXISTYPE_G_L As Short = 4
	Public Const CBOYAXISTYPE_PPB As Short = 5
	Public Const CBOYAXISTYPE_PPM As Short = 6
	Public Const CBOYAXISTYPE_NG_L As Short = 7






	Const MiscUI_declarations_end As Boolean = True
	
	
	'
	' THIS FUNCTION RETURNS THE CONVERSION FACTOR THAT THE VALUE OF
	' Results.CP() MUST BE MULTIPLIED BY TO GET THE VALUE IN
	' THE DESIRED UNITS OF DISPLAY.
	'
	'      If (intIsPSDMInRoomModel = False) Then
	'        'the .CP() units are C/Co
	'      Else
	'        If (intAnyCrCloseToZero = True) Then
	'          'the .CP() units are ug/L
	'        Else
	'          'the .CP() units are Cr/Cr,ss
	'        End If
	'      End If
	'
	Function CBOYAXISTYPE_GetUnitConversion(ByRef intCBOYAXISTYPE As Short, ByRef intIsPSDMInRoomModel As Short, ByRef intAnyCrCloseToZero As Short, ByRef intComponentNum As Short, ByRef intBedPhase As Short, ByRef OUT_strYAxisTitle As String) As Double
		Dim strConcName As String
		Dim dblRetVal As Double
		Dim dbl_Cr_ss As Double 'ug/L
		Dim dbl_Co As Double 'ug/L
		Dim dbl_ConvertTo_ug_L As Double 'what it takes to convert the .CP() value to ug/L
		Dim dbl_ConvertFrom_ppm_To_ug_L As Double
		Dim dbl_ConvertFrom_ug_L_To_ppm As Double
		Dim dbl_Pressure_Pa As Double
		Dim dbl_R_J_gmol_K As Double
		Dim dbl_T_K As Double
		Dim dbl_MolecWeight As Double
		'
		' SET UP THE CONVERSION FACTOR FOR ug/L <===> ppm.
		'
		Select Case intBedPhase
			'
			'////////////////////////////////////////////////////////////////////////////////////
			'////////////////////////   LIQUID PHASE
			Case 0
				dbl_ConvertFrom_ug_L_To_ppm = 1 / 1000#
				'
				'////////////////////////////////////////////////////////////////////////////////////
				'////////////////////////   GAS PHASE
			Case 1
				dbl_R_J_gmol_K = 8.31451
				dbl_Pressure_Pa = Results.Bed.Pressure * 101325#
				dbl_T_K = Results.Bed.Temperature + 273.15
				dbl_MolecWeight = Results.Component(intComponentNum).MW
				dbl_ConvertFrom_ppm_To_ug_L = 1# / 1000000# * (dbl_Pressure_Pa) / (dbl_R_J_gmol_K) / (dbl_T_K) * 1000000# * dbl_MolecWeight / 1000#
				dbl_ConvertFrom_ug_L_To_ppm = 1# / dbl_ConvertFrom_ppm_To_ug_L
		End Select
		'
		' DETERMINE HOW TO CONVERT .CP() INTO ug/L.
		'
		If (intIsPSDMInRoomModel = True) Then
			strConcName = "Cr"
			dbl_Cr_ss = Results.psdmroom_Crss(intComponentNum)
			If (intAnyCrCloseToZero = True) Then
				dbl_ConvertTo_ug_L = 1#
			Else
				dbl_ConvertTo_ug_L = dbl_Cr_ss
			End If
		Else
			strConcName = "C"
			dbl_Co = 1000# * Results.Component(intComponentNum).InitialConcentration
			' THE PREVIOUS LINE CONVERTS mg/L TO ug/L
			dbl_ConvertTo_ug_L = dbl_Co
		End If
		'
		' THE MAIN CODE.
		'
		Select Case intCBOYAXISTYPE                  'Plus 1 by Shang 0->1
			Case CBOYAXISTYPE_C_CO
				If (intIsPSDMInRoomModel = False) Then
					OUT_strYAxisTitle = "C/Co"
				Else
					If (intAnyCrCloseToZero = True) Then
						OUT_strYAxisTitle = "(  ERROR -- UNAVAILABLE!!!  )"
					Else
						OUT_strYAxisTitle = "Cr/Cr,ss"
					End If
				End If
				dblRetVal = 1.0#
			Case CBOYAXISTYPE_UG_L
				OUT_strYAxisTitle = strConcName & ", µg/L"
				dblRetVal = dbl_ConvertTo_ug_L
			Case CBOYAXISTYPE_MG_L
				OUT_strYAxisTitle = strConcName & ", mg/L"
				dblRetVal = dbl_ConvertTo_ug_L / 1000.0#
			Case CBOYAXISTYPE_G_L
				OUT_strYAxisTitle = strConcName & ", g/L"
				dblRetVal = dbl_ConvertTo_ug_L / 1000.0# / 1000.0#
			Case CBOYAXISTYPE_PPB
				OUT_strYAxisTitle = strConcName & ", ppb"
				dblRetVal = dbl_ConvertTo_ug_L * dbl_ConvertFrom_ug_L_To_ppm * 1000.0#
			Case CBOYAXISTYPE_PPM
				OUT_strYAxisTitle = strConcName & ", ppm"
				dblRetVal = dbl_ConvertTo_ug_L * dbl_ConvertFrom_ug_L_To_ppm
			Case CBOYAXISTYPE_NG_L
				OUT_strYAxisTitle = strConcName & ", ng/L"
				dblRetVal = dbl_ConvertTo_ug_L * 1000.0#
		End Select
		CBOYAXISTYPE_GetUnitConversion = dblRetVal
	End Function
	
	Sub ShellExecute_LocalFile(ByRef in_Filename As String)
		Call ShellExecute(0, vbNullString, in_Filename, vbNullString, vbNullString, AppWinStyle.NormalFocus)
	End Sub
	Sub ShellExecute_URL(ByRef in_URL As String)
		Call ShellExecute(0, vbNullString, in_URL, vbNullString, vbNullString, AppWinStyle.NormalFocus)
	End Sub
	
	
	Sub CalcStatus_Set(ByRef newVal As Boolean)
		If (newVal) Then
			Call GenericStatus_Set("Calculating -- please wait.")
		Else
			Call GenericStatus_Set("")
		End If
	End Sub
	Sub GenericStatus_Set(ByRef fn_Text As String)
		'UPGRADE_WARNING: Couldn't resolve default property of object frmMain.sspanel_Status. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'frmMain.sspanel_Status.Caption = fn_Text
		frmMain.ToolStripStatusLabelStatus.Text = fn_Text
	End Sub
	Sub DirtyStatus_Set(ByRef newVal As Boolean)
		If (IsThisADemo() = True) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object frmMain.sspanel_Dirty. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'frmMain.sspanel_Dirty.Caption = "* DEMO VERSION *"
			'UPGRADE_WARNING: Couldn't resolve default property of object frmMain.sspanel_Dirty.ForeColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'frmMain.sspanel_Dirty.ForeColor = Color.FromArgb(QBColor(12))
		Else
			If (newVal) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object frmMain.sspanel_Dirty. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'frmMain.sspanel_Dirty.Caption = "Data Changed"
				frmMain.ToolStripStatusLabelDirty.Text = "Data Changed"
				'UPGRADE_WARNING: Couldn't resolve default property of object frmMain.sspanel_Dirty.ForeColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'frmMain.sspanel_Dirty.ForeColor = Color.FromArgb(QBColor(12))
				frmMain.ToolStripStatusLabelDirty.ForeColor = Color.FromArgb(QBColor(12))
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object frmMain.sspanel_Dirty. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'frmMain.sspanel_Dirty.Caption = "Unchanged"
				frmMain.ToolStripStatusLabelDirty.Text = "Unchanged"
				'UPGRADE_WARNING: Couldn't resolve default property of object frmMain.sspanel_Dirty.ForeColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'frmMain.sspanel_Dirty.ForeColor = Color.FromArgb(QBColor(0))
				frmMain.ToolStripStatusLabelDirty.ForeColor = Color.FromArgb(QBColor(0))

			End If
		End If
	End Sub
	Sub DirtyStatus_Set_Current()
		Call DirtyStatus_Set(Project_Is_Dirty)
	End Sub
	Sub DirtyStatus_Throw()
		Project_Is_Dirty = True
		Call DirtyStatus_Set_Current()
	End Sub
	
	
	Sub frmMain_Close_All_Windows()
		Dim ifc As Short
		Dim i As Short
		On Error Resume Next
		ifc = Application.OpenForms.Count - 1
		For i = ifc To 0 Step -1
			'If (Forms(i%).name <> "frmMain") And _
			'(Forms(i%).name <> "frmProgress") Then
			If (Application.OpenForms.Item(i).Name <> "frmMain") Then
				'UPGRADE_ISSUE: Unload Forms() was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="875EBAD7-D704-4539-9969-BC7DBDAA62A2"'
				Application.OpenForms(i).Close()
			End If
		Next i
	End Sub
	
	
	Sub CenterOnScreen(ByRef frm_to_center As System.Windows.Forms.Form)
		frm_to_center.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(frm_to_center.Width)) / 2)
		frm_to_center.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(frm_to_center.Height)) / 2)
	End Sub
	Sub CenterOnForm(ByRef frm_to_center As System.Windows.Forms.Form, ByRef Frm As System.Windows.Forms.Form)
		frm_to_center.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Frm.Left) + (VB6.PixelsToTwipsX(Frm.Width) - VB6.PixelsToTwipsX(frm_to_center.Width)) / 2)
		frm_to_center.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Frm.Top) + (VB6.PixelsToTwipsY(Frm.Height) - VB6.PixelsToTwipsY(frm_to_center.Height)) / 2)
	End Sub
	
	
	Sub Show_Message00(ByRef msg As String, ByRef flags As Short, ByRef WinTitle As String)
		MsgBox(msg, flags, WinTitle)
	End Sub
	Sub Show_Message0(ByRef msg As String, ByRef flags As Short)
		Call Show_Message00(msg, MsgBoxStyle.Information, AppName_For_Display_Short)
	End Sub
	Sub Show_Message(ByRef msg As String)
		Call Show_Message0(msg, MsgBoxStyle.Information)
	End Sub
	Sub Show_Error(ByRef msg As String)
		Beep()
		Call Show_Message0(msg, MsgBoxStyle.Exclamation)
	End Sub
	Sub Show_Trapped_Error(ByRef subname As String)
		Call Show_Error("An error #" & Trim(Str(Err.Number)) & " has occurred in routine " & Trim(subname) & ": `" & Trim(ErrorToString()) & "`.  Ending this operation.")
	End Sub
	
	
	Sub Launch_Notepad(ByRef fn_edit As String)
		Dim CmdLine As String
		Dim RetVal As Short
		CmdLine = "notepad " & fn_edit
		RetVal = 0 * Shell(CmdLine, 3)
	End Sub
End Module