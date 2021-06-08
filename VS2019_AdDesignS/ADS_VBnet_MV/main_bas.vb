Option Strict Off
Option Explicit On
Module MainProgram
	
	'GLOBAL CONSTANTS -- APPLICATION RELATED.
	''''Global Const AppCopyright = "Michigan Technological University, 1994-99"
	Public Const AppRegisteredUser As String = ""
	Public Const AppRegisteredCompany As String = ""
	Public Const AppRegisteredSerial As String = ""

	Public MAIN_APP_PATH As String
	
	
	'splash_mode: 0 = Continue/Exit window
	'             1 = I Agree/I agree, never show again/Exit window
	Public splash_mode As Short
	
	'splash_button_pressed:
	'1 = Continue or I Agree
	'2 = I agree, never show again
	'3 = Exit
	Public splash_button_pressed As Short
	
	Public booActLikeBeta As Boolean
	
	
	
	
	
	Const MainProgram_declarations_end As Short = 0
	
	
	Function get_program_version_with_build_info_VB4(ByRef IsOnFrontWindow As Boolean) As String
		Dim ver As String
		Dim This_get_program_releasetype As String
		Dim Show_ReleaseType As Boolean
		ver = "Version "
		ver = ver & Trim(CStr(My.Application.Info.Version.Major))
		ver = ver & "." & Trim(CStr(My.Application.Info.Version.Minor))
		If (IsOnFrontWindow = False) Then
			ver = ver & "." & Trim(CStr(My.Application.Info.Version.Revision))
		End If
		This_get_program_releasetype = get_program_releasetype()
		Show_ReleaseType = True
		If (UCase(This_get_program_releasetype) = "STANDARD") Then
			If (IsOnFrontWindow = True) Then
				Show_ReleaseType = False
			End If
		End If
		If (Show_ReleaseType = True) Then
			ver = ver & " (" & get_program_releasetype() & ")"
		End If
		get_program_version_with_build_info_VB4 = ver
	End Function
	
	
	Function IsThisADemo() As Boolean
		Dim This_get_program_releasetype As String
		This_get_program_releasetype = get_program_releasetype()
		If (booActLikeBeta) Or (UCase(This_get_program_releasetype) = "BETA") Then
			IsThisADemo = True
		Else
			IsThisADemo = False
		End If
	End Function
	Sub Demo_ShowError(ByRef strMsg As String)
		Call Show_Error(strMsg & vbCrLf & vbCrLf & "For the full version of this program, please contact " & "Dr. David W. Hand (dwhand@mtu.edu or 906-487-2777). " & "Additional information about this program is available at " & "our web site (http://www.cpas.mtu.edu/etdot/).")
	End Sub
	Function Demo_AreValuesEqual(ByRef IN_dblVal1 As Double, ByRef IN_dblVal2 As Double) As Boolean
		Dim intSigFigs As Short
		Dim dblSigFigCutoff As Double
		Dim dblTest As Double
		If (IN_dblVal1 = 0#) Or (IN_dblVal2 = 0#) Then
			' DOES NOT HANDLE VALUES OF 0!
			Demo_AreValuesEqual = False
			Exit Function
		End If
		intSigFigs = 8
		dblSigFigCutoff = 10# ^ (-CDbl(intSigFigs))
		dblTest = System.Math.Abs((IN_dblVal1 / IN_dblVal2) - 1#)
		If (dblTest > dblSigFigCutoff) Then
			Demo_AreValuesEqual = False
		Else
			Demo_AreValuesEqual = True
		End If
	End Function
	Function Demo_CheckForValidFile(ByRef dblDemoChecksum As Double) As Boolean
		Dim booIsOkay As Boolean
		booIsOkay = False
		If (Demo_AreValuesEqual(dblDemoChecksum, 629.110977410767) = True) Then
			booIsOkay = True
		End If
		If (Demo_AreValuesEqual(dblDemoChecksum, 719.476218564802) = True) Then
			booIsOkay = True
		End If
		Demo_CheckForValidFile = booIsOkay
	End Function
	
	
	Function frmSplash_Run() As Short
		Dim tpath As String
		Dim tstr As String
		Dim must_read_disclaimer As Short
		Dim msg As String

		'	'''SET UP INI FILE PATH.
		''tpath$ = GetWindowsDir() & ProgramIniFile$

		'SHOW THE CONTINUE/EXIT FRONT WINDOW.
		splash_mode = 0
		splash_button_pressed = 0
		On Error GoTo err_frmSplash_Run
		'Error 5
		frmSplash.ShowDialog()
		'	asplashform = New frmSplash
		'   asplashform.ShowDialog()

		Select Case splash_button_pressed
			Case 1 'Hit Continue
				'DO NOTHING.
			Case 3 'Hit Exit
				End
		End Select
		
		'IS THE DISCLAIMER WINDOW STILL ACTIVE?
		must_read_disclaimer = True
		''tstr$ = INI_GetSetting0(fn_INI_path, "disclaimer", "has_read_disclaimer")
		tstr = INI_GetSetting0(fn_OldFileList, "disclaimer", "has_read_disclaimer")
		'tstr$ = ini_getsetting("has_seen_disclaimer")
		If (tstr = "1") Then
			must_read_disclaimer = False
		End If
		
		If (1 = 0) Then
			'''''if (must_read_disclaimer) Then
			'SHOW THE DISCLAIMER WINDOW.
			splash_mode = 1
			splash_button_pressed = 0
			frmSplash.ShowDialog()
			Select Case splash_button_pressed
				Case 1 'Hit I Agree
					'DO NOTHING.
				Case 2 'Hit I agree, never show again
					''Call ini_putsetting0(fn_INI_path, "disclaimer", "has_read_disclaimer", "1")
					Call ini_putsetting0(fn_OldFileList, "disclaimer", "has_read_disclaimer", "1")
					'Call ini_putsetting("has_seen_disclaimer", "1")
				Case 3 'Hit Exit
					End
			End Select
		End If
		
		frmSplash_Run = True
		Exit Function
		
exit_err_frmSplash_Run: 
		Call Show_Error("Halting due to an error.")
		End
err_frmSplash_Run: 
		msg = "Detected an error.  " & "Err.Number = " & Trim(Str(Err.Number)) & "; " & "Err.Source = `" & Err.Source & "`.  Now halting program."
		Call Show_Message(msg)
		Resume exit_err_frmSplash_Run
	End Function
	
	
	Sub ChangeDir_Exes()
		ChDrive(MAIN_APP_PATH)
		ChDir(MAIN_APP_PATH & "\EXES")
	End Sub
	Sub ChangeDir_Main()
		ChDrive(MAIN_APP_PATH)
		ChDir(MAIN_APP_PATH)
	End Sub
	
	
	Function CheckFileExistence_Critical(ByRef fn_This As String) As Boolean
		If (File_IsExists(fn_This)) Then
			'DO NOTHING; THIS IS OKAY.
			CheckFileExistence_Critical = True
		Else
			Call Show_Error("The file `" & fn_This & "` is missing.  " & "Therefore the software must have been improperly installed.  " & "Recommendation: Check the `Start In` path specified in the " & "program icon, or else perform a re-install of the software.")
			CheckFileExistence_Critical = False
		End If
	End Function
	'UPGRADE_WARNING: Application will terminate when Sub Main() finishes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="E08DDC71-66BA-424F-A612-80AF11498FF8"'
	Public Sub Main()
		Dim Workspaces As Object
		Dim fn_misc1 As String
		Dim LicFileLocation As Short
		Dim fpath_INI As String
		Dim msg As String
		Dim fn_Test As String
		'
		' SET UP MAIN APP PATH VARIABLE.
		'
		If (File_IsExists(My.Application.Info.DirectoryPath & "\debug_in_vb6.txt")) Then
			'FOR DEBUGGING IN THE VB5 ENVIRONMENT.
			MAIN_APP_PATH = "x:\ads\vb6"
			ChDrive(MAIN_APP_PATH)
			ChDir(MAIN_APP_PATH)
		Else
			'DO NOTHING.
			MAIN_APP_PATH = My.Application.Info.DirectoryPath
		End If
		'
		' VERIFY THAT PATHS ARE PROPERLY SET UP.
		'
		fn_misc1 = My.Application.Info.DirectoryPath & "\dbase\misc1.dat"
		If (CheckFileExistence_Critical(fn_misc1) = False) Then End
		'  If (File_IsExists(fn_misc1)) Then
		'    'DO NOTHING; THIS IS OKAY.
		'  Else
		'    Call Show_Error("The file `" & fn_misc1 & "` is missing.  " & _
		''        "Therefore the software must have been improperly installed.  " & _
		''        "Recommendation: Check the `Start In` path specified in the " & _
		''        "program icon, or else perform a re-install of the software.")
		'    End
		'  End If
		fn_Test = MAIN_APP_PATH & "\dbase\template.dat"
		If (CheckFileExistence_Critical(fn_Test) = False) Then End
		'  If (File_IsExists(fn_Test)) Then
		'    'DO NOTHING; THIS IS OKAY.
		'  Else
		'    Call Show_Error("The file `" & fn_Test & "` is missing.  " & _
		''        "Therefore the software must have been improperly installed.  " & _
		''        "Recommendation: Check the `Start In` path specified in the " & _
		''        "program icon, or else perform a re-install of the software.")
		'    End
		'  End If
		booActLikeBeta = False
		If (File_IsExists(MAIN_APP_PATH & "\actlikebeta.txt")) Then
			booActLikeBeta = True
		End If
		'
		' READ IN THE LICENSE FILE DATA.
		'
		If (TURN_LICENSING_OFF = True) Then
			lfd.Z_USERNAME = "Unspecified User"
			lfd.Z_USERCOMPANY = "Unspecified Company"
			lfd.Z_SERIALNUMBER = "Unspecified Serial Number"
			lfd.Z_RELEASETYPE = "STANDARD"
			fpath_INI = GetWindowsDir()
		Else
			Call LicFileData_Read(Global_fpath_dir_CPAS)
			fpath_INI = Global_fpath_dir_CPAS & "\DBASE"
		End If
		''READ IN THE LICENSE FILE DATA.
		'Call LicFileData_Read(LicFileLocation)
		'Select Case LicFileLocation
		'  Case LICFILELOCATION_WIN:
		'    fpath_INI = GetWindowsDir()
		'  Case LICFILELOCATION_APPPATH:
		'    fpath_INI = App.Path
		'End Select
		'
		' PSDM IN ROOM INITS.
		'
		If (Distribute_PSDMInRoom = False) Then
			Activate_PSDMInRoom = False
		Else
			If (FileExists(My.Application.Info.DirectoryPath & "\PSDMROOM.DAT") = True) Then
				Activate_PSDMInRoom = True
			Else
				Activate_PSDMInRoom = False
			End If
		End If
		If (Activate_PSDMInRoom = True) Then
			'AppName_For_Display_Short = "IAFM"
			'AppName_For_Display_Long = "Indoor Air Filtration Model"
			AppName_For_Display_Short = "AdDesignS" '///MOdefication///Sinan///07/04/06,Old///IndoorAirAdDesignS
			AppName_For_Display_Long = "Adsorption Design Software" '///Modefication///Sinan///07/04/06,Old///Indoor Air Adsorption Design Software
		Else
			AppName_For_Display_Short = "AdDesignS"
			AppName_For_Display_Long = "Adsorption Design Software"
		End If
		
		On Error GoTo err_main
		'
		''ENSURE THAT CODE REALIZES IT NEEDS TO CREATE A NEW PROJECT.
		'
		'NowProj_exists = False

		'app
		'
		' OPEN WORKSPACE TO HOLD DATABASES, STORE DATABASE NAMES.
		'
		DAOEngine = New DAO.DBEngine
		Ws1 = DAOEngine.Workspaces(0)       'out Shang
		fn_DB_dir = My.Application.Info.DirectoryPath & "\dbase"
		Database_Path = fn_DB_dir
		fn_DB_Isotherm = fn_DB_dir & "\isotherm.mdb"
		fn_DB_Carbon = fn_DB_dir & "\carbon.mdb"
		Exe_Path = My.Application.Info.DirectoryPath & "\exes"
		If (CheckFileExistence_Critical(fn_DB_Isotherm) = False) Then End
		If (CheckFileExistence_Critical(fn_DB_Carbon) = False) Then End
		If (CheckFileExistence_Critical(fn_DB_dir & "\beds1.txt") = False) Then End
		If (CheckFileExistence_Critical(fn_DB_dir & "\beds2.txt") = False) Then End
		If (CheckFileExistence_Critical(fn_DB_dir & "\corr_com.txt") = False) Then End
		If (CheckFileExistence_Critical(fn_DB_dir & "\water_co.txt") = False) Then End
		'TODOTODO: Add checks to verify that each of these
		'databases is available for exclusive use
		'by this program.
		'
		' SET UP INI FILENAME FOR VARIOUS USER PREFERENCES, INCLUDING LAST-FEW-FILES LISTS.
		'
		fn_OldFileList = fpath_INI & "\" & fn_INI_name
		
		''TEMPORARILY: DO NOT LOAD frmSplash.
		'If (1 = 0) Then
		'LOAD THE SPLASH WINDOW.
		If (frmSplash_Run() = False) Then
			End
		End If
		'End If
		'
		' SHOW THE DEMO WINDOW.
		'
		If (IsThisADemo() = True) Then
			Call frmDemo.frmDemo_GO()
		End If
		'
		' INITIALIZE THE UNIT STRUCTURES.
		'
		Call unitsys_initialize()
		'
		' LAUNCH THE MAIN WINDOW.
		'
		frmMain.ShowDialog()
		Exit Sub
		
exit_err_main: 
		Call Show_Error("Halting due to an error.")
		End
err_main: 
		msg = "Detected an error in main().  " & "Err.Number = " & Trim(Str(Err.Number)) & "; " & "Err.Source = `" & Err.Source & "`.  Now halting program."
		Call Show_Message(msg)
		Resume exit_err_main
	End Sub
	
	
	Sub debug_output(ByRef s As String)
		Dim f As Short
		f = FreeFile
		FileOpen(f, "c:\bug.txt", OpenMode.Append)
		WriteLine(f, "ADS", DateString & " " & TimeString & " -- " & s)
		FileClose(f)
	End Sub
	
	
	'Returns:
	'TRUE = The program is internal to MTU, thus show the hidden menu
	'FALSE = The program is distributed, hide the menu
	Function check_internal_to_mtu() As Short
		Dim file_1_not_found As Short
		Dim file_2_not_found As Short
		Dim is_internal_to_mtu As Short
		
		On Error GoTo err_check_internal_to_mtu1
		file_1_not_found = True
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If (Dir("c:\a_aaaaaa\internal.txt") <> "") Then file_1_not_found = False
		
res_err_check_internal_to_mtu1: 
		On Error GoTo err_check_internal_to_mtu1
		file_2_not_found = True
		'If (Dir("g:\a_aaaaaa\internal.txt") <> "") Then file_2_not_found = False
		'NOTE: Scanning the G: drive on some computers causes a
		'"hanging" problem so this scan was removed.  EJO, 1/6/98.
		
res_err_check_internal_to_mtu2: 
		is_internal_to_mtu = True
		If (file_1_not_found) And (file_2_not_found) Then
			is_internal_to_mtu = False
		End If
		check_internal_to_mtu = is_internal_to_mtu
		Exit Function
		
err_check_internal_to_mtu1: 
		file_1_not_found = True
		Resume res_err_check_internal_to_mtu1
		
err_check_internal_to_mtu2: 
		file_2_not_found = True
		Resume res_err_check_internal_to_mtu2
		
	End Function
End Module