Attribute VB_Name = "MainProgram"
Option Explicit

'GLOBAL CONSTANTS -- APPLICATION RELATED.
''''Global Const AppCopyright = "Michigan Technological University, 1994-99"
Global Const AppRegisteredUser = ""
Global Const AppRegisteredCompany = ""
Global Const AppRegisteredSerial = ""

Global MAIN_APP_PATH As String


'splash_mode: 0 = Continue/Exit window
'             1 = I Agree/I agree, never show again/Exit window
Global splash_mode As Integer

'splash_button_pressed:
'1 = Continue or I Agree
'2 = I agree, never show again
'3 = Exit
Global splash_button_pressed As Integer

Global booActLikeBeta As Boolean





Const MainProgram_declarations_end = 0


Function get_program_version_with_build_info_VB4( _
    IsOnFrontWindow As Boolean) _
    As String
Dim ver As String
Dim This_get_program_releasetype As String
Dim Show_ReleaseType As Boolean
  ver = "Version "
  ver = ver & Trim$(App.Major)
  ver = ver & "." & Trim$(App.Minor)
  If (IsOnFrontWindow = False) Then
    ver = ver & "." & Trim$(App.Revision)
  End If
  This_get_program_releasetype = get_program_releasetype()
  Show_ReleaseType = True
  If (UCase$(This_get_program_releasetype) = "STANDARD") Then
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
  If (booActLikeBeta) Or (UCase$(This_get_program_releasetype) = "BETA") Then
    IsThisADemo = True
  Else
    IsThisADemo = False
  End If
End Function
Sub Demo_ShowError(strMsg As String)
  Call Show_Error(strMsg & _
      vbCrLf & _
      vbCrLf & _
      "For the full version of this program, please contact " & _
      "Dr. David W. Hand (dwhand@mtu.edu or 906-487-2777). " & _
      "Additional information about this program is available at " & _
      "our web site (http://www.cpas.mtu.edu/etdot/).")
End Sub
Function Demo_AreValuesEqual( _
    IN_dblVal1 As Double, _
    IN_dblVal2 As Double) _
    As Boolean
Dim intSigFigs As Integer
Dim dblSigFigCutoff As Double
Dim dblTest As Double
  If (IN_dblVal1 = 0#) Or (IN_dblVal2 = 0#) Then
    ' DOES NOT HANDLE VALUES OF 0!
    Demo_AreValuesEqual = False
    Exit Function
  End If
  intSigFigs = 8
  dblSigFigCutoff = 10# ^ (-CDbl(intSigFigs))
  dblTest = Abs((IN_dblVal1 / IN_dblVal2) - 1#)
  If (dblTest > dblSigFigCutoff) Then
    Demo_AreValuesEqual = False
  Else
    Demo_AreValuesEqual = True
  End If
End Function
Function Demo_CheckForValidFile( _
    dblDemoChecksum As Double) _
    As Boolean
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


Function frmSplash_Run() As Integer
Dim tpath$
Dim tstr$
Dim must_read_disclaimer As Integer
Dim msg As String

  '''SET UP INI FILE PATH.
  ''tpath$ = GetWindowsDir() & ProgramIniFile$
  
  'SHOW THE CONTINUE/EXIT FRONT WINDOW.
  splash_mode = 0
  splash_button_pressed = 0
  On Error GoTo err_frmSplash_Run
'Error 5
  frmSplash.Show 1
  Select Case splash_button_pressed
    Case 1:         'Hit Continue
      'DO NOTHING.
    Case 3:         'Hit Exit
      End
  End Select
    
  'IS THE DISCLAIMER WINDOW STILL ACTIVE?
  must_read_disclaimer = True
  ''tstr$ = INI_GetSetting0(fn_INI_path, "disclaimer", "has_read_disclaimer")
  tstr$ = INI_GetSetting0(fn_OldFileList, "disclaimer", "has_read_disclaimer")
  'tstr$ = ini_getsetting("has_seen_disclaimer")
  If (tstr$ = "1") Then
    must_read_disclaimer = False
  End If
  
  If (1 = 0) Then
  '''''if (must_read_disclaimer) Then
    'SHOW THE DISCLAIMER WINDOW.
    splash_mode = 1
    splash_button_pressed = 0
    frmSplash.Show 1
    Select Case splash_button_pressed
      Case 1:         'Hit I Agree
        'DO NOTHING.
      Case 2:         'Hit I agree, never show again
        ''Call ini_putsetting0(fn_INI_path, "disclaimer", "has_read_disclaimer", "1")
        Call ini_putsetting0(fn_OldFileList, "disclaimer", "has_read_disclaimer", "1")
        'Call ini_putsetting("has_seen_disclaimer", "1")
      Case 3:         'Hit Exit
        End
    End Select
  End If

  frmSplash_Run = True
  Exit Function
  
exit_err_frmSplash_Run:
  Call Show_Error("Halting due to an error.")
  End
err_frmSplash_Run:
  msg = "Detected an error.  " & _
      "Err.Number = " & Trim$(Str$(Err.number)) & "; " & _
      "Err.Source = `" & Err.Source & "`.  Now halting program."
  Call Show_Message(msg)
  Resume exit_err_frmSplash_Run
End Function


Sub ChangeDir_Exes()
  ChDrive MAIN_APP_PATH
  ChDir MAIN_APP_PATH & "\EXES"
End Sub
Sub ChangeDir_Main()
  ChDrive MAIN_APP_PATH
  ChDir MAIN_APP_PATH
End Sub


Function CheckFileExistence_Critical( _
    fn_This As String) _
    As Boolean
  If (File_IsExists(fn_This)) Then
    'DO NOTHING; THIS IS OKAY.
    CheckFileExistence_Critical = True
  Else
    Call Show_Error("The file `" & fn_This & "` is missing.  " & _
        "Therefore the software must have been improperly installed.  " & _
        "Recommendation: Check the `Start In` path specified in the " & _
        "program icon, or else perform a re-install of the software.")
    CheckFileExistence_Critical = False
  End If
End Function
Sub Main()
Dim fn_misc1 As String
Dim LicFileLocation As Integer
Dim fpath_INI As String
Dim msg As String
Dim fn_Test As String
  '
  ' SET UP MAIN APP PATH VARIABLE.
  '
  If (File_IsExists(App.Path & "\debug_in_vb6.txt")) Then
    'FOR DEBUGGING IN THE VB5 ENVIRONMENT.
    MAIN_APP_PATH = "X:\etdot10\code\ads\vb6"
    ChDrive MAIN_APP_PATH
    ChDir MAIN_APP_PATH
  Else
    'DO NOTHING.
    MAIN_APP_PATH = App.Path
  End If
  '
  ' VERIFY THAT PATHS ARE PROPERLY SET UP.
  '
  fn_misc1 = App.Path & "\dbase\misc1.dat"
  If (CheckFileExistence_Critical(fn_misc1) = False) Then End
'  If (File_IsExists(fn_misc1)) Then
'    'DO NOTHING; THIS IS OKAY.
'  Else
'    Call Show_Error("The file `" & fn_misc1 & "` is missing.  " & _
'        "Therefore the software must have been improperly installed.  " & _
'        "Recommendation: Check the `Start In` path specified in the " & _
'        "program icon, or else perform a re-install of the software.")
'    End
'  End If
  fn_Test = MAIN_APP_PATH & "\dbase\template.dat"
  If (CheckFileExistence_Critical(fn_Test) = False) Then End
'  If (File_IsExists(fn_Test)) Then
'    'DO NOTHING; THIS IS OKAY.
'  Else
'    Call Show_Error("The file `" & fn_Test & "` is missing.  " & _
'        "Therefore the software must have been improperly installed.  " & _
'        "Recommendation: Check the `Start In` path specified in the " & _
'        "program icon, or else perform a re-install of the software.")
'    End
'  End If
  booActLikeBeta = False
  If (File_IsExists(MAIN_APP_PATH & "\actlikebeta.txt")) Then
    booActLikeBeta = True
  End If
  '
  ' READ IN THE LICENSE FILE DATA.
  '
  Call LicFileData_Read(Global_fpath_dir_CPAS)
  fpath_INI = Global_fpath_dir_CPAS & "\DBASE"
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
    If (FileExists(App.Path & "\PSDMROOM.DAT") = True) Then
      Activate_PSDMInRoom = True
    Else
      Activate_PSDMInRoom = False
    End If
  End If
  If (Activate_PSDMInRoom = True) Then
    'AppName_For_Display_Short = "IAFM"
    'AppName_For_Display_Long = "Indoor Air Filtration Model"
    AppName_For_Display_Short = "IndoorAirAdDesignS"
    AppName_For_Display_Long = "Indoor Air Adsorption Design Software"
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
  Set Ws1 = Workspaces(0)
  fn_DB_dir = App.Path & "\dbase"
  Database_Path = fn_DB_dir
  fn_DB_Isotherm = fn_DB_dir & "\isotherm.mdb"
  fn_DB_Carbon = fn_DB_dir & "\carbon.mdb"
  Exe_Path = App.Path & "\exes"
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
    Call frmDemo.frmDemo_GO
  End If
  '
  ' INITIALIZE THE UNIT STRUCTURES.
  '
  Call unitsys_initialize
  '
  ' LAUNCH THE MAIN WINDOW.
  '
  frmMain.Show
  Exit Sub

exit_err_main:
  Call Show_Error("Halting due to an error.")
  End
err_main:
  msg = "Detected an error in main().  " & _
      "Err.Number = " & Trim$(Str$(Err.number)) & "; " & _
      "Err.Source = `" & Err.Source & "`.  Now halting program."
  Call Show_Message(msg)
  Resume exit_err_main
End Sub


Sub debug_output(s As String)
Dim f As Integer
  f = FreeFile
  Open "c:\bug.txt" For Append As #f
  Write #f, "ADS", Date$ & " " & Time$ & " -- " & s
  Close #f
End Sub


'Returns:
'TRUE = The program is internal to MTU, thus show the hidden menu
'FALSE = The program is distributed, hide the menu
Function check_internal_to_mtu() As Integer
Dim file_1_not_found As Integer
Dim file_2_not_found As Integer
Dim is_internal_to_mtu As Integer

  On Error GoTo err_check_internal_to_mtu1
  file_1_not_found = True
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

