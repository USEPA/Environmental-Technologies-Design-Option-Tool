Attribute VB_Name = "MainMod"
Option Explicit


'splash_mode: 0 = Continue/Exit window
'             1 = I Agree/I agree, never show again/Exit window
Global splash_mode As Integer

'splash_button_pressed:
'1 = Continue or I Agree
'2 = I agree, never show again
'3 = Exit
Global splash_button_pressed As Integer

Global MAIN_APP_PATH As String

Global Ws1 As Workspace
Global DB_Main As database
Global RS_Main As Recordset

Global booActLikeBeta As Boolean



Const MainMod_declarations_end = True


'Function get_program_version_with_build_info_VB4( _
'    IsOnFrontWindow As Boolean) _
'    As String
'Dim ver As String
'  ver = ver & Trim$(App.Major)
'  ver = ver & "." & Trim$(App.Minor)
'  If (IsOnFrontWindow = False) Then
'    ver = ver & "." & Trim$(App.Revision)
'  End If
'  get_program_version_with_build_info_VB4 = ver
'End Function
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


Function frmSplash_Run() As Integer
Dim tpath$
Dim tstr$
Dim must_read_disclaimer As Integer

  '''SET UP INI FILE PATH.
  ''tpath$ = GetWindowsDir() & ProgramIniFile$
  
  'SHOW THE CONTINUE/EXIT FRONT WINDOW.
  splash_mode = 0
  splash_button_pressed = 0
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
  'tstr$ = INI_GetSetting0(fn_OldFileList, "disclaimer", "has_read_disclaimer")
  'tstr$ = ini_getsetting(INI_FileName, INI_ProgramType, "has_seen_disclaimer")
  tstr$ = INI_Getsetting("has_seen_disclaimer")
  If (tstr$ = "1") Then
    must_read_disclaimer = False
  End If
  
  ''''If (must_read_disclaimer) Then
  If (1 = 0) Then
    'SHOW THE DISCLAIMER WINDOW.
    splash_mode = 1
    splash_button_pressed = 0
    frmSplash.Show 1
    Select Case splash_button_pressed
      Case 1:         'Hit I Agree
        'DO NOTHING.
      Case 2:         'Hit I agree, never show again
        ''Call ini_putsetting0(fn_INI_path, "disclaimer", "has_read_disclaimer", "1")
        'Call ini_putsetting0(fn_OldFileList, "disclaimer", "has_read_disclaimer", "1")
        Call INI_PutSetting("has_seen_disclaimer", "1")
      Case 3:         'Hit Exit
        End
    End Select
  End If

  frmSplash_Run = True

End Function


Sub ChangeDir_Exes()
  ChDrive MAIN_APP_PATH
  ChDir MAIN_APP_PATH & "\EXES"
End Sub
Sub ChangeDir_Main()
  ChDrive MAIN_APP_PATH
  ChDir MAIN_APP_PATH
End Sub


Sub Main()
Dim fn_Misc1 As String
Dim fpath_INI As String
  '
  ' SET UP MAIN APP PATH VARIABLE.
  '
  If (File_IsExists(App.Path & "\debug_in_vb6.txt")) Then
    'FOR DEBUGGING IN THE VB5 ENVIRONMENT.
    MAIN_APP_PATH = "X:\etdot10\code\stepp\vb6"
    ChDrive MAIN_APP_PATH
    ChDir MAIN_APP_PATH
  Else
    'DO NOTHING.
    MAIN_APP_PATH = App.Path
  End If
  '
  ' VERIFY THAT PATHS ARE PROPERLY SET UP.
  '
  fn_Misc1 = App.Path & "\dbase\misc1.dat"
  If (File_IsExists(fn_Misc1)) Then
    'DO NOTHING; THIS IS OKAY.
  Else
    Call Show_Error("The file `" & fn_Misc1 & "` is missing.  " & _
        "Therefore the software must have been improperly installed.  " & _
        "Recommendation: Check the `Start In` path specified in the " & _
        "program icon, or else perform a re-install of the software.")
    End
  End If
  booActLikeBeta = False
  If (File_IsExists(MAIN_APP_PATH & "\actlikebeta.txt")) Then
    booActLikeBeta = True
  End If
  '
  ' READ IN THE LICENSE FILE DATA.
  '
  If (TURN_LICENSING_OFF = False) Then
    Call LicFileData_Read(Global_fpath_dir_CPAS)
  End If
  fpath_INI = Global_fpath_dir_CPAS & "\DBASE"
  '
  ' SET UP INI FILENAME FOR VARIOUS USER PREFERENCES, INCLUDING LAST-FEW-FILES LISTS.
  '
  fn_OldFileList = fpath_INI & "\" & fn_INI_name
  '
  'MISC INITIALIZATIONS.
  '
  ''''Call ini_initializethisprogram("stepp")
  
  If (fileexists(App.Path & "\help\stepp.hlp")) Then App.HelpFile = App.Path & "\help\stepp.hlp"
  
    steppPath = App.Path
    SaveAndLoadPath = App.Path
    Database_Path = App.Path + "\dbase"
    
    'ChDrive Database_Path
    'ChDir Database_Path
    ''' THE ENCRYPTION IS KEPT IN DEMOMODE.BAS
    ''If (SecureDBMode) Then
    ''    On Error GoTo Security_Database
    ''    SetDefaultWorkspace decrypt_string(Encrypted_User_Name), decrypt_string(Encrypted_User_Password)
    ''    On Error Resume Next
    ''End If
    'ChDrive steppPath
    'ChDir steppPath

    Call LoadErrorMessages
    Call InitializeHierarchy
    Call InitializeBIPdbHierarchy

    'Initialize Number of Chemicals currently selected to zero
    NumSelectedChemicals = 0

    '
    ' OPEN THE PASSWORD-PROTECTED MAIN DATABASE.
    '
    Set Ws1 = Workspaces(0)
    'NOTE: THE FOLLOWING OpenDatabase() COMMAND MUST BE
    'SPECIFIED EXACTLY AS-IS, OR ELSE IT WILL FAIL.
    Set DB_Main = _
        Ws1.OpenDatabase(Database_Path + "\stepp_db.mdb", _
              True, _
              False, _
              ";pwd=" & decrypt_string(Encrypted_User_Password))
    
    ''Open the database (Microsoft Access)
    'If (DemoMode) Then
    '    contam_prop_form!Data1.DatabaseName = Database_Path + "\demo_db.mdb"
    'Else
    '    contam_prop_form!Data1.DatabaseName = Database_Path + "\stepp_db.mdb"
    'End If
    'contam_prop_form!Data1.RecordSource = "SELECT * FROM [Names (Master)]"
    'contam_prop_form!Data1.Refresh
  
    read_blist_file


    'Temporarily set initially selected chemical to
    'Carbon Tetrachloride
    contam_prop_form!contam_combo.ListIndex = 5
    contam_prop_form!contam_combo.TopIndex = contam_prop_form!contam_combo.ListIndex
    ''''contam_prop_form!contam_combo.Selected(0) = True
    
    'Load all property forms to save time when SHOWing them later
    Load contam_prop_form

    Load aqsol_form
    Load gas_diff_form
    Load hc_form
    Load Infinite_dilution_form
    Load ldens_form
    Load liquid_diff_form
    Load molar_vol_form
    Load mv_nbp_form
    Load mwt_form
    Load nbp_form
    Load octanol_form
    Load rindex_form
    Load vp_form
    Load frmWaterDensity
    Load frmWaterViscosity
    Load frmWaterSurfaceTension
    Load frmAirDensity
    Load frmAirViscosity


  
  'Call ini_initializethisprogram("asap")
  ''---- Setup helpfiles
  'If (fileexists(app.Path & "\help\asap.hlp")) Then app.HelpFile = app.Path & "\help\asap.hlp"
  'ChDrive app.Path
  'ChDir app.Path
  'SaveAndLoadPath = app.Path
  '
  ''Initialize Default Power Variables
  'Scr1.Power.BlowerEfficiency = 35#
  'Scr1.Power.PumpEfficiency = 80#
  'Scr2.Power.BlowerEfficiency = 35#
  'Scr2.Power.PumpEfficiency = 80#
  '
  'bub.Power.BlowerEfficiency = 35#
  'bub.Power.TankWaterDepth = 4#
  'bub.Power.NumberOfBlowersInEachTank = 1
  '
  'ReadMainPackingDB
  'ReadUserPackingDB
  '
  'NL = Chr$(13) & Chr$(10)

  'LOAD THE SPLASH WINDOW.
  If (frmSplash_Run() = False) Then
    End
  End If
  
  '
  ' SHOW THE DEMO WINDOW.
  '
  If (IsThisADemo() = True) Then
    Call frmDemo.frmDemo_GO
  End If
  '
  ' LOAD THE MAIN WINDOW.
  '
  contam_prop_form.Show
  Exit Sub
  
Security_Database:
Dim temp As String, Error_Code As Integer
  Error_Code = Err
  temp = "Error " & Format$(Error_Code, "0") & " : " & error$(Error_Code)
  If Err = 3024 Then
      MsgBox "The File SYSTEM.MDA is missing.  The database is not accessible.  The program will be terminated."
  Else
      MsgBox "Error while checking the security system.  " & Chr$(13) & temp & Chr$(13) & "The database is not accessible.  The program will be terminated."
  End If

  Resume ExitProgram

ExitProgram:
  End

End Sub


'NOTE: THIS FUNCTION WORKS EQUALLY WELL ON
'EITHER FILES OR DIRECTORIES.
Function File_IsExists(fn As String) As Boolean
Dim Dummy As Long
  On Error GoTo err_File_IsExists
  Dummy = GetAttr(fn)   'TRIGGERS ERROR IF FILE DOES NOT EXIST.
  File_IsExists = True
  Exit Function
exit_err_File_IsExists:
  File_IsExists = False
  Exit Function
err_File_IsExists:
  Resume exit_err_File_IsExists
End Function
Function FileExists0(fn As String) As Boolean
  FileExists0 = File_IsExists(fn)
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


