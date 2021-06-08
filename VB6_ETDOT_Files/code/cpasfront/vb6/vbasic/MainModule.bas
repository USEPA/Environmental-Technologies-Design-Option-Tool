Attribute VB_Name = "MainModule"
Option Explicit

Global frmMain_MODE As Integer
Global Const frmMain_MODE_DESIGN = 1
Global Const frmMain_MODE_USER = 2

Global fpath_dbase As String
Global fpath_Icons As String
Global fpath_Backgrounds As String
Global Const fn_Short_MainDataFile = "MAIN.DAT"
Global fn_Full_MainDataFile As String

'splash_mode: 0 = Continue/Exit window
'             1 = I Agree/I agree, never show again/Exit window
Global splash_mode As Integer

'splash_button_pressed:
'1 = Continue or I Agree
'2 = I agree, never show again
'3 = Exit
Global splash_button_pressed As Integer

Global ALLOW_DESIGN_MODE As Boolean

Global MAIN_APP_PATH As String




Const MainModule_declarations_end = 0


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
  'tstr$ = INI_GetSetting0(fn_INI_path, "disclaimer", "has_read_disclaimer")
  tstr$ = INI_GetSetting0(fn_OldFileList, "disclaimer", "has_read_disclaimer")
  
  'tstr$ = ini_getsetting("has_seen_disclaimer")
  If (tstr$ = "1") Then
    must_read_disclaimer = False
  End If
  
  If (must_read_disclaimer) Then
    'SHOW THE DISCLAIMER WINDOW.
    splash_mode = 1
    splash_button_pressed = 0
    frmSplash.Show 1
    Select Case splash_button_pressed
      Case 1:         'Hit I Agree
        'DO NOTHING.
      Case 2:         'Hit I agree, never show again
        ''Call ini_putsetting0(fn_INI_path, "disclaimer", "has_read_disclaimer", "1")
        'Call ini_putsetting0(fn_INI_path, "disclaimer", "has_read_disclaimer", "1")
        Call ini_putsetting0(fn_OldFileList, "disclaimer", "has_read_disclaimer", "1")
        'Call ini_putsetting("has_seen_disclaimer", "1")
      Case 3:         'Hit Exit
        End
    End Select
  End If

  frmSplash_Run = True

End Function


Sub Main()
Dim fn_misc1 As String
Dim fpath_INI As String

  AppName_For_Display_Short = "CPAS Desktop"
  AppName_For_Display_Long = "CPAS Desktop"
  '
  ' DO NOT ALLOW DESIGN MODE UNLESS THIS "SECRET" COMMAND-LINE
  ' PARAMETER IS USED.
  '
  If (Trim$(UCase$(Command$)) = "EJOMAN") Then
    ALLOW_DESIGN_MODE = True
  Else
    ALLOW_DESIGN_MODE = False
  End If
  '
  ' READ THE LICENSING FILE.
  '
  Call LicFileData_Read(Global_fpath_dir_CPAS)
  '
  ' VERIFY THAT PATHS ARE PROPERLY SET UP.
  '
  MAIN_APP_PATH = App.Path
  fpath_dbase = MAIN_APP_PATH & "\dbase"
  fpath_Icons = MAIN_APP_PATH & "\dbase\icons"
  fpath_Backgrounds = MAIN_APP_PATH & "\dbase\backs"
  fn_misc1 = fpath_dbase & "\misc1.dat"
  If (File_IsExists(fn_misc1)) Then
    'DO NOTHING; THIS IS OKAY.
  Else
    Call Show_Error("The file `" & fn_misc1 & "` is missing.  " & _
        "Therefore the software must have been improperly installed.  " & _
        "Recommendation: Check the `Start In` path specified in the " & _
        "program icon, or else perform a re-install of the software.")
    End
  End If
  ''
  ''READ IN THE LICENSE FILE DATA.
  ''
  'Call LicFileData_Read(LicFileLocation)
  'Select Case LicFileLocation
  '  Case LICFILELOCATION_WIN:
      'fpath_INI = GetWindowsDir()
  '  Case LICFILELOCATION_APPPATH:
  '    fpath_INI = App.Path
  'End Select
  '
  ' SET UP INI FILENAME.
  '
  fpath_INI = Global_fpath_dir_CPAS & "\DBASE"
  fn_OldFileList = fpath_INI & "\" & fn_INI_name
  '
  ' SET UP MAIN DATAFILE FILENAME.
  '
  fn_Full_MainDataFile = App.Path & "\dbase\" & fn_Short_MainDataFile
  '
  ' LOAD THE SPLASH WINDOW.
  '
  If (frmSplash_Run() = False) Then
    End
  End If
  '
  ' SHOW THE MAIN FORM, MODE-LESS.
  '
  frmMain.Show
  '
  ' EXIT OUT OF HERE.
  '
End Sub


Function get_program_version_with_build_info_VB4() As String
Dim ver As String
Dim capped As String
  'capped = LCase$(Trim$(lfd.Z_RELEASETYPE))
  'If (Len(capped) >= 1) Then
  '  Mid$(capped, 1, 1) = UCase$(Mid$(capped, 1, 1))
  'End If
  'ver = lfd.Z_VERSIONCODE & " (" & capped & ")"
  ver = ver & Trim$(App.Major) & "."
  ver = ver & Trim$(App.Minor) & "."
  ver = ver & Trim$(App.Revision)
  get_program_version_with_build_info_VB4 = ver
End Function


Sub DebugOutputFile(OutStr As String)
Dim f As Integer
  f = FreeFile
  Open "c:\bug.txt" For Append As #f
  Print #f, Now, OutStr
  Close #f
End Sub
