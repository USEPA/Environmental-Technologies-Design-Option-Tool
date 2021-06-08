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

'''''Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathName" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
''''Declare Function GetShortPathName Lib "c:\winnt\system32\kernel32.dll" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Global MAIN_APP_PATH As String

Global booActLikeBeta As Boolean




Const MainMod_declarations_end = True


''''Function get_program_version_with_build_info_VB4( _
''''    IsOnFrontWindow As Boolean) _
''''    As String
''''Dim ver As String
''''  ver = ver & Trim$(App.Major)
''''  ver = ver & "." & Trim$(App.Minor)
''''  If (IsOnFrontWindow = False) Then
''''    ver = ver & "." & Trim$(App.Revision)
''''  End If
''''  get_program_version_with_build_info_VB4 = ver
''''End Function
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
  tstr$ = INI_Getsetting("has_seen_disclaimer")
  If (tstr$ = "1") Then
    must_read_disclaimer = False
  End If
  
  If (1 = 0) Then
  ''''If (must_read_disclaimer) Then
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
''''MsgBox "ChangeDir_Main: CurDir = " & CurDir
End Sub


Sub Do_The_DLL_Test()
'Dim DG As Double
'Dim TEMP As Double
'Dim PRES As Double
'  TEMP = 298.15
'  PRES = 1#
'  Call AIRDENS(DG, TEMP, PRES)
'  MsgBox "DG = " & Trim$(Str$(DG))

Dim CS As Double
Dim VQ As Double
Dim HC As Double
Dim CI As Double
Dim CE As Double
  VQ = 1#
  HC = 1#
  CI = 100#
  CE = 5#
  Call GETCSPT(CS, VQ, HC, CI, CE)
  MsgBox "CS = " & Trim$(Str$(CS))
 
'          CS = (1# / (VQ * HC)) * (CI - CE)

End Sub


Sub Main()
Dim fn_Misc1 As String
Dim fpath_INI As String
  '
  ' SET UP MAIN APP PATH VARIABLE.
  '
  If (File_IsExists(App.Path & "\debug_in_vb6.txt")) Then
    'FOR DEBUGGING IN THE VB5 ENVIRONMENT.
    MAIN_APP_PATH = "X:\etdot10\code\asap\vb6"
    ChDrive MAIN_APP_PATH
    ChDir MAIN_APP_PATH
  Else
    'DO NOTHING.
    MAIN_APP_PATH = App.Path
'MsgBox "(1.) CurDir = " & CurDir
'ChDrive MAIN_APP_PATH
'ChDir MAIN_APP_PATH
'MsgBox "(2.) CurDir = " & CurDir
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

'MsgBox "Test (1a)"
'Call Do_The_DLL_Test
'MsgBox "Test (1b)"

  ''temp
  'ChDir "d:\program files\etdot10\asap"
  'ChDrive "d:\program files\etdot10\asap"
  'Dim RetVal As Long
  'Dim lpszLongPath As String * 120
  'Dim lpszShortPath As String * 120
  'Dim cchBuffer As Long
  'lpszLongPath = "d:\program files\etdot10\asap"
  'RetVal = GetShortPathName(ByVal lpszLongPath, ByVal lpszShortPath, ByVal cchBuffer)
  'End
   
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
  '
  ' SET UP INI FILENAME FOR VARIOUS USER PREFERENCES, INCLUDING LAST-FEW-FILES LISTS.
  '
  fn_OldFileList = fpath_INI & "\" & fn_INI_name
  '
  ' MISC INITIALIZATIONS.
  '
  ''''Call ini_initializethisprogram("asap")
  '
  '---- Setup helpfiles
  '
  If (fileexists(MAIN_APP_PATH & "\help\asap.hlp")) Then
    App.HelpFile = MAIN_APP_PATH & "\help\asap.hlp"
  End If
  Call ChangeDir_Main
  'ChDrive App.Path
  'ChDir App.Path
  SaveAndLoadPath = App.Path
  
  'Initialize Default Power Variables
  scr1.Power.BlowerEfficiency = 35#
  scr1.Power.PumpEfficiency = 80#
  Scr2.Power.BlowerEfficiency = 35#
  Scr2.Power.PumpEfficiency = 80#
  
  bub.Power.BlowerEfficiency = 35#
  bub.Power.TankWaterDepth = 4#
  bub.Power.NumberOfBlowersinEachTank = 1

  ReadMainPackingDB
  ReadUserPackingDB
  
  NL = Chr$(13) & Chr$(10)
  '
  ' LOAD THE SPLASH WINDOW.
  '
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
  frmMainMenu.Show
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


Sub Launch_ASAP_HLP_File()
Dim fn_This As String
  fn_This = MAIN_APP_PATH & "\help\asap.hlp"
  If (fileexists(fn_This) = False) Then
    Call Show_Message("The file `" & fn_This & "` is missing.")
    Exit Sub
  End If
  Call LaunchFile_General("", fn_This)
  'Call LaunchFile_General("", MAIN_APP_PATH & "\help\asap.hlp")
End Sub


Sub Launch_ASAP_mnuHelp_Item( _
    Index As Integer)
Dim fn_This As String
  Select Case Index
'    Case 5:       'CONTENTS.
'      SendKeys "{F1}", True
    Case 5:       'ONLINE HELP.
      Call Launch_ASAP_HLP_File
    Case 6:       'ONLINE MANUAL.
      ''''fn_This = MAIN_APP_PATH & "\help\asap.pdf"
      fn_This = MAIN_APP_PATH & "\help\readme.doc"
      If (fileexists(fn_This) = False) Then
        Call Show_Message("The file `" & fn_This & "` is missing.")
        Exit Sub
      End If
      Call ShellExecute_LocalFile(fn_This)
      ''''Call LaunchFile_General("", fn_This)
      '''''Call LaunchFile_General("", MAIN_APP_PATH & "\help\asap.pdf")
    Case 7:       'MANUAL PRINTING INSTRUCTIONS.
      fn_This = Global_fpath_dir_CPAS & "\dbase\printing.txt"
      If (fileexists(fn_This) = False) Then
        Call Show_Message("The file `" & fn_This & "` is missing.")
        Exit Sub
      End If
      Call Launch_Notepad(fn_This)
    Case 10:      'VIEW VERSION HISTORY.
      fn_This = App.Path & "\dbase\readme.txt"
      If (fileexists(fn_This) = False) Then
        Call Show_Message("The file `" & fn_This & "` is missing.")
        Exit Sub
      End If
      Call Launch_Notepad(fn_This)
    Case 20:      'VIEW DISCLAIMER.
      'SHOW THE DISCLAIMER WINDOW.
      splash_mode = 101
      splash_button_pressed = 0
      frmSplash.Show 1
    Case 30:      'TECHNICAL ASSISTANCE PROVIDED BY.
      frmTechAssistance.Show 1
    Case 200:     'ABOUT.
      frmAbout.Show 1
  End Select
End Sub


Sub Error_Unavailable_File( _
    fn_This As String, _
    Model_Type As String)
  Call Show_Error("The file `" & fn_This & "` does not exist. " & _
      "A valid version of this file must exist in order to run " & _
      "the " & Model_Type & " model. Please re-install the software.")
End Sub


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


