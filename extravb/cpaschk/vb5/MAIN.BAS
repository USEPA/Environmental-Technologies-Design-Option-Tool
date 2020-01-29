Attribute VB_Name = "MainModule"
Option Explicit

Global Const LICFILE_SoftwareTitle = "CPAS Software"
Global Const LICFILE_ExpireContactInfo = "Eric J. Oman at (906) 487-1624 or David W. Hand at (906) 487-2777"

Global Const LICFILE_StudentCheckFile = "R:\esbeam\data\tack\esbeam.tck"
Global Const LICFILE_LicName = "CPAS.LIC"
Global Const LICFILE_NewLicInfo = "NEWLIC.X"
Global Const LICFILE_GoodSerialNumber = "OKNUM.X"
Global Const LICFILE_BadSerialNumber = "BADNUM.X"
Global Const LICFILE_GoodLicenseFile = "GO.X"
Global Const LICFILE_BadLicenseFile = "EXIT.X"
Global Const LICFILE_DATE_NEVER = "NEVER"
Global Const LICFILE_ExtraCheckFile = "MTCHK.LIC"
Global Const LICFILE_ExtraCheckFile_Text = "IOW2EK4FV832"

Global Const DEBUG_Pause_Before_Exit = 0   '1

'Global Const LICFILE_MAXSIZE = 5823
'
'Global Const LICFILE_Z_EXPIRATIONDATE = 1492
'Global Const LICFILE_Z_PROGRAMKEY_ADS = 102
'Global Const LICFILE_Z_PROGRAMKEY_ASAP = 175
'Global Const LICFILE_Z_PROGRAMKEY_STEPP = 293
'Global Const LICFILE_Z_RELEASETYPE = 417
'Global Const LICFILE_Z_SERIALNUMBER = 2032
'Global Const LICFILE_Z_USERCOMPANY = 1004
'Global Const LICFILE_Z_USERNAME = 1602
'Global Const LICFILE_Z_VERSIONCODE = 502
'Global Const LICFILE_Z_VERSIONTYPE = 774
'Global Const LICFILE_ZZ_LASTEXECUTIONDATE = 1232
'Global Const LICFILE_ZZ_LASTEXECUTIONTIME = 1741

Global Const lfdt_order_Z_SERIALNUMBER = 2
Global Const lfdt_order_Z_USERNAME = 3
Global Const lfdt_order_Z_USERCOMPANY = 4
Global Const lfdt_order_ZZ_LASTEXECUTIONDATE = 5
Global Const lfdt_order_ZZ_LASTEXECUTIONTIME = 6
Global Const lfdt_order_ZZ_NUMPROGRAMKEYS = 7

Type LicFile_Data_Type
  Z_SERIALNUMBER As String
  Z_USERNAME As String
  Z_USERCOMPANY As String
  ZZ_LASTEXECUTIONDATE As String
  ZZ_LASTEXECUTIONTIME As String
  ZZ_NUMPROGRAMKEYS As Integer
End Type

Global Const pkdt_order_Z_PROGRAMKEY = 0
Global Const pkdt_order_Z_EXPIRATIONDATE = 1
Global Const pkdt_order_Z_RELEASETYPE = 2
Global Const pkdt_order_Z_VERSIONCODE = 3
Global Const pkdt_order_Z_VERSIONTYPE = 4

Type ProgramKey_Data_Type
  Z_PROGRAMKEY As String
  Z_EXPIRATIONDATE As String
  Z_RELEASETYPE As String
  Z_VERSIONCODE As String
  Z_VERSIONTYPE As String
End Type

Global fn_ControlFile As String

Global MAIN_APP_PATH As String





Const MainModule_declarations_end = True


Sub ExitCode(arg_OutputDir As String, fn_nopath As String)
Dim f As Integer
Dim fn_full As String
  f = FreeFile
  fn_full = arg_OutputDir & "\" & fn_nopath
  Open fn_full For Output As #f
  Print #f, "0"
  Close #f
  If (DEBUG_Pause_Before_Exit = 1) Then
    MsgBox "Exiting with code `" & arg_OutputDir & "`,`" & fn_nopath & "` ..."
  End If
  End
End Sub
Sub ExitError(arg_OutputDir As String, fn_nopath As String)
  MsgBox "There appears to be an error in your licensing configuration."
  Call ExitCode(arg_OutputDir, fn_nopath)
End Sub
Sub ShowError(msg As String)
  MsgBox msg, vbExclamation, LICFILE_SoftwareTitle
End Sub


Sub Main()
Dim cmdline As String
Dim bad_params As Integer
Dim arg1 As String
Dim arg2 As String
Dim arg3 As String
Dim arg_CpasDir As String
Dim arg_ProgramKey As String
Dim num_args As Integer
Dim arg_AllButFirst As String
Dim arg_AllButFirstAndLast As String
Dim arg_Last As String
Dim arg_PastComma As String
Dim arg_Remaining As String
Dim i As Integer
  '
  'SAMPLE COMMAND PARAMETERS:
  '  -CREATE_FILE i:\cpastest
  '  -DISPLAY_ALL i:\cpastest
  '  -GET_INFO i:\cpastest ads
  '
  'ADDITIONAL SAMPLE COMMAND PARAMETERS:
  '  -CREATE_FILE x:\etdot10\extravb\cpaschk\test
  '  -DISPLAY_ALL x:\etdot10\extravb\cpaschk\test
  '  -GET_INFO x:\etdot10\extravb\cpaschk\test ads
  '
  'NEW SAMPLES FOR FILENAMES-WITH-SPACES CAPABILITY:
  '  -CREATE_FILE C:\PROGRAM FILES\CPAS10
  '  -DISPLAY_ALL C:\PROGRAM FILES\CPAS10
  '  -GET_INFO C:\PROGRAM FILES\CPAS10 ADS
  '  -GET_INFO C:\PROGRAM FILES\CPAS10 ADS ,C:\PROGRAM FILES\CPAS10\ASAP
  '
  '  -GET_INFO D:\PROGRAM FILES\ETDOT10 ASAP ,X:\CPAS10\CODE\ASAP\VB3
  '
  ''''Call LicFile_AppendGarbage("i:\cpastest\test.txt", 10)
  ''''End
  '
  'SET UP MAIN APP PATH VARIABLE.
  '
  If (FileExists(App.Path & "\debug_in_vb5.txt")) Then
    'FOR DEBUGGING IN THE VB5 ENVIRONMENT.
    MAIN_APP_PATH = "X:\etdot10\extravb\cpaschk\vb5"
    ChDrive MAIN_APP_PATH
    ChDir MAIN_APP_PATH
  Else
    'DO NOTHING.
    MAIN_APP_PATH = App.Path
  End If
  '
  'MAIN CODE SEQUENCE.
  '
  Call get_win_paths
  'fn_ControlFile = WinPathWindows$ & LICFILE_LicName
  cmdline = Trim$(UCase$(CStr(Command$)))
  num_args = Parser_GetNumArgs(" ", cmdline)
  bad_params = True
  If (num_args >= 1) Then
    'GET FIRST ARGUMENT.
    Call Parser_GetArg(" ", cmdline, 1, arg1)
    'GET ARGUMENT PAST COMMA (IF ANY).
    If (InStr(cmdline, ",") <> 0) Then
      i = InStr(cmdline, ",")
      arg_Remaining = Left$(cmdline, i - 1)
      arg_PastComma = Trim$(Right$(cmdline, Len(cmdline) - Len(arg_Remaining) - 1))
      arg_Remaining = Trim$(arg_Remaining)
    Else
      arg_Remaining = cmdline
    End If
    'GET REST OF ARGUMENTS (ALL BUT THE FIRST).
    If (Len(arg_Remaining) > Len(arg1)) Then
      arg_AllButFirst = Right$(arg_Remaining, Len(arg_Remaining) - Len(arg1))
      arg_AllButFirst = Trim$(arg_AllButFirst)
    Else
      arg_AllButFirst = ""
    End If
    'GET LAST ARGUMENT.
    arg_Last = ""
    i = Len(arg_AllButFirst)
    Do While (1 = 1)
      If (Mid$(arg_AllButFirst, i, 1) = " ") Then Exit Do
      arg_Last = Mid$(arg_AllButFirst, i, 1) & arg_Last
      i = i - 1
      If (i < 1) Then Exit Do
    Loop
    'GET ALL BUT FIRST AND LAST ARGUMENTS.
    arg_AllButFirstAndLast = ""
    If (arg_Last <> "") And (arg_AllButFirst <> "") Then
      If (Len(arg_Last) < Len(arg_AllButFirst)) Then
        arg_AllButFirstAndLast = _
            Left$(arg_AllButFirst, _
            Len(arg_AllButFirst) - Len(arg_Last))
        arg_AllButFirstAndLast = Trim$(arg_AllButFirstAndLast)
      End If
    End If
    'RUN COMMAND (BASED ON FIRST ARGUMENT).
    '=============  -CREATE_FILE  ===============================================================================
    If (Trim$(UCase$(arg1)) = Trim$(UCase$("-CREATE_FILE"))) Then
      If (arg_AllButFirst <> "") Then
        arg_CpasDir = Trim$(arg_AllButFirst)
        Call Do_Create_File_vX(arg_CpasDir)
      End If
      'If (num_args = 2) Then
      '  Call Parser_GetArg(" ", cmdline, 2, arg2)
      '  arg_CpasDir = arg2
      '  bad_params = False
      '  Call Do_Create_File_vX(arg_CpasDir)
      'End If
    End If
    ''=============  -GET_INFOX  ===============================================================================
    'If (Trim$(UCase$(arg1)) = Trim$(UCase$("-GET_INFOX"))) Then
    '  If (arg_Last <> "") And (arg_AllButFirstAndLast <> "") And (arg_PastComma <> "") Then
    '    arg_CpasDir = Trim$(arg_AllButFirstAndLast)
    '    arg_ProgramKey = Trim$(arg_Last)
    '    Call Do_Get_Info_vX(arg_CpasDir, arg_ProgramKey, arg_PastComma)
    '  End If
    'End If
    '=============  -GET_INFO  ===============================================================================
    If (Trim$(UCase$(arg1)) = Trim$(UCase$("-GET_INFO"))) Then
      If (arg_Last <> "") And (arg_AllButFirstAndLast <> "") And (arg_PastComma <> "") Then
        arg_CpasDir = Trim$(arg_AllButFirstAndLast)
        arg_ProgramKey = Trim$(arg_Last)
        Call Do_Get_Info_vX(arg_CpasDir, arg_ProgramKey, arg_PastComma)
      End If
      If (arg_Last <> "") And (arg_AllButFirstAndLast <> "") And (arg_PastComma = "") Then
        arg_CpasDir = Trim$(arg_AllButFirstAndLast)
        arg_ProgramKey = Trim$(arg_Last)
        Call Do_Get_Info_vX(arg_CpasDir, arg_ProgramKey, "")
      End If
      'If (num_args = 3) Then
      '  Call Parser_GetArg(" ", cmdline, 2, arg2)
      '  Call Parser_GetArg(" ", cmdline, 3, arg3)
      '  arg_CpasDir = arg2
      '  arg_ProgramKey = arg3
      '  bad_params = False
      '  Call Do_Get_Info_vX(arg_CpasDir, arg_ProgramKey)
      'End If
    End If
    '=============  -DISPLAY_ALL  ===============================================================================
    If (Trim$(UCase$(arg1)) = Trim$(UCase$("-DISPLAY_ALL"))) Then
      If (arg_AllButFirst <> "") Then
        arg_CpasDir = Trim$(arg_AllButFirst)
        Call Do_Display_All_vX(arg_CpasDir)
      End If
      'If (num_args = 2) Then
      '  Call Parser_GetArg(" ", cmdline, 2, arg2)
      '  arg_CpasDir = arg2
      '  bad_params = False
      '  Call Do_Display_All_vX(arg_CpasDir)
      'End If
    End If
  End If
  End
End Sub



