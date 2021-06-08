Attribute VB_Name = "LicdataMod"
Option Explicit

Global Const TURN_LICENSING_OFF = False

Global Const AppProgramKey = "STEPP"
Global Const AppCopyrightYears = "1993-1998"
Global Const AppName = "StEPP"
Global AppWillExpire As Integer     'true/false
Global AppExpireYear As Integer
Global AppExpireMonth As Integer
Global AppExpireDay As Integer
Global Global_fpath_dir_CPAS As String

Global Const OSTYPE_WIN95 = 1
Global Const OSTYPE_WINNT = 2
Global Const LAUNCHFILEVIA_IS_DEBUG_MODE_ON = False

'Global Const LICFILE_GetInfoProgram = "MTCHK.EXE"
Global Const LICFILE_GetInfoProgram = "CPASCHK.EXE"
Global Const LICFILE_GetInfoProgramParams = "-GET_INFO"

'Global Const LICFILE_LicName = "ETDOT10.LIC"
Global Const LICFILE_LicName = "CPAS.LIC"
'Global Const LICFILE_NewLicInfo = "MTNEWLIC.X"
'Global Const LICFILE_GoodSerialNumber = "OKNUM.X"
'Global Const LICFILE_BadSerialNumber = "BADNUM.X"
Global Const LICFILE_GoodLicenseFile = "GO.X"
Global Const LICFILE_BadLicenseFile = "EXIT.X"

Type LicFile_Data_Type
  'Z_PROGRAMKEY_ADS As String
  'Z_PROGRAMKEY_ASAP As String
  'Z_PROGRAMKEY_STEPP As String
  Z_SERIALNUMBER As String
  Z_USERNAME As String
  Z_USERCOMPANY As String
  Z_PROGRAMKEY As String
  Z_EXPIRATIONDATE As String
  Z_RELEASETYPE As String
  Z_VERSIONCODE As String
  Z_VERSIONTYPE As String
  'ZZ_LASTEXECUTIONDATE As String
  'ZZ_LASTEXECUTIONTIME As String
End Type
Global lfd As LicFile_Data_Type

Function get_expiration_info() As String
  Select Case Trim$(UCase$(lfd.Z_VERSIONTYPE))
    Case Trim$(UCase$("VER_INTERNAL_STUDENT")):
      get_expiration_info = "No Expiration Date (Student Copy)"
    Case Trim$(UCase$("VER_WONT_EXPIRE")):
      get_expiration_info = "No Expiration Date (Professional Copy)"
    Case Else:
      get_expiration_info = "Expires on " & Trim$(Str$(AppExpireMonth)) & "/" & Trim$(Str$(AppExpireDay)) & "/" & Trim$(Str$(AppExpireYear))
  End Select
End Function

Function get_program_version_with_build_info() As String
Dim ver As String
Dim capped As String
  capped = LCase$(Trim$(lfd.Z_RELEASETYPE))
  If (Len(capped) >= 1) Then
    Mid$(capped, 1, 1) = UCase$(Mid$(capped, 1, 1))
  End If
  ver = lfd.Z_VERSIONCODE & " (" & capped & ")"
  'ver = ver & Trim$(App.Major) & "."
  'ver = ver & Trim$(App.Minor) & "."
  'ver = ver & Trim$(App.Revision)
  get_program_version_with_build_info = ver
End Function

'RETURNS:
'    TRUE = SUCCEEDED.
'    FALSE = FAILED.
Function LaunchFileViaStartMethod(fn_Dir As String, fn_File As String) As Integer
Dim RetValBool As Integer
  RetValBool = LaunchFileViaStartMethod_0(Trim$(fn_Dir), Trim$(fn_File), OSTYPE_WINNT)
  If (Not RetValBool) Then
    RetValBool = LaunchFileViaStartMethod_0(Trim$(fn_Dir), Trim$(fn_File), OSTYPE_WIN95)
  End If
  LaunchFileViaStartMethod = RetValBool
End Function

'RETURNS:
'    TRUE = SUCCEEDED.
'    FALSE = FAILED.
Function LaunchFileViaStartMethod_0(fn_Dir As String, fn_File As String, OSTYPE As Integer) As Integer
Dim RetVal As Integer
Dim cmdline As String
    
  On Error GoTo err_LaunchFileViaStartMethod_0
  
  If (Trim$(fn_Dir) <> "") Then
    ChDir Trim$(fn_Dir)
  End If
  Select Case OSTYPE
    Case OSTYPE_WIN95:
      'CMDLINE = "command.com /c start " & Trim$(fn_File)
      cmdline = "command.com /c " & Trim$(fn_File)
    Case OSTYPE_WINNT:
      'CMDLINE = "cmd /c start " & Trim$(fn_File)
      cmdline = "cmd /c " & Trim$(fn_File)
  End Select
  'If (LAUNCHFILEVIA_IS_DEBUG_MODE_ON) Then
    MsgBox "CmdLine = `" & cmdline & "`"
  'End If
  RetVal = 0 * Shell(cmdline, 1)
  
  LaunchFileViaStartMethod_0 = True
  Exit Function
    
exit_err_LaunchFileViaStartMethod_0:
  LaunchFileViaStartMethod_0 = False
  Exit Function
err_LaunchFileViaStartMethod_0:
  Resume exit_err_LaunchFileViaStartMethod_0
End Function

Sub LicFileData_Read(return_fpath_dir_CPAS As String)
Dim WinDir As String
Dim fn_CPASCHK As String
Dim cmdline As String
Dim time_start As Double
Dim fn_GoodLicenseFile As String
Dim fn_BadLicenseFile As String
Dim time_elapsed As Double
Dim f As Integer
Dim RetVal As Integer
Dim copy_z_expirationdate As String
Dim temp As String
Dim fn_CPASDIR_INI As String
Dim fpath_Dir_CPAS As String
Dim AnyErrors As Integer
Dim CMDLINE0 As String
Dim fn_ResultsFile As String
Dim OLD_fpath_Dir_CPAS As String

  'GET CPAS DIRECTORY NAME.
  fn_CPASDIR_INI = App.Path & "\CPASDIR.INI"
  If (Not fileexists(fn_CPASDIR_INI)) Then
    'UNABLE TO READ LICENSE FILE DATA.
    GoTo err_Cant_Read_Licensing_Data
  End If
  temp = Trim$(INI_GetSetting00(fn_CPASDIR_INI, "Directory", "CPASDIR"))
  fpath_Dir_CPAS = temp
  return_fpath_dir_CPAS = temp

  'CONVERT CPAS DIRECTORY PATH TO SHORT-FILENAME CONVENTION (IF NEEDED).
  ChDir App.Path
  ChDrive App.Path
  CMDLINE0 = "fnconv " & fpath_Dir_CPAS
  fn_ResultsFile = "shortp.x"
  If (fileexists(fn_ResultsFile)) Then
    Kill fn_ResultsFile
  End If
  RetVal = 0 * Shell(CMDLINE0, 1)
  time_start = Timer
  Do While (1 = 1)
    DoEvents
    If (fileexists(fn_ResultsFile)) Then
      'Kill fn_ResultsFile    'DELETED BELOW.
      time_start = Timer
      Do While (time_start = Timer)
        DoEvents
      Loop
      DoEvents
      Exit Do
    End If
    time_elapsed = Timer - time_start
    If (time_elapsed > 10#) Then
      'UNABLE TO READ LICENSE FILE DATA.
      GoTo err_Cant_Read_Licensing_Data
    End If
  Loop
  f = FreeFile
  OLD_fpath_Dir_CPAS = fpath_Dir_CPAS
  Open fn_ResultsFile For Input As #f
  Line Input #f, fpath_Dir_CPAS
  Close #f
  Kill fn_ResultsFile

  'CHECK ON LICENSE FILE.
  WinDir = GetWindowsDir()
  'fn_MTCHK = WinDir & "\" & LICFILE_GetInfoProgram
  fn_CPASCHK = fpath_Dir_CPAS & "\DBASE\" & LICFILE_GetInfoProgram
  'fn_CPASCHK = LICFILE_GetInfoProgram
  'If (fileexists(fn_CPASCHK)) Then
  '  'THAT'S OKAY.
  'Else
  '  'UNABLE TO READ LICENSE FILE DATA.
  '  GoTo err_Cant_Read_Licensing_Data
  'End If
  'CmdLine = LICFILE_GetInfoProgram & " " & LICFILE_GetInfoProgramParams
  'CmdLine = CmdLine & " " & fpath_dir_CPAS
  'CmdLine = CmdLine & " " & AppProgramKey
  'CMDLINE = fn_CPASCHK & " " & LICFILE_GetInfoProgramParams
  'CMDLINE = CMDLINE & " " & fpath_dir_CPAS
  'CMDLINE = CMDLINE & " " & AppProgramKey
  cmdline = Chr$(34) & fn_CPASCHK & Chr$(34) & " " & LICFILE_GetInfoProgramParams
  cmdline = cmdline & " " & fpath_Dir_CPAS
  cmdline = cmdline & " " & AppProgramKey
  cmdline = cmdline & " ," & App.Path
  ''''MsgBox CMDLINE
  'fn_GoodLicenseFile = WinDir & "\" & LICFILE_GoodLicenseFile
  'fn_BadLicenseFile = WinDir & "\" & LICFILE_BadLicenseFile
  'fn_GoodLicenseFile = fpath_dir_CPAS & "\DBASE\" & LICFILE_GoodLicenseFile
  'fn_BadLicenseFile = fpath_dir_CPAS & "\DBASE\" & LICFILE_BadLicenseFile
  fn_GoodLicenseFile = App.Path & "\" & LICFILE_GoodLicenseFile
  fn_BadLicenseFile = App.Path & "\" & LICFILE_BadLicenseFile
  time_start = Timer
  
  On Error Resume Next
  AnyErrors = False
  ChDir fpath_Dir_CPAS & "\DBASE": If (Err <> 0) Then AnyErrors = True
  ChDrive fpath_Dir_CPAS & "\DBASE": If (Err <> 0) Then AnyErrors = True
  ''''MsgBox cmdline
  RetVal = 0 * Shell(cmdline, 1): If (Err <> 0) Then AnyErrors = True
  On Error GoTo 0
  If (AnyErrors) Then
    If (False = LaunchFileViaStartMethod("", cmdline)) Then
      'UNABLE TO READ LICENSE FILE DATA.
      GoTo err_Cant_Read_Licensing_Data
    End If
  End If
  
  Do While (1 = 1)
    DoEvents
    If (fileexists(fn_GoodLicenseFile)) Then
      'Kill fn_GoodLicenseFile    'DELETED BELOW.
      time_start = Timer
      Do While (time_start = Timer)
        DoEvents
      Loop
      DoEvents
      Exit Do
    End If
    If (fileexists(fn_BadLicenseFile)) Then
      Kill fn_BadLicenseFile
      time_start = Timer
      Do While (time_start = Timer)
        DoEvents
      Loop
      DoEvents
      End
    End If
    time_elapsed = Timer - time_start
    If (time_elapsed > 10#) Then
      'UNABLE TO READ LICENSE FILE DATA.
      GoTo err_Cant_Read_Licensing_Data
    End If
  Loop
  ChDir App.Path
  ChDrive App.Path

  'READ IN LICENSE FILE INFO.
  f = FreeFile
  Open fn_GoodLicenseFile For Input As #f
  Line Input #f, lfd.Z_SERIALNUMBER
  Line Input #f, lfd.Z_USERNAME
  Line Input #f, lfd.Z_USERCOMPANY
  Line Input #f, lfd.Z_PROGRAMKEY
  Line Input #f, lfd.Z_EXPIRATIONDATE
  Line Input #f, lfd.Z_RELEASETYPE
  Line Input #f, lfd.Z_VERSIONCODE
  Line Input #f, lfd.Z_VERSIONTYPE
  Close #f
  Kill fn_GoodLicenseFile
  
  Select Case Trim$(UCase$(lfd.Z_VERSIONTYPE))
    Case Trim$(UCase$("VER_INTERNAL_STUDENT")):
      AppWillExpire = False
    Case Trim$(UCase$("VER_WONT_EXPIRE")):
      AppWillExpire = False
    Case Else:
      AppWillExpire = True
      copy_z_expirationdate = Trim$(UCase$(lfd.Z_EXPIRATIONDATE))
      copy_z_expirationdate = Parser_RemoveCharacters(" ", copy_z_expirationdate)
      If (Parser_GetNumArgs(",", copy_z_expirationdate) = 3) Then
        Call Parser_GetArg(",", copy_z_expirationdate, 1, temp)
        AppExpireMonth = CInt(Val(temp))
        Call Parser_GetArg(",", copy_z_expirationdate, 2, temp)
        AppExpireDay = CInt(Val(temp))
        Call Parser_GetArg(",", copy_z_expirationdate, 3, temp)
        AppExpireYear = CInt(Val(temp))
      End If
  End Select
  
  Exit Sub

err_Cant_Read_Licensing_Data:
  MsgBox "Unable to read licensing data.  You may need to re-install the software.", 48, AppName
  End
End Sub

Sub Parser_GetArg(sepchar As String, inline As String, ArgNum As Integer, RetStr As String)
Dim I As Integer
Dim J As Integer
  RetStr = ""
  J = 1
  For I = 1 To Len(inline)
    If (Mid$(inline, I, 1) = sepchar) Then
      J = J + 1
      If (J > ArgNum) Then Exit For
    Else
      If (J = ArgNum) Then
        RetStr = RetStr + Mid$(inline, I, 1)
      End If
    End If
  Next I
End Sub

Function Parser_GetNumArgs(sepchar As String, inline As String) As Integer
Dim NumArgs As Integer
Dim I As Integer
  NumArgs = 1     'between chr #1 and first separator char.
  For I = 1 To Len(inline)
    If (Mid$(inline, I, 1) = sepchar) Then
      NumArgs = NumArgs + 1
    End If
  Next I
  Parser_GetNumArgs = NumArgs
End Function

Function Parser_RemoveCharacters(remove_char As String, inline As String) As String
Dim RetStr As String
Dim I As Integer
Dim ok_append As Integer
Dim thisc As String
  RetStr = ""
  For I = 1 To Len(inline)
    ok_append = True
    thisc = Mid$(inline, I, 1)
    If (thisc = remove_char) Then ok_append = False
    If (ok_append) Then
      RetStr = RetStr & thisc
    End If
  Next I
  Parser_RemoveCharacters = RetStr
End Function

Function Parser_RemoveDuplicateSeparators(sepchar As String, inline As String) As String
Dim RetStr As String
Dim I As Integer
Dim ok_append As Integer
Dim thisc As String
  RetStr = ""
  For I = 1 To Len(inline)
    ok_append = True
    thisc = Mid$(inline, I, 1)
    If (I > 1) Then
      If (thisc = sepchar) Then
        If (Right$(RetStr, 1) = sepchar) Then
          ok_append = False
        End If
      End If
    End If
    If (ok_append) Then
      RetStr = RetStr & thisc
    End If
  Next I
  Parser_RemoveDuplicateSeparators = RetStr
End Function

