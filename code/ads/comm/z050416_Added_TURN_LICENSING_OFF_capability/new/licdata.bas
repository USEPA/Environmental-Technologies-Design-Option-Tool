Attribute VB_Name = "LicData"
Option Explicit

Global Const AppProgramKey = "ADS"
Global Const TURN_LICENSING_OFF = True
Global Const AppCopyrightYears = "1994-2005"
''''Global Const AppName = "AdXDesignS"
Global AppName_For_Display_Short As String
Global AppName_For_Display_Long As String

Global AppWillExpire As Integer     'true/false
Global AppExpireYear As Integer
Global AppExpireMonth As Integer
Global AppExpireDay As Integer
Global Global_fpath_dir_CPAS As String


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
  Z_USERCODE As String    'xaxa
  'ZZ_LASTEXECUTIONDATE As String
  'ZZ_LASTEXECUTIONTIME As String
End Type
Global lfd As LicFile_Data_Type

Function get_expiration_info( _
    IsOnFrontWindow As Boolean) _
    As String
Dim This_UserCode As Long
Dim Is_Commercial_License As Boolean
If (TURN_LICENSING_OFF = True) Then
  get_expiration_info = "No Expiration Date"
Else
  Select Case Trim$(UCase$(lfd.Z_VERSIONTYPE))
    Case Trim$(UCase$("VER_INTERNAL_STUDENT")):
      If (IsOnFrontWindow = True) Then
        get_expiration_info = ""
      Else
        get_expiration_info = "No Expiration Date (Student Copy)"
      End If
    Case Trim$(UCase$("VER_WONT_EXPIRE")):
      This_UserCode = CLng(Val(lfd.Z_USERCODE))
      Is_Commercial_License = True
      If (This_UserCode >= 500000) Then
        Is_Commercial_License = False
      End If
      If (IsOnFrontWindow = True) Then
        If (Is_Commercial_License = True) Then
          get_expiration_info = "Commercial Version"
        Else
          get_expiration_info = "Academic Version"
        End If
      Else
        If (Is_Commercial_License = True) Then
          get_expiration_info = "Commercial Version (No Expiration Date)"
        Else
          get_expiration_info = "Academic Version (No Expiration Date)"
        End If
      End If
    Case Else:
      get_expiration_info = _
          "Expires on " & _
          Trim$(Str$(AppExpireMonth)) & "/" & _
          Trim$(Str$(AppExpireDay)) & "/" & _
          Trim$(Str$(AppExpireYear))
  End Select
End If
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
Function get_program_releasetype() As String
Dim ver As String
Dim capped As String
  capped = LCase$(Trim$(lfd.Z_RELEASETYPE))
  If (Len(capped) >= 1) Then
    Mid$(capped, 1, 1) = UCase$(Mid$(capped, 1, 1))
  End If
  get_program_releasetype = capped
End Function

Sub LicFileData_Read(return_fpath_dir_CPAS As String)
Dim WinDir As String
Dim fn_CPASCHK As String
Dim CmdLine As String
Dim time_start As Double
Dim fn_GoodLicenseFile As String
Dim fn_BadLicenseFile As String
Dim time_elapsed As Double
Dim f As Integer
Dim RetVal As Integer
Dim copy_z_expirationdate As String
Dim temp As String
Dim fn_CPASDIR_INI As String
Dim fpath_dir_CPAS As String
On Error GoTo err_Cant_Read_Licensing_Data

  'GET CPAS DIRECTORY NAME.
  fn_CPASDIR_INI = App.Path & "\CPASDIR.INI"
  If (Not FileExists(fn_CPASDIR_INI)) Then
    'UNABLE TO READ LICENSE FILE DATA.
    GoTo err_Cant_Read_Licensing_Data
  End If
  temp = Trim$(INI_GetSetting00(fn_CPASDIR_INI, "Directory", "CPASDIR"))
  fpath_dir_CPAS = temp
  return_fpath_dir_CPAS = temp

  'CHECK ON LICENSE FILE.
  WinDir = GetWindowsDir()
  'fn_MTCHK = WinDir & "\" & LICFILE_GetInfoProgram
  fn_CPASCHK = fpath_dir_CPAS & "\DBASE\" & LICFILE_GetInfoProgram
  If (FileExists(fn_CPASCHK)) Then
    'THAT'S OKAY.
  Else
    'UNABLE TO READ LICENSE FILE DATA.
    GoTo err_Cant_Read_Licensing_Data
  End If
  CmdLine = fn_CPASCHK & " " & LICFILE_GetInfoProgramParams
  CmdLine = CmdLine & " " & fpath_dir_CPAS
  CmdLine = CmdLine & " " & AppProgramKey
  'fn_GoodLicenseFile = WinDir & "\" & LICFILE_GoodLicenseFile
  'fn_BadLicenseFile = WinDir & "\" & LICFILE_BadLicenseFile
  fn_GoodLicenseFile = fpath_dir_CPAS & "\DBASE\" & LICFILE_GoodLicenseFile
  fn_BadLicenseFile = fpath_dir_CPAS & "\DBASE\" & LICFILE_BadLicenseFile
  time_start = Timer
  RetVal = 0 * Shell(CmdLine, 1)
  Do While (1 = 1)
    DoEvents
    If (FileExists(fn_GoodLicenseFile)) Then
      'Kill fn_GoodLicenseFile    'DELETED BELOW.
      DoEvents
      Exit Do
    End If
    If (FileExists(fn_BadLicenseFile)) Then
      Kill fn_BadLicenseFile
      DoEvents
      End
    End If
    time_elapsed = Timer - time_start
    If (time_elapsed > 10#) Then
      'UNABLE TO READ LICENSE FILE DATA.
      GoTo err_Cant_Read_Licensing_Data
    End If
  Loop

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
  Line Input #f, lfd.Z_USERCODE       'xaxa
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
  MsgBox "Unable to read licensing data.  You may need to re-install the software.", 48, AppName_For_Display_Short
  End
End Sub

Sub Parser_GetArg(sepchar As String, inline As String, ArgNum As Integer, retStr As String)
Dim i As Integer
Dim J As Integer
  retStr = ""
  J = 1
  For i = 1 To Len(inline)
    If (Mid$(inline, i, 1) = sepchar) Then
      J = J + 1
      If (J > ArgNum) Then Exit For
    Else
      If (J = ArgNum) Then
        retStr = retStr + Mid$(inline, i, 1)
      End If
    End If
  Next i
End Sub

Function Parser_GetNumArgs(sepchar As String, inline As String) As Integer
Dim NumArgs As Integer
Dim i As Integer
  NumArgs = 1     'between chr #1 and first separator char.
  For i = 1 To Len(inline)
    If (Mid$(inline, i, 1) = sepchar) Then
      NumArgs = NumArgs + 1
    End If
  Next i
  Parser_GetNumArgs = NumArgs
End Function

Function Parser_RemoveCharacters(remove_char As String, inline As String) As String
Dim retStr As String
Dim i As Integer
Dim ok_append As Integer
Dim thisc As String
  retStr = ""
  For i = 1 To Len(inline)
    ok_append = True
    thisc = Mid$(inline, i, 1)
    If (thisc = remove_char) Then ok_append = False
    If (ok_append) Then
      retStr = retStr & thisc
    End If
  Next i
  Parser_RemoveCharacters = retStr
End Function

Function Parser_RemoveDuplicateSeparators(sepchar As String, inline As String) As String
Dim retStr As String
Dim i As Integer
Dim ok_append As Integer
Dim thisc As String
  retStr = ""
  For i = 1 To Len(inline)
    ok_append = True
    thisc = Mid$(inline, i, 1)
    If (i > 1) Then
      If (thisc = sepchar) Then
        If (Right$(retStr, 1) = sepchar) Then
          ok_append = False
        End If
      End If
    End If
    If (ok_append) Then
      retStr = retStr & thisc
    End If
  Next i
  Parser_RemoveDuplicateSeparators = retStr
End Function

'NOTE: THERE IS NO RECURSION CHECKER!  IT IS POSSIBLE
'TO SEND THIS ROUTINE INTO AN INFINITE LOOP WITH
'POORLY CHOSEN PARAMETERS.
Function Parser_ReplaceStrings( _
    InputStr As String, _
    OldStr As String, _
    NewStr As String) As String
'Dim Instr_Result As String
Dim Instr_Result As Integer
Dim WorkingStr As String
Dim Part1 As String
Dim Part2 As String
  WorkingStr = InputStr
  
''temp
'Open "c:\test.out" For Output As #1
'Dim i As Integer
'For i = 1 To Len(WorkingStr)
'  Print #1, Asc(Mid$(WorkingStr, i, 1))
'Next i
'Close #1
'  MsgBox WorkingStr
  
  Do While (1 = 1)
    Instr_Result = InStr(WorkingStr, OldStr)
    If (Instr_Result = 0) Then
      Exit Do
    End If
    If (Instr_Result > 1) Then
      Part1 = Left$(WorkingStr, Instr_Result - 1)
    End If
    If (Instr_Result < Len(WorkingStr) - Len(OldStr) + 1) Then
      Part2 = Right$(WorkingStr, Len(WorkingStr) - Instr_Result - Len(OldStr) + 1)
    End If
    WorkingStr = Part1 & NewStr & Part2
'123456789012
'testingXXout           12-2+1=11       12-8-2+1=3
'testingXXo             10-2+1=9        10-8-2+1=1
'testingXX              9-2+1=8         9-8-2+1=0
'-----------------------------------------------------
'12345678901
'testingXout            12-2+1=11       11-8-1+1=3
'testingXo              10-2+1=9        9-8-1+1=1
'testingX               9-2+1=8         8-8-1+1=0
  Loop
  
'Open "c:\test.out" For Output As #1
'For i = 1 To Len(WorkingStr)
'  Print #1, Asc(Mid$(WorkingStr, i, 1))
'Next i
'Close #1
  
  Parser_ReplaceStrings = WorkingStr
End Function



