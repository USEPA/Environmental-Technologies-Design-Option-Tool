Option Explicit

Global Const AppCopyrightYears = "1993-1998"
Global Const AppName = "StEPP"
Global AppWillExpire As Integer     'true/false
Global AppExpireYear As Integer
Global AppExpireMonth As Integer
Global AppExpireDay As Integer




Global Const LICFILE_GetInfoProgram = "MTCHK.EXE"
Global Const LICFILE_GetInfoProgramParams = "-GET_INFO"

Global Const LICFILE_LicName = "ETDOT10.LIC"
'Global Const LICFILE_NewLicInfo = "MTNEWLIC.X"
'Global Const LICFILE_GoodSerialNumber = "OKNUM.X"
'Global Const LICFILE_BadSerialNumber = "BADNUM.X"
Global Const LICFILE_GoodLicenseFile = "GO.X"
Global Const LICFILE_BadLicenseFile = "EXIT.X"

Type LicFile_Data_Type
  Z_EXPIRATIONDATE As String
  'Z_PROGRAMKEY_ADS As String
  'Z_PROGRAMKEY_ASAP As String
  'Z_PROGRAMKEY_STEPP As String
  Z_RELEASETYPE As String
  Z_SERIALNUMBER As String
  Z_USERCOMPANY As String
  Z_USERNAME As String
  Z_VERSIONCODE As String
  Z_VERSIONTYPE As String
  'ZZ_LASTEXECUTIONDATE As String
  'ZZ_LASTEXECUTIONTIME As String
End Type
Global lfd As LicFile_Data_Type

Function get_expiration_info () As String
  Select Case Trim$(UCase$(lfd.Z_VERSIONTYPE))
    Case Trim$(UCase$("INTERNAL_STUDENT")):
      get_expiration_info = "No Expiration Date (Student Copy)"
    Case Trim$(UCase$("EXTERNAL_WONT_EXPIRE")):
      get_expiration_info = "No Expiration Date (Professional Copy)"
    Case Else:
      get_expiration_info = "Expires on " & Trim$(Str$(AppExpireMonth)) & "/" & Trim$(Str$(AppExpireDay)) & "/" & Trim$(Str$(AppExpireYear))
  End Select
End Function

Function get_program_version_with_build_info () As String
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

Sub LicFileData_Read ()
Dim WinDir As String
Dim fn_MTCHK As String
Dim cmdline As String
Dim time_start As Double
Dim fn_GoodLicenseFile As String
Dim fn_BadLicenseFile As String
Dim time_elapsed As Double
Dim f As Integer
Dim retval As Integer
Dim copy_z_expirationdate As String
Dim temp As String

  WinDir = GetWindowsDir()
  fn_MTCHK = WinDir & "\" & LICFILE_GetInfoProgram
  If (fileexists(fn_MTCHK)) Then
    'THAT'S OKAY.
  Else
    'UNABLE TO READ LICENSE FILE DATA.
    GoTo err_Cant_Read_Licensing_Data
  End If
  cmdline = fn_MTCHK & " " & LICFILE_GetInfoProgramParams
  time_start = Timer
  fn_GoodLicenseFile = WinDir & "\" & LICFILE_GoodLicenseFile
  fn_BadLicenseFile = WinDir & "\" & LICFILE_BadLicenseFile
  retval = Shell(cmdline, 1)
  Do While (1 = 1)
    DoEvents
    If (fileexists(fn_GoodLicenseFile)) Then
      'Kill fn_GoodLicenseFile    'DELETED BELOW.
      DoEvents
      Exit Do
    End If
    If (fileexists(fn_BadLicenseFile)) Then
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
  Line Input #f, lfd.Z_EXPIRATIONDATE
  Line Input #f, lfd.Z_RELEASETYPE
  Line Input #f, lfd.Z_SERIALNUMBER
  Line Input #f, lfd.Z_USERCOMPANY
  Line Input #f, lfd.Z_USERNAME
  Line Input #f, lfd.Z_VERSIONCODE
  Line Input #f, lfd.Z_VERSIONTYPE
  Close #f
  Kill fn_GoodLicenseFile
  
  Select Case Trim$(UCase$(lfd.Z_VERSIONTYPE))
    Case Trim$(UCase$("INTERNAL_STUDENT")):
      AppWillExpire = False
    Case Trim$(UCase$("EXTERNAL_WONT_EXPIRE")):
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

Sub Parser_GetArg (sepchar As String, inline As String, ArgNum As Integer, retStr As String)
Dim i As Integer
Dim j As Integer
  retStr = ""
  j = 1
  For i = 1 To Len(inline)
    If (Mid$(inline, i, 1) = sepchar) Then
      j = j + 1
      If (j > ArgNum) Then Exit For
    Else
      If (j = ArgNum) Then
        retStr = retStr + Mid$(inline, i, 1)
      End If
    End If
  Next i
End Sub

Function Parser_GetNumArgs (sepchar As String, inline As String) As Integer
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

Function Parser_RemoveCharacters (remove_char As String, inline As String) As String
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

Function Parser_RemoveDuplicateSeparators (sepchar As String, inline As String) As String
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

