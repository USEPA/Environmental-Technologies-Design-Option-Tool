Attribute VB_Name = "CPASCHK_11"
Option Explicit





Const CPASCHK_11_declarations_end = True


Sub Do_Create_File_v11(arg_CpasDir As String)
Dim fn_NEWLIC As String
Dim f As Integer
Dim result As Integer
Dim fn_test As String

Dim lfdt As LicFile_Data_Type
Dim pkdt_new() As ProgramKey_Data_Type
Dim num_pkdt_new As Integer
Dim num_pkdt_existing As Integer
Dim fn_CPASLIC As String
Dim filesize_CPASLIC As Long
Dim offset_This_PK As Long
Dim offset_This_Data As Long
Dim i As Integer
Dim j As Integer
Dim dummy As String
Dim slot_to_use() As Integer
Dim num_new_slots As Integer
Dim This_Position As Integer

Dim fn_MTCHKLIC As String

Dim fpath_OutputDir As String
Dim RetVal As Integer

Dim ForceAll_Z_EXPIRATIONDATE As String
Dim ForceAll_Z_RELEASETYPE As String
Dim ForceAll_Z_VERSIONCODE As String
Dim ForceAll_Z_VERSIONTYPE As String
Dim iExpiresDay As Integer
Dim iExpiresMonth As Integer
Dim iExpiresYear As Integer
Dim ThisSerialNumber As String

  On Error GoTo err_Do_Create_File_v11
  
  'DETERMINE OUTPUT DIRECTORY.
  fpath_OutputDir = arg_CpasDir & "\DBASE"


  ''LET WORLD KNOW THE CALL WORKED (temporary).
  'f = FreeFile
  'fn_test = WinPathWindows$ & "WORKED.X"
  'Open fn_test For Output As #f
  'Print #f, "0"
  'Close #f

  'DELETE {CPAS}\DBASE\OKNUM.X AND {CPAS}\DBASE\BADNUM.X IF PRESENT.
  fn_test = arg_CpasDir & "\DBASE\" & LICFILE_GoodSerialNumber
  If (FileExists(fn_test)) Then
    Kill fn_test
  End If
  fn_test = arg_CpasDir & "\DBASE\" & LICFILE_BadSerialNumber
  If (FileExists(fn_test)) Then
    Kill fn_test
  End If

  'READ IN THE {CPAS}\DBASE\NEWLIC.X FILE.
  fn_NEWLIC = arg_CpasDir & "\DBASE\" & LICFILE_NewLicInfo
  If (FileExists(fn_NEWLIC)) Then
    'THAT'S OKAY.
  Else
    'EXIT WITH AN ERROR: CREATE {CPAS}\DBASE\BADNUM.X.
    'Call ExitError(arg_CpasDir, LICFILE_BadSerialNumber)
    Call ExitError(fpath_OutputDir, LICFILE_BadSerialNumber)
  End If
  f = FreeFile
  Open fn_NEWLIC For Input As #f
  Line Input #f, lfdt.Z_SERIALNUMBER
  Line Input #f, lfdt.Z_USERNAME
  Line Input #f, lfdt.Z_USERCOMPANY
'  num_pkdt_new = 0
'  Do While (1 = 1)
'    Line Input #f, dummy
'    If (Trim$(UCase$(dummy)) = Trim$(UCase$("END"))) Then
'      Exit Do
'    End If
'    num_pkdt_new = num_pkdt_new + 1
'    ReDim Preserve pkdt_new(1 To num_pkdt_new)
'    'pkdt_new(num_pkdt_new).Z_PROGRAMKEY = dummy
'    'Line Input #f, pkdt_new(num_pkdt_new).Z_EXPIRATIONDATE
'    'Line Input #f, pkdt_new(num_pkdt_new).Z_RELEASETYPE
'    'Line Input #f, pkdt_new(num_pkdt_new).Z_VERSIONCODE
'    'Line Input #f, pkdt_new(num_pkdt_new).Z_VERSIONTYPE
'  Loop
  Close #f

  'ERASE THE {CPAS}\DBASE\NEWLIC.X FILE.
  Kill fn_NEWLIC
  
  'VERIFY THE SERIAL NUMBER.
  ''''result = Snums_Verify(Trim$(lfdt.Z_SERIALNUMBER))
  result = snumVerify(Trim$(lfdt.Z_SERIALNUMBER))
  If (result = 0) Then
    'EXIT WITH AN ERROR: CREATE {CPAS}\DBASE\BADNUM.X.
    Call ExitError(fpath_OutputDir, LICFILE_BadSerialNumber)
    'Call ExitError(arg_CpasDir, LICFILE_BadSerialNumber)
  Else
    'THAT'S OKAY.
  End If

  '
  'THIS IS WHERE THE FIELDS Z_PROGRAMKEY, Z_EXPIRATIONDATE,
  'Z_RELEASETYPE, AND Z_VERSIONTYPE ARE SET BASED ON THE Z_SERIALNUMBER FIELD.
  'NOTE THAT Z_VERSIONCODE IS FORCED TO "1.0" BECAUSE IT IS
  'NO LONGER USED.
  '
  'NOTE THAT NO UPDATES TO CPAS.LIC ARE PERFORMED ANYMORE.
  'NOW, THE FILE IS REBUILT FROM SCRATCH EACH TIME THE -CREATE_FILE
  'OPTION IS RUN.
  '
  ThisSerialNumber = Trim$(lfdt.Z_SERIALNUMBER)
  ForceAll_Z_VERSIONCODE = "1.0"  'FORCE ALL TO "1.0"; THIS FIELD IS NO LONGER USED.
  RetVal = snumGetVersionType(ThisSerialNumber)
  Select Case RetVal
    Case 1:     'ALPHA VERSION.
      ForceAll_Z_RELEASETYPE = "ALPHA"
    Case 2:     'BETA VERSION.
      ForceAll_Z_RELEASETYPE = "BETA"
    Case 3:     'STANDARD VERSION.
      ForceAll_Z_RELEASETYPE = "STANDARD"
    Case Else:
      'EXIT WITH AN ERROR: CREATE {CPAS}\DBASE\BADNUM.X.
      Call ExitError(fpath_OutputDir, LICFILE_BadSerialNumber)
  End Select
  RetVal = snumIsExpirationPresent(ThisSerialNumber)
  If (RetVal = 0) Then
    'NO EXPIRATION DATE PRESENT.
    ForceAll_Z_EXPIRATIONDATE = "NEVER"
    ForceAll_Z_VERSIONTYPE = "VER_WONT_EXPIRE"
  Else
    'EXPIRATION DATE IS PRESENT.
    iExpiresDay = snumGetExpirationDay(ThisSerialNumber)
    If (iExpiresDay = 0) Then
      'EXIT WITH AN ERROR: CREATE {CPAS}\DBASE\BADNUM.X.
      Call ExitError(fpath_OutputDir, LICFILE_BadSerialNumber)
    End If
    iExpiresMonth = snumGetExpirationMonth(ThisSerialNumber)
    If (iExpiresMonth = 0) Then
      'EXIT WITH AN ERROR: CREATE {CPAS}\DBASE\BADNUM.X.
      Call ExitError(fpath_OutputDir, LICFILE_BadSerialNumber)
    End If
    iExpiresYear = snumGetExpirationYear(ThisSerialNumber)
    ForceAll_Z_EXPIRATIONDATE = _
        Trim$(Str$(iExpiresMonth)) & "," & _
        Trim$(Str$(iExpiresDay)) & "," & _
        Trim$(Str$(iExpiresYear))
    ForceAll_Z_VERSIONTYPE = "VER_WILL_EXPIRE"
  End If
  '
  'THIS IS WHERE THE LIST OF RECOGNIZED MODULES IS GENERATED.
  '
  ReDim Recognized_Idx_iModule(0 To 2) As Integer
  ReDim Recognized_Z_PROGRAMKEY(0 To 2) As String
  Recognized_Idx_iModule(0) = 0
  Recognized_Z_PROGRAMKEY(0) = "ADS"
  Recognized_Idx_iModule(1) = 1
  Recognized_Z_PROGRAMKEY(1) = "ASAP"
  Recognized_Idx_iModule(2) = 2
  Recognized_Z_PROGRAMKEY(2) = "STEPP"
  '
  'THIS IS WHERE THE LIST OF PURCHASED MODULES IS GENERATED.
  '
  num_pkdt_new = 0
  For i = LBound(Recognized_Idx_iModule) To UBound(Recognized_Idx_iModule)
    If (snumIsModulePurchased(ThisSerialNumber, i) = 1) Then
      num_pkdt_new = num_pkdt_new + 1
      ReDim Preserve pkdt_new(1 To num_pkdt_new)
      pkdt_new(num_pkdt_new).Z_PROGRAMKEY = Recognized_Z_PROGRAMKEY(i)
      pkdt_new(num_pkdt_new).Z_EXPIRATIONDATE = ForceAll_Z_EXPIRATIONDATE
      pkdt_new(num_pkdt_new).Z_RELEASETYPE = ForceAll_Z_RELEASETYPE
      pkdt_new(num_pkdt_new).Z_VERSIONCODE = ForceAll_Z_VERSIONCODE
      pkdt_new(num_pkdt_new).Z_VERSIONTYPE = ForceAll_Z_VERSIONTYPE
    End If
  Next i
  '
  'GENERATE ENCRYPTION KEY.
  '
  Call LicFile_GenerateEncryptionKey    'must be called once before any other LicFile_*() routine.
  '
  'DOES THE LICENSE FILE CURRENTLY EXIST?
  '
  fn_CPASLIC = arg_CpasDir & "\DBASE\" & LICFILE_LicName
'  If (FileExists(fn_CPASLIC)) Then
  If (True = False) Then
'    'FILE CURRENTLY EXISTS: MODIFY EXISTING FILE.
'    Call LicFile_GetEncryptedString(fn_CPASLIC, CLng(lfdt_order_ZZ_NUMPROGRAMKEYS * 100), dummy)
'    num_pkdt_existing = CInt(Val(dummy))
'    lfdt.ZZ_NUMPROGRAMKEYS = num_pkdt_existing
'    ReDim slot_to_use(1 To num_pkdt_new)
'    For i = 1 To num_pkdt_new
'      slot_to_use(i) = 0
'    Next i
'    'PLAN TO UPDATE EXISTING PROGRAM KEY ENTRIES (IF ANY).
'    For i = 1 To lfdt.ZZ_NUMPROGRAMKEYS
'      offset_This_PK = 1000 + (i - 1) * 1000
'      offset_This_Data = CLng(offset_This_PK + pkdt_order_Z_PROGRAMKEY * 100)
'      Call LicFile_GetEncryptedString(fn_CPASLIC, offset_This_Data, dummy)
'      For j = 1 To num_pkdt_new
'        If (Trim$(UCase$(dummy)) = Trim$(UCase$(pkdt_new(j).Z_PROGRAMKEY))) Then
'          slot_to_use(j) = i
'          Exit For
'        End If
'      Next j
'    Next i
'    num_new_slots = 0
'    For i = 1 To num_pkdt_new
'      If (slot_to_use(i) = 0) Then
'        num_new_slots = num_new_slots + 1
'        lfdt.ZZ_NUMPROGRAMKEYS = lfdt.ZZ_NUMPROGRAMKEYS + 1
'        slot_to_use(i) = lfdt.ZZ_NUMPROGRAMKEYS
'      End If
'    Next i
'    'APPEND GARBAGE CHARACTERS TO END OF FILE TO ALLOW SPACE FOR NEW PROGRAM KEY ENTRIES.
'    If (num_new_slots <> 0) Then
'      Call LicFile_AppendGarbage(fn_CPASLIC, CLng(num_new_slots * 1000))
'    End If
'    'RE-OUTPUT NUMBER OF PROGRAM KEYS.
'    dummy = Trim$(Str$(lfdt.ZZ_NUMPROGRAMKEYS))
'    Call LicFile_PutEncryptedString(fn_CPASLIC, CLng(lfdt_order_ZZ_NUMPROGRAMKEYS * 100), dummy)
'    'OUTPUT PROGRAM KEY ENTRIES.
'    For i = 1 To num_pkdt_new
'      This_Position = slot_to_use(i)
'      offset_This_PK = 1000 + (This_Position - 1) * 1000
'      offset_This_Data = CLng(offset_This_PK + pkdt_order_Z_PROGRAMKEY * 100)
'      Call LicFile_PutEncryptedString(fn_CPASLIC, offset_This_Data, pkdt_new(i).Z_PROGRAMKEY)
'      offset_This_Data = CLng(offset_This_PK + pkdt_order_Z_EXPIRATIONDATE * 100)
'      Call LicFile_PutEncryptedString(fn_CPASLIC, offset_This_Data, pkdt_new(i).Z_EXPIRATIONDATE)
'      offset_This_Data = CLng(offset_This_PK + pkdt_order_Z_RELEASETYPE * 100)
'      Call LicFile_PutEncryptedString(fn_CPASLIC, offset_This_Data, pkdt_new(i).Z_RELEASETYPE)
'      offset_This_Data = CLng(offset_This_PK + pkdt_order_Z_VERSIONCODE * 100)
'      Call LicFile_PutEncryptedString(fn_CPASLIC, offset_This_Data, pkdt_new(i).Z_VERSIONCODE)
'      offset_This_Data = CLng(offset_This_PK + pkdt_order_Z_VERSIONTYPE * 100)
'      Call LicFile_PutEncryptedString(fn_CPASLIC, offset_This_Data, pkdt_new(i).Z_VERSIONTYPE)
'    Next i
'    'MODIFY EXISTING FILE: OPERATION COMPLETE.
  Else
    'DELETE THE FILE IF IT EXISTS.
    If (FileExists(fn_CPASLIC)) Then
      Kill fn_CPASLIC
    End If
    '
    'FILE DOES NOT CURRENTLY EXIST: START FROM SCRATCH.
    '
    filesize_CPASLIC = 1000 + (num_pkdt_new) * 1000 + 374
    Call LicFile_Create(fn_CPASLIC, filesize_CPASLIC)
    'IMPORTANT STEP!!  SET DATE/TIME TO "NEVER" STRINGS.
    lfdt.ZZ_LASTEXECUTIONDATE = LICFILE_DATE_NEVER
    lfdt.ZZ_LASTEXECUTIONTIME = LICFILE_DATE_NEVER
    'UPDATE PROGRAM KEY COUNT.
    lfdt.ZZ_NUMPROGRAMKEYS = num_pkdt_new
    dummy = Trim$(Str$(lfdt.ZZ_NUMPROGRAMKEYS))
    Call LicFile_PutEncryptedString(fn_CPASLIC, CLng(lfdt_order_ZZ_NUMPROGRAMKEYS * 100), dummy)
    For i = 1 To num_pkdt_new
      This_Position = i
      offset_This_PK = 1000 + (This_Position - 1) * 1000
      offset_This_Data = CLng(offset_This_PK + pkdt_order_Z_PROGRAMKEY * 100)
      Call LicFile_PutEncryptedString(fn_CPASLIC, offset_This_Data, pkdt_new(i).Z_PROGRAMKEY)
      offset_This_Data = CLng(offset_This_PK + pkdt_order_Z_EXPIRATIONDATE * 100)
      Call LicFile_PutEncryptedString(fn_CPASLIC, offset_This_Data, pkdt_new(i).Z_EXPIRATIONDATE)
      offset_This_Data = CLng(offset_This_PK + pkdt_order_Z_RELEASETYPE * 100)
      Call LicFile_PutEncryptedString(fn_CPASLIC, offset_This_Data, pkdt_new(i).Z_RELEASETYPE)
      offset_This_Data = CLng(offset_This_PK + pkdt_order_Z_VERSIONCODE * 100)
      Call LicFile_PutEncryptedString(fn_CPASLIC, offset_This_Data, pkdt_new(i).Z_VERSIONCODE)
      offset_This_Data = CLng(offset_This_PK + pkdt_order_Z_VERSIONTYPE * 100)
      Call LicFile_PutEncryptedString(fn_CPASLIC, offset_This_Data, pkdt_new(i).Z_VERSIONTYPE)
    Next i
    'CREATE NEW FILE: OPERATION COMPLETE.
  End If
  
  'IN EITHER FILE-SCENARIO, OUTPUT THE HEADER INFO.
  Call LicFile_PutEncryptedString(fn_CPASLIC, CLng(lfdt_order_Z_SERIALNUMBER * 100), lfdt.Z_SERIALNUMBER)
  Call LicFile_PutEncryptedString(fn_CPASLIC, CLng(lfdt_order_Z_USERNAME * 100), lfdt.Z_USERNAME)
  Call LicFile_PutEncryptedString(fn_CPASLIC, CLng(lfdt_order_Z_USERCOMPANY * 100), lfdt.Z_USERCOMPANY)
  Call LicFile_PutEncryptedString(fn_CPASLIC, CLng(lfdt_order_ZZ_LASTEXECUTIONDATE * 100), lfdt.ZZ_LASTEXECUTIONDATE)
  Call LicFile_PutEncryptedString(fn_CPASLIC, CLng(lfdt_order_ZZ_LASTEXECUTIONTIME * 100), lfdt.ZZ_LASTEXECUTIONTIME)

  ''CREATE THE NEW LICENSE FILE.
  'Call LicFile_GenerateEncryptionKey    'must be called once before any other LicFile_*() routine.
  'Call LicFile_Create(fn_ControlFile, LICFILE_MAXSIZE)
  'Call LicFile_PutEncryptedString(fn_ControlFile, LICFILE_Z_EXPIRATIONDATE, lfdt.Z_EXPIRATIONDATE)
  'Call LicFile_PutEncryptedString(fn_ControlFile, LICFILE_Z_PROGRAMKEY_ADS, lfdt.Z_PROGRAMKEY_ADS)
  'Call LicFile_PutEncryptedString(fn_ControlFile, LICFILE_Z_PROGRAMKEY_ASAP, lfdt.Z_PROGRAMKEY_ASAP)
  'Call LicFile_PutEncryptedString(fn_ControlFile, LICFILE_Z_PROGRAMKEY_STEPP, lfdt.Z_PROGRAMKEY_STEPP)
  'Call LicFile_PutEncryptedString(fn_ControlFile, LICFILE_Z_RELEASETYPE, lfdt.Z_RELEASETYPE)
  'Call LicFile_PutEncryptedString(fn_ControlFile, LICFILE_Z_SERIALNUMBER, lfdt.Z_SERIALNUMBER)
  'Call LicFile_PutEncryptedString(fn_ControlFile, LICFILE_Z_USERCOMPANY, lfdt.Z_USERCOMPANY)
  'Call LicFile_PutEncryptedString(fn_ControlFile, LICFILE_Z_USERNAME, lfdt.Z_USERNAME)
  'Call LicFile_PutEncryptedString(fn_ControlFile, LICFILE_Z_VERSIONCODE, lfdt.Z_VERSIONCODE)
  'Call LicFile_PutEncryptedString(fn_ControlFile, LICFILE_Z_VERSIONTYPE, lfdt.Z_VERSIONTYPE)
  'Call LicFile_PutEncryptedString(fn_ControlFile, LICFILE_ZZ_LASTEXECUTIONDATE, LICFILE_DATE_NEVER)
  'Call LicFile_PutEncryptedString(fn_ControlFile, LICFILE_ZZ_LASTEXECUTIONTIME, LICFILE_DATE_NEVER)

  'ATTEMPT TO CREATE THE FILE {WIN}\MTCHK.LIC;
  'IF UNABLE TO CREATE THE FILE, NO BIG DEAL.
  On Error Resume Next
  'f = FreeFile
  fn_MTCHKLIC = WinPathWindows$ & LICFILE_ExtraCheckFile
  Call LicFile_Create(fn_MTCHKLIC, 1277)
  Call LicFile_PutEncryptedString(fn_MTCHKLIC, 671, LICFILE_ExtraCheckFile_Text)
  'Kill fn_MTCHKLIC
  On Error GoTo err_Do_Create_File_v11

  'EXIT WITH SUCCESS: CREATE {CPAS}\DBASE\OKNUM.X.
  Call ExitCode(fpath_OutputDir, LICFILE_GoodSerialNumber)
  'Call ExitCode(arg_CpasDir, LICFILE_GoodSerialNumber)
  End   'redundant

exit_err_Do_Create_File_v11:
  'ERRORS OCCURED; EXIT WITH AN ERROR CODE.
  Call ExitError(fpath_OutputDir, LICFILE_BadSerialNumber)
  'Call ExitError(arg_CpasDir, LICFILE_BadSerialNumber)
  End   'redundant
err_Do_Create_File_v11:
  'err.description
  Resume exit_err_Do_Create_File_v11
End Sub


Sub Do_Display_All_v11(arg_CpasDir As String)
Dim fn_CPASLIC As String
Dim dummy As String
Dim i As Integer
Dim lfdt As LicFile_Data_Type
Dim pkdt_this As ProgramKey_Data_Type
Dim msg As String
Dim num_pkdt As Integer
Dim This_Position As Integer
Dim offset_This_PK As Long
Dim offset_This_Data As Long

Dim fpath_OutputDir As String

  'DETERMINE OUTPUT DIRECTORY.
  fpath_OutputDir = arg_CpasDir & "\DBASE"

  'DOES THE LICENSE FILE CURRENTLY EXIST?
  fn_CPASLIC = arg_CpasDir & "\DBASE\" & LICFILE_LicName
  If (FileExists(fn_CPASLIC)) Then
    'GENERATE ENCRYPTION KEY.
    Call LicFile_GenerateEncryptionKey    'must be called once before any other LicFile_*() routine.
    'FILE CURRENTLY EXISTS: DISPLAY CONTENTS.
    Call LicFile_GetEncryptedString(fn_CPASLIC, CLng(lfdt_order_ZZ_NUMPROGRAMKEYS * 100), dummy)
    num_pkdt = CInt(Val(dummy))
    msg = ""
    Call LicFile_GetEncryptedString(fn_CPASLIC, CLng(lfdt_order_ZZ_LASTEXECUTIONDATE * 100), lfdt.ZZ_LASTEXECUTIONDATE)
    Call LicFile_GetEncryptedString(fn_CPASLIC, CLng(lfdt_order_ZZ_LASTEXECUTIONTIME * 100), lfdt.ZZ_LASTEXECUTIONTIME)
    msg = msg & "lfdt.ZZ_LASTEXECUTIONDATE = `" & lfdt.ZZ_LASTEXECUTIONDATE & "`"
    msg = msg & Chr$(13) & Chr$(10)
    msg = msg & "lfdt.ZZ_LASTEXECUTIONTIME = `" & lfdt.ZZ_LASTEXECUTIONTIME & "`"
    msg = msg & Chr$(13) & Chr$(10)
    msg = msg & Chr$(13) & Chr$(10)
    For i = 1 To num_pkdt
      This_Position = i
      offset_This_PK = 1000 + (This_Position - 1) * 1000
      offset_This_Data = CLng(offset_This_PK + pkdt_order_Z_PROGRAMKEY * 100)
      Call LicFile_GetEncryptedString(fn_CPASLIC, offset_This_Data, pkdt_this.Z_PROGRAMKEY)
      offset_This_Data = CLng(offset_This_PK + pkdt_order_Z_EXPIRATIONDATE * 100)
      Call LicFile_GetEncryptedString(fn_CPASLIC, offset_This_Data, pkdt_this.Z_EXPIRATIONDATE)
      offset_This_Data = CLng(offset_This_PK + pkdt_order_Z_RELEASETYPE * 100)
      Call LicFile_GetEncryptedString(fn_CPASLIC, offset_This_Data, pkdt_this.Z_RELEASETYPE)
      offset_This_Data = CLng(offset_This_PK + pkdt_order_Z_VERSIONCODE * 100)
      Call LicFile_GetEncryptedString(fn_CPASLIC, offset_This_Data, pkdt_this.Z_VERSIONCODE)
      offset_This_Data = CLng(offset_This_PK + pkdt_order_Z_VERSIONTYPE * 100)
      Call LicFile_GetEncryptedString(fn_CPASLIC, offset_This_Data, pkdt_this.Z_VERSIONTYPE)
      msg = msg & Trim$(Str$(i)) & ": " & Trim$(pkdt_this.Z_PROGRAMKEY) & ","
      msg = msg & "" & Trim$(pkdt_this.Z_EXPIRATIONDATE) & ","
      msg = msg & "" & Trim$(pkdt_this.Z_RELEASETYPE) & ","
      msg = msg & "" & Trim$(pkdt_this.Z_VERSIONCODE) & ","
      msg = msg & "" & Trim$(pkdt_this.Z_VERSIONTYPE) & "."
      msg = msg & Chr$(13) & Chr$(10)
    Next i
    MsgBox msg
  Else
    MsgBox "File does not exist: `" & fn_CPASLIC & "`."
  End If

End Sub


Sub Do_Get_Info_v11(arg_CpasDir As String, arg_ProgramKey As String, arg_ResultsDir As String)
Dim fn_test As String
Dim lfdt As LicFile_Data_Type
Dim copy_z_versiontype As String
Dim copy_z_expirationdate As String
Dim has_expired As Integer
Dim f As Integer
Dim fn_GO As String
Dim msg As String
Dim result As Integer

Dim s_date_now As String
ReDim date_year(1 To 2) As Integer
ReDim date_month(1 To 2) As Integer
ReDim date_day(1 To 2) As Integer
ReDim time_hour(1 To 2) As Integer
ReDim time_minute(1 To 2) As Integer
ReDim time_second(1 To 2) As Integer
ReDim s_date(1 To 2) As String
ReDim s_time(1 To 2) As String
Dim DateExpires_year As Integer
Dim DateExpires_month As Integer
Dim DateExpires_day As Integer
Dim temp As String
Dim has_set_clock_back As Integer
Dim has_set_clock_back_2 As Integer

Dim fn_CPASLIC As String
Dim num_pkdt_existing As Integer
Dim i As Integer
Dim offset_This_PK As Long
Dim offset_This_Data As Long
Dim This_Position As Integer
Dim pkdt As ProgramKey_Data_Type
Dim dummy As String
Dim found_it As Integer
Dim fn_MTCHKLIC As String

Dim fpath_OutputDir As String

  On Error GoTo err_Do_Get_Info_v11
  '
  ' DETERMINE OUTPUT DIRECTORY.
  '
  If (Trim$(arg_ResultsDir) = "") Then
    fpath_OutputDir = arg_CpasDir & "\DBASE"
  Else
    fpath_OutputDir = Trim$(arg_ResultsDir)
  End If
  '
  ' DELETE {CPAS}\DBASE\GO.X AND {CPAS}\DBASE\EXIT.X IF PRESENT.
  '
  fn_test = fpath_OutputDir & "\" & LICFILE_GoodLicenseFile
  'fn_test = arg_CpasDir & "\DBASE\" & LICFILE_GoodLicenseFile
  If (FileExists(fn_test)) Then
    Kill fn_test
  End If
  fn_test = fpath_OutputDir & "\" & LICFILE_BadLicenseFile
  'fn_test = arg_CpasDir & "\DBASE\" & LICFILE_BadLicenseFile
  If (FileExists(fn_test)) Then
    Kill fn_test
  End If
  '
  ' READ IN THE {CPAS}\DBASE\CPAS.LIC FILE.
  '
  fn_CPASLIC = arg_CpasDir & "\DBASE\" & LICFILE_LicName
  If (FileExists(fn_CPASLIC)) Then
    'THAT'S OKAY.
  Else
    'EXIT WITH AN ERROR: CREATE {CPAS}\DBASE\EXIT.X.
    Call ExitError(fpath_OutputDir, LICFILE_BadLicenseFile)
    'Call ExitError(arg_CpasDir, LICFILE_BadLicenseFile)
  End If
  Call LicFile_GenerateEncryptionKey    'must be called once before any other LicFile_*() routine.
  Call LicFile_GetEncryptedString(fn_CPASLIC, CLng(lfdt_order_Z_SERIALNUMBER * 100), lfdt.Z_SERIALNUMBER)
  Call LicFile_GetEncryptedString(fn_CPASLIC, CLng(lfdt_order_Z_USERNAME * 100), lfdt.Z_USERNAME)
  Call LicFile_GetEncryptedString(fn_CPASLIC, CLng(lfdt_order_Z_USERCOMPANY * 100), lfdt.Z_USERCOMPANY)
  Call LicFile_GetEncryptedString(fn_CPASLIC, CLng(lfdt_order_ZZ_LASTEXECUTIONDATE * 100), lfdt.ZZ_LASTEXECUTIONDATE)
  Call LicFile_GetEncryptedString(fn_CPASLIC, CLng(lfdt_order_ZZ_LASTEXECUTIONTIME * 100), lfdt.ZZ_LASTEXECUTIONTIME)
  Call LicFile_GetEncryptedString(fn_CPASLIC, CLng(lfdt_order_ZZ_NUMPROGRAMKEYS * 100), dummy)
  lfdt.ZZ_NUMPROGRAMKEYS = CInt(Val(dummy))
  '
  ' LOOK FOR THAT PROGRAM KEY.
  '
  num_pkdt_existing = lfdt.ZZ_NUMPROGRAMKEYS
  found_it = False
  For i = 1 To num_pkdt_existing
    This_Position = i
    offset_This_PK = 1000 + (This_Position - 1) * 1000
    offset_This_Data = CLng(offset_This_PK + pkdt_order_Z_PROGRAMKEY * 100)
    Call LicFile_GetEncryptedString(fn_CPASLIC, offset_This_Data, dummy)
    If (Trim$(UCase$(dummy)) = Trim$(UCase$(arg_ProgramKey))) Then
      pkdt.Z_PROGRAMKEY = dummy
      offset_This_Data = CLng(offset_This_PK + pkdt_order_Z_EXPIRATIONDATE * 100)
      Call LicFile_GetEncryptedString(fn_CPASLIC, offset_This_Data, pkdt.Z_EXPIRATIONDATE)
      offset_This_Data = CLng(offset_This_PK + pkdt_order_Z_RELEASETYPE * 100)
      Call LicFile_GetEncryptedString(fn_CPASLIC, offset_This_Data, pkdt.Z_RELEASETYPE)
      offset_This_Data = CLng(offset_This_PK + pkdt_order_Z_VERSIONCODE * 100)
      Call LicFile_GetEncryptedString(fn_CPASLIC, offset_This_Data, pkdt.Z_VERSIONCODE)
      offset_This_Data = CLng(offset_This_PK + pkdt_order_Z_VERSIONTYPE * 100)
      Call LicFile_GetEncryptedString(fn_CPASLIC, offset_This_Data, pkdt.Z_VERSIONTYPE)
      found_it = True
      Exit For
    End If
  Next i
  If (Not found_it) Then
    'UNABLE TO FIND THAT PROGRAM KEY!
    'EXIT WITH AN ERROR: CREATE {CPAS}\DBASE\EXIT.X.
    Call ExitError(fpath_OutputDir, LICFILE_BadLicenseFile)
    'Call ExitError(arg_CpasDir, LICFILE_BadLicenseFile)
  End If
  '
  'CHECK THE {WIN}\MTCHK.LIC FILE; IF IT DOES NOT EXIST, AND THE USER
  'HAS WRITE ACCESS IN THE {WIN} DIRECTORY, EXIT WITH AN ERROR.
  '
  fn_MTCHKLIC = WinPathWindows$ & LICFILE_ExtraCheckFile
  If (FileExists(fn_MTCHKLIC)) Then
    Call LicFile_GetEncryptedString(fn_MTCHKLIC, 671, dummy)
    If (Trim$(UCase$(dummy)) <> Trim$(UCase$(LICFILE_ExtraCheckFile_Text))) Then
      'EXIT WITH AN ERROR: CREATE {CPAS}\DBASE\EXIT.X.
      Call ExitError(fpath_OutputDir, LICFILE_BadLicenseFile)
      'Call ExitError(arg_CpasDir, LICFILE_BadLicenseFile)
    End If
    'FILE VERIFIED, OKAY TO CONTINUE.
  Else
    'TEST TO SEE IF USER HAS WRITE ACCESS TO THAT FILE.
    On Error Resume Next
    f = FreeFile
    Call LicFile_Create(fn_MTCHKLIC, 1277)
    If (Err <> 0) Then
      'OKAY, USER DOES NOT HAVE ACCESS.
    Else
      'USER HAS WRITE ACCESS!
      'EXIT WITH AN ERROR: CREATE {CPAS}\DBASE\EXIT.X.
      Call ExitError(fpath_OutputDir, LICFILE_BadLicenseFile)
      'Call ExitError(arg_CpasDir, LICFILE_BadLicenseFile)
    End If
    'Call LicFile_PutEncryptedString(fn_MTCHKLIC, 671, LICFILE_ExtraCheckFile_Text)
    On Error GoTo err_Do_Get_Info_v11
  End If
  '
  'VERIFY THE SERIAL NUMBER.
  '
  ''''result = Snums_Verify(Trim$(lfdt.Z_SERIALNUMBER))
  result = snumVerify(Trim$(lfdt.Z_SERIALNUMBER))
  If (result = 0) Then
    'EXIT WITH AN ERROR: CREATE {CPAS}\DBASE\EXIT.X.
    msg = "The serial number for the software is invalid.  "
    msg = msg & "Please contact "
    msg = msg & LICFILE_ExpireContactInfo & " to correct "
    msg = msg & "this problem."
    Call ShowError(msg)
    GoTo exit_err_Do_Get_Info_v11
  Else
    'THAT'S OKAY.
  End If
  '
  ' GET THE CURRENT DATE/TIME.
  '
  s_date_now = Now
  date_year(1) = Year(s_date_now)
  date_month(1) = Month(s_date_now)
  date_day(1) = Day(s_date_now)
  time_hour(1) = Hour(s_date_now)
  time_minute(1) = Minute(s_date_now)
  time_second(1) = Second(s_date_now)
  
        'date_year(1) = 1998
        'date_month(1) = 11
        'date_day(1) = 1

  s_date(1) = Format$(date_year(1), "0000") & Format$(date_month(1), "00") & Format$(date_day(1), "00")
  s_time(1) = Format$(time_hour(1), "00") & Format$(time_minute(1), "00") & Format$(time_second(1), "00")
  '
  ' PERFORM DATE CHECKING FOR EXPIRATION DATES (IF ANY).
  '
  copy_z_versiontype = Trim$(UCase$(pkdt.Z_VERSIONTYPE))
  Select Case copy_z_versiontype
    Case Trim$(UCase$("VER_WONT_EXPIRE")):
      'DO NOTHING; THIS TYPE DOES NOT EXPIRE.
    Case Trim$(UCase$("VER_INTERNAL_STUDENT")):
      'CHECK THE STUDENT EXPIRATION FILE.
      fn_test = LICFILE_StudentCheckFile
      If (FileExists(fn_test)) Then
        'THAT'S OKAY.
      Else
        'EXIT WITH AN ERROR.
        msg = "This internal student version of the software "
        msg = msg & "cannot be verified.  Please contact "
        msg = msg & LICFILE_ExpireContactInfo & " to correct "
        msg = msg & "this problem."
        Call ShowError(msg)
        GoTo exit_err_Do_Get_Info_v11
      End If
    Case Else:      'ASSUME EXTERNAL_WILL_EXPIRE.
      'PERFORM DATE CHECKING: HAS THE EXPIRATION DATE PASSED?
      
      copy_z_expirationdate = Trim$(UCase$(pkdt.Z_EXPIRATIONDATE))
      copy_z_expirationdate = Parser_RemoveCharacters(" ", copy_z_expirationdate)
'MsgBox "copy_z_expirationdate = `" & copy_z_expirationdate & "`"
      has_expired = True
      If (Parser_GetNumArgs(",", copy_z_expirationdate) = 3) Then
        Call Parser_GetArg(",", copy_z_expirationdate, 1, temp)
        DateExpires_month = CInt(Val(temp))
        Call Parser_GetArg(",", copy_z_expirationdate, 2, temp)
        DateExpires_day = CInt(Val(temp))
        Call Parser_GetArg(",", copy_z_expirationdate, 3, temp)
        DateExpires_year = CInt(Val(temp))
        'HAS THE DATE PASSED YET?  (IS DATE #1 -- the current date -- GREATER THAN THE EXPIRATION DATE?)
        If (date_year(1) = DateExpires_year) Then
          If (date_month(1) = DateExpires_month) Then
            If (date_day(1) <= DateExpires_day) Then
              has_expired = False
            End If
          End If
          If (date_month(1) < DateExpires_month) Then
            has_expired = False
          End If
        End If
        If (date_year(1) < DateExpires_year) Then
          has_expired = False
        End If
      End If
  'MsgBox "got here (a)"
      If (has_expired) Then
        'SOFTWARE HAS EXPIRED; WRITE THE EXECUTION DATE/TIME TO THE FILE.
        'THIS IS IMPORTANT: IF THE USER SETS THEIR CLOCK BACK NOW, THE SOFTWARE WILL
        'STILL REFUSE TO RUN.
        lfdt.ZZ_LASTEXECUTIONDATE = s_date(1)
        lfdt.ZZ_LASTEXECUTIONTIME = s_time(1)
        'Call LicFile_PutEncryptedString(fn_ControlFile, LICFILE_ZZ_LASTEXECUTIONDATE,lfdt.ZZ_LASTEXECUTIONDATE)
        'Call LicFile_PutEncryptedString(fn_ControlFile, LICFILE_ZZ_LASTEXECUTIONTIME,lfdt.ZZ_LASTEXECUTIONTIME)
        Call LicFile_PutEncryptedString(fn_CPASLIC, CLng(lfdt_order_ZZ_LASTEXECUTIONDATE * 100), lfdt.ZZ_LASTEXECUTIONDATE)
        Call LicFile_PutEncryptedString(fn_CPASLIC, CLng(lfdt_order_ZZ_LASTEXECUTIONTIME * 100), lfdt.ZZ_LASTEXECUTIONTIME)
        'TELL THE USER.
        msg = "This software expired on " & Trim$(Str$(DateExpires_month)) & "/"
        msg = msg & Trim$(Str$(DateExpires_day)) & "/" & Trim$(Str$(DateExpires_year))
        msg = msg & ".  Please contact "
        msg = msg & LICFILE_ExpireContactInfo & " for a "
        msg = msg & "new copy of the software."
        Call ShowError(msg)
        GoTo exit_err_Do_Get_Info_v11
      End If
  'MsgBox "got here (b)"
      'PERFORM DATE CHECKING: DID THE USER SET THEIR CLOCK BACK?
      has_set_clock_back = True
      s_date(2) = Trim$(lfdt.ZZ_LASTEXECUTIONDATE)
      s_time(2) = Trim$(lfdt.ZZ_LASTEXECUTIONTIME)
'MsgBox "s_date(2) = `" & s_date(2) & "`,s_time(2) = `" & s_time(2) & "`"
      If (s_date(2) = LICFILE_DATE_NEVER) Or (s_time(2) = LICFILE_DATE_NEVER) Or (s_date(2) = "") Or (s_time(2) = "") Then
        'SKIP THE CHECK; THE PROGRAM HAS NEVER BEEN EXECUTED BEFORE.
      Else
        'PERFORM THE CHECK.
        If (Len(s_date(2)) <> 8) Then GoTo exit_err_Do_Get_Info_v11
        If (Len(s_time(2)) <> 6) Then GoTo exit_err_Do_Get_Info_v11
        date_year(2) = CInt(Val(Mid$(s_date(2), 1, 4)))
        date_month(2) = CInt(Val(Mid$(s_date(2), 5, 2)))
        date_day(2) = CInt(Val(Mid$(s_date(2), 7, 2)))
        time_hour(2) = CInt(Val(Mid$(s_time(2), 1, 2)))
        time_minute(2) = CInt(Val(Mid$(s_time(2), 3, 2)))
        time_second(2) = CInt(Val(Mid$(s_time(2), 5, 2)))
        'IS THE LAST EXECUTION DATE (#2) GREATER THAN THE CURRENT DATE (#1)?
        If (date_year(2) = date_year(1)) Then
          If (date_month(2) = date_month(1)) Then
            If (date_day(2) <= date_day(1)) Then
              has_set_clock_back = False
            End If
          End If
          If (date_month(2) < date_month(1)) Then
            has_set_clock_back = False
          End If
        End If
        If (date_year(2) < date_year(1)) Then
          has_set_clock_back = False
        End If
        has_set_clock_back_2 = True
  'MsgBox "got here (c)"
        'IS THE LAST EXECUTION DATE (#2) GREATER THAN THE EXPIRATION DATE?
        If (date_year(2) = DateExpires_year) Then
          If (date_month(2) = DateExpires_month) Then
            If (date_day(2) <= DateExpires_day) Then
              has_set_clock_back_2 = False
            End If
          End If
          If (date_month(2) < DateExpires_month) Then
            has_set_clock_back_2 = False
          End If
        End If
        If (date_year(2) < DateExpires_year) Then
          has_set_clock_back_2 = False
        End If
        If (has_set_clock_back_2) And (has_set_clock_back) Then
          'NOT ONLY IS THE CLOCK SET BACK, BUT THE LAST EXECUTION DATE WAS
          'AFTER THE EXPIRATION DATE.  REFUSE TO RUN.
          msg = "This software expired on " & Trim$(Str$(DateExpires_month)) & "/"
          msg = msg & Trim$(Str$(DateExpires_day)) & "/" & Trim$(Str$(DateExpires_year))
          msg = msg & "!  Please contact "
          msg = msg & LICFILE_ExpireContactInfo & " for a "
          msg = msg & "new copy of the software."
          Call ShowError(msg)
          GoTo exit_err_Do_Get_Info_v11
        End If
      End If
  'MsgBox "got here (z)"
  End Select
'MsgBox "got here (1)"
'MsgBox "fn_CPASLIC = `" & fn_CPASLIC & "`"
  '
  'STORE THE LAST EXECUTION DATE/TIME.
  '
  lfdt.ZZ_LASTEXECUTIONDATE = s_date(1)
  lfdt.ZZ_LASTEXECUTIONTIME = s_time(1)
  'Call LicFile_PutEncryptedString(fn_ControlFile, LICFILE_ZZ_LASTEXECUTIONDATE, lfdt.ZZ_LASTEXECUTIONDATE)
  'Call LicFile_PutEncryptedString(fn_ControlFile, LICFILE_ZZ_LASTEXECUTIONTIME, lfdt.ZZ_LASTEXECUTIONTIME)
  Call LicFile_PutEncryptedString(fn_CPASLIC, CLng(lfdt_order_ZZ_LASTEXECUTIONDATE * 100), lfdt.ZZ_LASTEXECUTIONDATE)
  Call LicFile_PutEncryptedString(fn_CPASLIC, CLng(lfdt_order_ZZ_LASTEXECUTIONTIME * 100), lfdt.ZZ_LASTEXECUTIONTIME)
  '
  'OUTPUT THE {CPAS}\DBASE\GO.X FILE.
  '
  f = FreeFile
  'MsgBox "got here (5)"
  fn_GO = fpath_OutputDir & "\" & LICFILE_GoodLicenseFile
  'fn_GO = arg_CpasDir & "\DBASE\" & LICFILE_GoodLicenseFile
  Open fn_GO For Output As #f
  Print #f, lfdt.Z_SERIALNUMBER
  Print #f, lfdt.Z_USERNAME
  Print #f, lfdt.Z_USERCOMPANY
  Print #f, pkdt.Z_PROGRAMKEY
  Print #f, pkdt.Z_EXPIRATIONDATE
  Print #f, pkdt.Z_RELEASETYPE
  Print #f, pkdt.Z_VERSIONCODE
  Print #f, pkdt.Z_VERSIONTYPE
  Close #f
  'MsgBox "got here (9)"
  End

exit_err_Do_Get_Info_v11:
  '
  'ERRORS OCCURED; EXIT WITH AN ERROR CODE ({CPAS}\DBASE\EXIT.X).
  '
  Call ExitError(fpath_OutputDir, LICFILE_BadLicenseFile)
  'Call ExitError(arg_CpasDir, LICFILE_BadLicenseFile)
  End   'redundant
err_Do_Get_Info_v11:
  Resume exit_err_Do_Get_Info_v11

End Sub




