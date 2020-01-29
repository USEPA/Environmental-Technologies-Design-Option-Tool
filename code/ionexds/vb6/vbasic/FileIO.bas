Attribute VB_Name = "FileIO"
Option Explicit

Global Project_Is_Dirty As Boolean





Const FileIO_declarations_end = True


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
Function FileExists(fn As String) As Boolean
  FileExists = File_IsExists(fn)
End Function
Sub KillFile_If_Exists(fn As String)
  If (File_IsExists(fn)) Then
    On Error Resume Next
    Kill fn
  End If
End Sub


Sub file_new()
'  'DISABLE RESULTS MENU: PSDM, CPHSDM, ECM, COMPARE PSDM, COMPARE CPHSDM.
'  frmMain.mnuResultsItem(0).Enabled = False
'  frmMain.mnuResultsItem(1).Enabled = False
'  frmMain.mnuResultsItem(2).Enabled = False
'  frmMain.mnuResultsItem(3).Enabled = False
'  frmMain.mnuResultsItem(4).Enabled = False
'  frmMain.mnuResultsItem(10).Enabled = False      'PSDM-IN-ROOM.
'  'DISABLE OPTIONS MENU: FOULING, INFLUENT CONC, EFFLUENT CONC.
'  frmMain.mnuOptionsItem(0).Enabled = False
'  frmMain.mnuOptionsItem(1).Enabled = False
'  frmMain.mnuOptionsItem(2).Enabled = False
'  'DISABLE RUN MENU: PSDM, CPHSDM, ECM.
'  frmMain.mnuRunItem(0).Enabled = False
'  frmMain.mnuRunItem(1).Enabled = False
'  frmMain.mnuRunItem(2).Enabled = False
'  frmMain.mnuRunItem(10).Enabled = False      'PSDM-IN-ROOM.
'  'DISABLE FILE MENU: SAVE.
'  frmMain.mnuFileItem(2).Enabled = False
  
  'INITIALIZE FOR LIQUID PHASE DEFAULTS.
  Call Project_SetDefaults(NowProj)
  Current_Filename = ""
  frmMain.Caption = Name_App_Short & "  -  (Untitled)"
  Call frmMain_Refresh

  'CLEAR DIRTY (CHANGES) FLAG.
  ''''Project_Is_Dirty = False
  ''''Call DirtyStatus_Set_Current
  Call Global_DirtyStatus_Set(frmMain, Project_Is_Dirty, False)
End Sub


Sub File_Open(fn_Open As String)
Dim f As Integer
Dim ThisVersion As Double
Dim ShowLegacyWarning As Boolean
Dim OpenedOkay As Boolean
Dim IsInvalidFormat As Boolean
Dim DbTest1 As Database
Dim Rs1 As Recordset
Dim DataVersion_Major As Integer
Dim DataVersion_Minor As Integer
  On Error GoTo err_file_open
  If (Not FileExists(fn_Open)) Then
    Call Show_Error("File `" & fn_Open & "` does not exist.")
    GoTo exit_sub
  End If
  frmMain.MousePointer = 11
  'DETERMINE WHETHER THIS IS A VALID MDB DATA FILE.
  IsInvalidFormat = True
  On Error Resume Next
  Set DbTest1 = OpenDatabase(fn_Open)
  If (Err.Number = 0) Then
    IsInvalidFormat = False
    DbTest1.Close
  End If
  On Error GoTo err_file_open
  If (IsInvalidFormat = True) Then
    'FILE IS INVALID; EXIT WITH AN ERROR.
    GoTo exit_invalid_format
  Else
    'DETERMINE WHETHER THE MDB FORMAT FILE IS A LEGACY-VERSION.
    IsInvalidFormat = True
    Set DbTest1 = OpenDatabase(fn_Open)
    If (Database_IsTableExist(DbTest1, "Version") = False) Then
      'INVALID FORMAT.
    Else
      Set Rs1 = DbTest1.OpenRecordset("Version")
      If (Database_NoRecordsInRecordset(Rs1)) Then
        'INVALID FORMAT.
      Else
        DataVersion_Major = 0#
        DataVersion_Minor = 0#
        Rs1.MoveFirst
        Do Until Rs1.EOF
          Select Case Trim$(UCase$(Database_Get_String(Rs1, "FieldName")))
            Case Trim$(UCase$("DataVersion_Major")): Call Database_LoadProperty(Rs1, DataVersion_Major)
            Case Trim$(UCase$("DataVersion_Minor")): Call Database_LoadProperty(Rs1, DataVersion_Minor)
            ''''Case Trim$(UCase$("ContainsPSDMInRoomData")): Call Database_LoadProperty(Rs1, ContainsPSDMInRoomData)
          End Select
          Rs1.MoveNext
        Loop
        Rs1.Close
        DbTest1.Close
      End If
    End If
    ShowLegacyWarning = False
    If (DataVersion_Major = 1) Then
      Select Case DataVersion_Minor
        Case 0:
          'OPEN A NON-LEGACY-VERSION FILE.
          Call file_new
          OpenedOkay = File_Open_Latest_v1_00(fn_Open)
          If (OpenedOkay) Then IsInvalidFormat = False
          ShowLegacyWarning = False
        Case Else:
          'OPEN A LEGACY-VERSION FILE.
          Call file_new
          OpenedOkay = File_Open_Latest_v1_00(fn_Open)
          If (OpenedOkay) Then IsInvalidFormat = False
          ShowLegacyWarning = True
      End Select
    End If
    If (IsInvalidFormat) Then
      'FILE IS INVALID; EXIT WITH AN ERROR.
      GoTo exit_invalid_format
    End If
  End If
  'SHOW LEGACY WARNING IF NECESSARY.
  If (ShowLegacyWarning) Then
    Call Show_Message00("Warning: This file is formatted as a " & _
        "Version " & Trim$(Str$(DataVersion_Major)) & "." & _
        Trim$(Str$(DataVersion_Minor)) & _
        " file.  If saved, it will be saved as a " & _
        "Version " & Trim$(Str$(Latest_DataVersion_Major)) & "." & _
        Trim$(Str$(Latest_DataVersion_Minor)) & _
        " file.", _
        vbInformation, _
        App.Title & " : Legacy File Version Warning")
  End If
  Close #f
  'UPDATE DISPLAY.
  Current_Filename = fn_Open
  frmMain.Caption = Name_App_Short & "  -  " & Trim$(Current_Filename)
  Call frmMain_Refresh
  'CLEAR DIRTY FLAG.
  ''''Project_Is_Dirty = False
  ''''Call DirtyStatus_Set_Current
  Call Global_DirtyStatus_Set(frmMain, Project_Is_Dirty, False)
  'MOVE THIS FILENAME TO TOP OF LAST-FEW-FILES LIST.
  Call OldFileList_Promote( _
      Current_Filename, _
      1, _
      frmMain.mnuFileItem(199), _
      frmMain.mnuFileItem(191), _
      frmMain.mnuFileItem(192), _
      frmMain.mnuFileItem(193), _
      frmMain.mnuFileItem(194))
  GoTo exit_sub
exit_sub:
  frmMain.MousePointer = 0
  Exit Sub
exit_invalid_format:
  Call Show_Error("The selected file is not a " & _
      "valid file.")
  GoTo exit_sub
exit_err_file_open:
  Call file_new
  GoTo exit_sub
err_file_open:
  Call Show_Trapped_Error("file_open")
  On Error Resume Next
  Close f
  Resume exit_err_file_open
End Sub
Sub File_OpenAs(fn_force As String)
Dim fn_openas As String
  If (fn_force <> "") Then
    fn_openas = fn_force
  Else
    'INPUT NEW FILENAME.
    On Error GoTo err_file_openas
    frmMain.CommonDialog1.DialogTitle = "Open " & Name_App_Short & " File"
    'frmMain.CommonDialog1.Filter = "All Files (*.*)|*.*|" & _
    '    Name_App_Short & " Files (*.dat)|*.dat"
    frmMain.CommonDialog1.Filter = "All Files (*.*)|*.*|" & _
        Name_App_Short & " Files (*." & FileExt_App & ")|*." & _
        FileExt_App
    frmMain.CommonDialog1.FilterIndex = 2
    frmMain.CommonDialog1.CancelError = True
    frmMain.CommonDialog1.flags = _
        cdlOFNFileMustExist + _
        cdlOFNPathMustExist
    frmMain.CommonDialog1.ShowOpen
    fn_openas = Trim$(frmMain.CommonDialog1.filename)
    If (fn_openas = "") Then
      'DO NOTHING.
      Exit Sub
    End If
  End If
  'OPEN THIS FILE.
  Call File_Open(fn_openas)
exit_err_file_openas:
  Exit Sub
err_file_openas:
  If (Err.Number = cdlCancel) Then
    'CANCEL BUTTON WAS SELECTED.
    Resume exit_err_file_openas
  End If
  Resume exit_err_file_openas
End Sub


'RETURNS:
'- true = save went okay.
'- false = save failed.
Function File_Save(fn_Save As String) As Boolean
Dim f As Integer
Dim SavedOkay As Boolean
  On Error GoTo err_File_Save
  'SAVE FILE.
  frmMain.MousePointer = 11
  SavedOkay = File_Save_Latest_v1_00(fn_Save)
  If (SavedOkay = False) Then
    GoTo exit_err_File_Save
  End If
  'CLEAR DIRTY FLAG.
  ''''Project_Is_Dirty = False
  ''''Call DirtyStatus_Set_Current
  Call Global_DirtyStatus_Set(frmMain, Project_Is_Dirty, False)
  File_Save = True
  'UPDATE DISPLAY.
  Current_Filename = fn_Save
  frmMain.Caption = Name_App_Short & "  -  " & Trim$(Current_Filename)
  Call frmMain_Refresh
  'MOVE THIS FILENAME TO TOP OF LAST-FEW-FILES LIST.
  Call OldFileList_Promote( _
      Current_Filename, _
      1, _
      frmMain.mnuFileItem(199), _
      frmMain.mnuFileItem(191), _
      frmMain.mnuFileItem(192), _
      frmMain.mnuFileItem(193), _
      frmMain.mnuFileItem(194))
  GoTo Exit_Function
Exit_Function:
  frmMain.MousePointer = 0
  Exit Function
exit_err_File_Save:
  File_Save = False
  GoTo Exit_Function
err_File_Save:
  Call Show_Trapped_Error("File_Save")
  On Error Resume Next
  Close f
  Resume exit_err_File_Save
End Function
'RETURNS:
'- true = save went okay.
'- false = save failed.
Function File_SaveAs(fn_force As String) As Boolean
Dim f As Integer
Dim fn_saveas As String
Dim RetVal As Integer
  If (fn_force <> "") Then
    fn_saveas = fn_force
  Else
    Do While (1 = 1)
      'INPUT NEW FILENAME.
      On Error GoTo err_File_SaveAs
      frmMain.CommonDialog1.DialogTitle = "Save " & _
          Name_App_Short & " File"
      frmMain.CommonDialog1.Filter = _
          "All Files (*.*)|*.*|" & Name_App_Short & _
          " Files (*." & FileExt_App & ")|*." & FileExt_App
      frmMain.CommonDialog1.FilterIndex = 2
      frmMain.CommonDialog1.CancelError = True
      frmMain.CommonDialog1.flags = _
          cdlOFNOverwritePrompt + _
          cdlOFNPathMustExist
      frmMain.CommonDialog1.ShowSave
      fn_saveas = Trim$(frmMain.CommonDialog1.filename)
      If (fn_saveas = "") Then
        'DO NOTHING.
        Exit Function
      End If
      'If (Not File_IsExists(fn_saveas)) Then
      '  Exit Do
      'End If
      'RetVal = MsgBox("File " & fn_saveas & _
      '    " already exists.  Do you want to replace it?", _
      '    vbQuestion + vbYesNo, _
      '    App.Title & " : Overwrite File ?")
      'If (RetVal = vbYes) Then Exit Do
      '
      ' NOTE: "REPLACE?" CHECK HANDLED IN COMMON
      ' DIALOG CONTROL NOW.
      '
      Exit Do
    Loop
  End If
  'OPEN THIS FILE.
  File_SaveAs = File_Save(fn_saveas)
  Exit Function
exit_err_File_SaveAs:
  File_SaveAs = False
  Exit Function
err_File_SaveAs:
  If (Err.Number = cdlCancel) Then
    'CANCEL BUTTON WAS SELECTED.
    Resume exit_err_File_SaveAs
  End If
  Call Show_Trapped_Error("File_SaveAs")
  Resume exit_err_File_SaveAs
End Function


Function Project_IsDirtyFlagThrown() As Boolean
  If (Project_Is_Dirty) Then
    Project_IsDirtyFlagThrown = True
  Else
    Project_IsDirtyFlagThrown = False
  End If
End Function


'RETURNS:
'- true = it's okay to unload this file now.
'- false = cancel the unload.
Function file_query_unload() As Integer
Dim RetVal As Integer
Dim msg As String
  If (Not Project_IsDirtyFlagThrown()) Then
    file_query_unload = True
    Exit Function
  End If
  msg = "Do you want to save the changes you made to "
  If (Current_Filename = "") Then
    msg = msg & "this new project"
  Else
    msg = msg & "your project of filename " & Current_Filename
  End If
  msg = msg & " ?"
  RetVal = MsgBox(msg, vbCritical + vbYesNoCancel, App.Title & " : Save Changes ?")
  Select Case RetVal
    Case vbYes:
      If (File_SaveAs(Current_Filename) = True) Then
        'SAVE WENT OK; IT'S NOW OKAY TO UNLOAD THIS FILE.
        file_query_unload = True
      Else
        'SAVE FAILED; DON'T UNLOAD THIS FILE.
        file_query_unload = False
      End If
      Exit Function
    Case vbNo:
      file_query_unload = True
      Exit Function
    Case vbCancel:
      file_query_unload = False
      Exit Function
  End Select
End Function


Sub ProjectFile_Read(f As Integer, ByRef RetVal As Variant, Optional optDummy1 As Variant)
Dim outputstr$
Dim outlin As String
Dim sub_name As String
Dim input1 As String
Dim input2 As String
  Input #f, input1, input2
  sub_name = "ProjectFile_Read"
  Select Case VarType(RetVal)
    Case vbBoolean
      RetVal = Val(input1)
    Case vbByte, vbInteger, vbLong, vbCurrency
      RetVal = Val(input1)
    Case vbSingle, vbDouble
      RetVal = Val(input1)
    Case vbString, vbDate
      RetVal = input1
    Case vbObject
        MsgBox sub_name & " vbObject not implemented"
        GoTo EXIT_FALSE_VALUE
    Case vbError
        MsgBox sub_name & " vbError not implemented"
        GoTo EXIT_FALSE_VALUE
    Case vbDataObject
        MsgBox sub_name & " vbDataObject not implemented"
        GoTo EXIT_FALSE_VALUE
    Case vbVariant
        MsgBox sub_name & " vbVariant not implemented"
        GoTo EXIT_FALSE_VALUE
    Case vbArray
        MsgBox sub_name & " vbArray not implemented"
        GoTo EXIT_FALSE_VALUE
    Case vbEmpty
        MsgBox sub_name & " vbEmpty not implemented"
        GoTo EXIT_FALSE_VALUE
    Case vbNull
        MsgBox sub_name & " vbNull not implemented"
        GoTo EXIT_FALSE_VALUE
  End Select
  GoTo EXIT_OK
EXIT_FALSE_VALUE:
  Print #f, "   - - - ERROR IN " & sub_name & "() - - -"
  Exit Sub
EXIT_OK:
  Exit Sub
End Sub
Sub ProjectFile_Write(f As Integer, v As Variant, s As String)
Dim outputstr$
Dim outlin As String
Dim sub_name As String
  sub_name = "ProjectFile_Write"
  Select Case VarType(v)
    Case vbBoolean
        outputstr$ = IIf(v, "1", "0")
    Case vbByte, vbInteger, vbLong, vbCurrency
        outputstr$ = Trim$(CStr(v))
    Case vbSingle, vbDouble
        outputstr$ = Trim$(CStr(v))
    Case vbString, vbDate
        outputstr$ = CStr(v)
    Case vbObject
        MsgBox sub_name & " vbObject not implemented"
        GoTo EXIT_FALSE_VALUE
    Case vbError
        MsgBox sub_name & " vbError not implemented"
        GoTo EXIT_FALSE_VALUE
    Case vbDataObject
        MsgBox sub_name & " vbDataObject not implemented"
        GoTo EXIT_FALSE_VALUE
    Case vbVariant
        MsgBox sub_name & " vbVariant not implemented"
        GoTo EXIT_FALSE_VALUE
    Case vbArray
        MsgBox sub_name & " vbArray not implemented"
        GoTo EXIT_FALSE_VALUE
    Case vbEmpty
        MsgBox sub_name & " vbEmpty not implemented"
        GoTo EXIT_FALSE_VALUE
    Case vbNull
        MsgBox sub_name & " vbNull not implemented"
        GoTo EXIT_FALSE_VALUE
  End Select
  outlin = Chr$(34) & Trim$(outputstr$) & Chr$(34) & "," & _
      Chr$(34) & s & Chr$(34)
  'outlin = Trim$(outputstr$)
  'If (Len(outlin) > 27) Then
  '  outlin = outlin & "    "
  'Else
  '  Do While (1 = 1)
  '    If (Len(outlin) >= 27) Then Exit Do
  '    outlin = outlin & " "
  '  Loop
  'End If
  'outlin = outlin & s
  Print #f, outlin
  GoTo EXIT_OK
EXIT_FALSE_VALUE:
  Print #f, "   - - - ERROR IN " & sub_name & "() - - -"
  Exit Sub
EXIT_OK:
  Exit Sub
End Sub



