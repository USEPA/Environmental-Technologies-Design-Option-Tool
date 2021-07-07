Attribute VB_Name = "FileIO_LatestMDB"
Option Explicit





Const FileIO_LatestMDB_declarations_end = True


'RETURNS:
'         TRUE = SUCCEEDED IN LOADING.
'         FALSE = FAILED IN LOADING.
Function File_Open_Latest_v1_00( _
    fn_this As String) As Boolean
Dim Ws1 As Workspace
Dim Db1 As Database
Dim rs1 As Recordset
Dim Use_FieldIndex As Integer
Dim Use_FieldIndex2 As Integer
Dim ContainsTable_PSDMInRoomData As Boolean
Dim prj As Project_Type
Dim i As Integer
Dim f As Integer
Dim s As String
Dim fn_outputtxt As String

'>>>>>>>>>>>>>>>>>>>> *TODO* UPDATE THIS ENTIRE FUNCTION <<<<<<<<<<<<<<<<<<<<<<<<<

  '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
  '>>>>>>>>>>>>>>>>>>>>>>>>>>>  INPUT FROM MAIN DATABASE  <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
  '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
  If (Not FileExists(fn_this)) Then
    'ERROR: UNABLE TO FIND THE FILE!
    File_Open_Latest_v1_00 = False
    Exit Function
  End If
  'OPEN DATABASE.
  Set Db1 = OpenDatabase(fn_this)

  '=========== INPUT DATA FROM DATABASE TABLES. =================
  
  '------ INPUT DATA FROM TABLE "Version". ------------------------------------------------------------------------------------------------------
  'APPLICABLE DEFAULT VALUES:
  ContainsTable_PSDMInRoomData = False
  If (Database_IsTableExist(Db1, "Version") = False) Then
    'DO NOTHING: USE DEFAULT VALUES.
  Else
    Set rs1 = Db1.OpenRecordset("Version")
    If (Database_NoRecordsInRecordset(rs1)) Then
      'DO NOTHING: USE DEFAULT VALUES.
    Else
      rs1.MoveFirst
      Do Until rs1.EOF
        ''''Use_FieldIndex = CInt(Database_Get_Long(Rs1, "FieldIndex"))
        Select Case Trim$(UCase$(Database_Get_String(rs1, "FieldName")))
          'HEADER BLOCK.
          Case Trim$(UCase$("ContainsTable_PSDMInRoomData")): Call Database_LoadProperty(rs1, ContainsTable_PSDMInRoomData)
        End Select
        rs1.MoveNext
      Loop
    End If
    rs1.Close
  End If
  
  '------ INPUT DATA FROM TABLE "Main". ------------------------------------------------------------------------------------------------------
  If (Database_IsTableExist(Db1, "Main") = False) Then
    'DO NOTHING: USE DEFAULT VALUES.
  Else
    Set rs1 = Db1.OpenRecordset("Main")
    If (Database_NoRecordsInRecordset(rs1)) Then
      'DO NOTHING: USE DEFAULT VALUES.
    Else
      '
      ' SET DEFAULT PROJECT DATA IN TEMPORARY VARIABLE.
      '
      Call Project_SetDefaults(prj)
      '
      ' READ IN THE PROJECT DATA TO TEMPORARY VARIABLE.
      '
      rs1.MoveFirst
      Do Until rs1.EOF
        Use_FieldIndex = CInt(Database_Get_Long(rs1, "FieldIndex"))
        Select Case Trim$(UCase$(Database_Get_String(rs1, "FieldName")))
           
        Case Trim$(UCase$("DyeStudy_Count")): Call Database_LoadProperty(rs1, prj.dyestudy_count)
        If (prj.dyestudy_count > 0) Then
          ReDim prj.DyeStudy(1 To prj.dyestudy_count)
          rs1.MoveNext
          For i = 1 To prj.dyestudy_count
              Do While i = Trim$(UCase$(Database_Get_Integer(rs1, "FieldIndex")))
                Select Case Trim$(UCase$(Database_Get_String(rs1, "FieldName")))
                  Case Trim$(UCase$("time")): Call Database_LoadProperty(rs1, prj.DyeStudy(i).time)
                  Case Trim$(UCase$("concentration")): Call Database_LoadProperty(rs1, prj.DyeStudy(i).concentration)
                End Select
                rs1.MoveNext
                Loop
            Next i
          rs1.MovePrevious
        End If
        
        Case Trim$(UCase$("DyeStudy_Output")): _
          Call Database_LoadProperty(rs1, prj.dyestudy_output, True)
        f = FreeFile
        fn_outputtxt = App.Path + "\exes\outpt.txt"
        Open fn_outputtxt For Output As #f
        Print #f, prj.dyestudy_output
        Close #f
        
        Case Trim$(UCase$("DyeStudyDisp_Output")): _
          Call Database_LoadProperty(rs1, prj.dyestudydisp_output, True)
        f = FreeFile
        fn_outputtxt = App.Path + "\exes\pecoutpt.txt"
        Open fn_outputtxt For Output As #f
        Print #f, prj.dyestudydisp_output
        Close #f
        
        Case Trim$(UCase$("DyeStudy_Calcdate")): _
          Call Database_LoadProperty(rs1, prj.dyestudy_calcdate)
        
        Case Trim$(UCase$("Predicted_Available")): _
          Call Database_LoadProperty(rs1, prj.Predicted_Available)
        If prj.Predicted_Available Then
          Predicted_Available = True
        Else
          Predicted_Available = False
        End If
        
        Case Trim$(UCase$("Predicted_Count")): Call Database_LoadProperty(rs1, prj.Predicted_count)
        If (prj.Predicted_count > 0) Then
          ReDim prj.Predicted(1 To prj.Predicted_count)
          rs1.MoveNext
          For i = 1 To prj.Predicted_count
              Do While i = Trim$(UCase$(Database_Get_Integer(rs1, "FieldIndex")))
                
                Select Case Trim$(UCase$(Database_Get_String(rs1, "FieldName")))
                  Case Trim$(UCase$("Predicted_Theta")): Call Database_LoadProperty(rs1, prj.Predicted(i).Predicted_Theta)
                  Case Trim$(UCase$("Predicted_E")): Call Database_LoadProperty(rs1, prj.Predicted(i).Predicted_E)
                End Select
                rs1.MoveNext
                If rs1.EOF Then
                  Exit Do
                End If
                Loop
            Next i
          rs1.MovePrevious
        End If
        
        Case Trim$(UCase$("Experimental_Count")): Call Database_LoadProperty(rs1, prj.Experimental_count)
        If (prj.Experimental_count > 0) Then
          ReDim prj.Experimental(1 To prj.Experimental_count)
          rs1.MoveNext
          For i = 1 To prj.Experimental_count
              Do While i = Trim$(UCase$(Database_Get_Integer(rs1, "FieldIndex")))

                Select Case Trim$(UCase$(Database_Get_String(rs1, "FieldName")))
                  Case Trim$(UCase$("Experimental_Theta")): Call Database_LoadProperty(rs1, prj.Experimental(i).Experimental_Theta)
                  Case Trim$(UCase$("Experimental_E")): Call Database_LoadProperty(rs1, prj.Experimental(i).Experimental_E)
                End Select
                rs1.MoveNext
                If rs1.EOF Then
                  Exit Do
                End If
                Loop
            Next i
          rs1.MovePrevious
        End If
        
        Case Trim$(UCase$("PredictedDispClosed_Count")): _
          Call Database_LoadProperty(rs1, prj.PredictedDispClosed_count)
        If (prj.PredictedDispClosed_count > 0) Then
          ReDim prj.DispClosed(1 To prj.PredictedDispClosed_count)
          rs1.MoveNext
          For i = 1 To prj.PredictedDispClosed_count
              Do While i = Trim$(UCase$(Database_Get_Integer(rs1, "FieldIndex")))

                Select Case Trim$(UCase$(Database_Get_String(rs1, "FieldName")))
                  Case Trim$(UCase$("PredictedDispClosed_Theta")): _
                    Call Database_LoadProperty(rs1, prj.DispClosed(i).PredictedDispClosed_Theta)
                  Case Trim$(UCase$("PredictedDispClosed_E")): _
                    Call Database_LoadProperty(rs1, prj.DispClosed(i).PredictedDispClosed_E)
                End Select
                rs1.MoveNext
                If rs1.EOF Then
                  Exit Do
                End If
                Loop
            Next i
          rs1.MovePrevious
        End If
        
        Case Trim$(UCase$("PredictedDispOpen_Count")): _
          Call Database_LoadProperty(rs1, prj.PredictedDispOpen_count)
        If (prj.PredictedDispOpen_count > 0) Then
          ReDim prj.DispOpen(1 To prj.PredictedDispOpen_count)
          rs1.MoveNext
          For i = 1 To prj.PredictedDispOpen_count
              Do While i = Trim$(UCase$(Database_Get_Integer(rs1, "FieldIndex")))

                Select Case Trim$(UCase$(Database_Get_String(rs1, "FieldName")))
                  Case Trim$(UCase$("PredictedDispOpen_Theta")): _
                    Call Database_LoadProperty(rs1, prj.DispOpen(i).PredictedDispOpen_Theta)
                  Case Trim$(UCase$("PredictedDispOpen_E")): _
                    Call Database_LoadProperty(rs1, prj.DispOpen(i).PredictedDispOpen_E)
                End Select
                rs1.MoveNext
                If rs1.EOF Then
                  Exit Do
                End If
                Loop
            Next i
          rs1.MovePrevious
        End If
      
       Case Trim$(UCase$("plottype")): Call Database_LoadProperty(rs1, prj.plottype)
      
       End Select
       rs1.MoveNext
       
      Loop
      '
      ' TRANSFER PROJECT DATA TO MEMORY.
      '
      IsCalculated = True
      nowproj = prj
    End If
    rs1.Close
  End If

  'CLOSE THE DATABASE FILE.
  Db1.Close

  'RETURN A "SUCCESS" MESSAGE TO CALLER.
  File_Open_Latest_v1_00 = True

End Function


'RETURNS:
'         TRUE = SUCCEEDED IN SAVING.
'         FALSE = FAILED IN SAVING.
Function File_Save_Latest_v1_00( _
    fn_this As String) As Boolean
Dim Ws1 As Workspace
Dim Db1 As Database
Dim rs1 As Recordset
Dim i As Integer
Dim j As Integer
Dim IsInvalidFormat As Boolean
Dim NeedToCreateNewDatabase As Boolean
Dim prj As Project_Type

'>>>>>>>>>>>>>>>>>>>> *TODO* UPDATE THIS ENTIRE FUNCTION <<<<<<<<<<<<<<<<<<<<<<<<<

  '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
  '>>>>>>>>>>>>>>>>>>>>>>>>>>>  SAVE TO MAIN DATABASE  <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
  '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
  
  'IF FILE DOES NOT EXIST, CREATE IT.
  'FOR EACH TABLE, IF IT EXISTS, DELETE IT.
  If (Not FileExists(fn_this)) Then
    'CREATE NEW DATABASE.
    NeedToCreateNewDatabase = True
  Else
    'DETERMINE WHETHER OLD FILE IS AN INVALID VERSION (i.e. A NON-MDB FILE).
    IsInvalidFormat = True
    On Error Resume Next
    Set Db1 = OpenDatabase(fn_this)
    If (Err.Number = 0) Then
      IsInvalidFormat = False
      Db1.Close
    End If
    On Error GoTo 0
    If (IsInvalidFormat) Then
      'DELETE OLD FILE, CREATE NEW DATABASE (SEE BELOW).
      Kill fn_this
      NeedToCreateNewDatabase = True
    Else
      'OPEN DATABASE NORMALLY.
      Set Db1 = OpenDatabase(fn_this)
    End If
  End If
  If (NeedToCreateNewDatabase) Then
 '   Set Db1 = CreateDatabase(fn_this, dbLangGeneral)
    FileCopy MAIN_APP_PATH & "\dbase\template.dye", fn_this
  End If
  
  'CREATE NEW TABLES WITHIN DATABASE, IF NECESSARY.
  Call Database_CreateMFBTable_IfNoExist(Db1, "Version", True)
  Call Database_CreateMFBTable_IfNoExist(Db1, "Main", True)
  
  '=========== OUTPUT DATA TO DATABASE TABLES. =================
  
  '------ OUTPUT DATA TO TABLE "Version". ------------------------------------------------------------------------------------------------------
  'START SAVE TO THIS TABLE.
  Call Database_DeleteTableContents(Db1, "Version")
  Set rs1 = Db1.OpenRecordset("Version")
  Call Database_SaveProperty(rs1, "DataVersion_Major", CInt(Latest_DataVersion_Major))
  Call Database_SaveProperty(rs1, "DataVersion_Minor", CInt(Latest_DataVersion_Minor))
  ''''Call Database_SaveProperty(Rs1, "ContainsTable_PSDMInRoomData", True)
  'END SAVE TO THIS TABLE.
  rs1.Close
  
  '------ OUTPUT DATA TO TABLE "Main". ------------------------------------------------------------------------------------------------------
  'START SAVE TO THIS TABLE.
  Call Database_DeleteTableContents(Db1, "Main")
  Set rs1 = Db1.OpenRecordset("Main")
  'MAIN BLOCK.
  prj = nowproj
  
  'DYE STUDY PARAMETERS.
  Call Database_SaveProperty(rs1, "DyeStudy_Count", prj.dyestudy_count)
  For i = 1 To prj.dyestudy_count
    Call Database_SavePropertyWithIndex(rs1, "time", i, prj.DyeStudy(i).time)
    Call Database_SavePropertyWithIndex(rs1, "concentration", i, prj.DyeStudy(i).concentration)
  Next i
  
  Call Database_SaveProperty(rs1, "*DyeStudy_Output", prj.dyestudy_output)
  Call Database_SaveProperty(rs1, "*DyeStudyDisp_Output", prj.dyestudydisp_output)
  
  Call Database_SaveProperty(rs1, "DyeStudy_Calcdate", prj.dyestudy_calcdate)
  
  'PLOT
  Call Database_SaveProperty(rs1, "Predicted_Available", prj.Predicted_Available)
  Call Database_SaveProperty(rs1, "Predicted_Count", prj.Predicted_count)
  For i = 1 To prj.Predicted_count
    Call Database_SavePropertyWithIndex(rs1, "Predicted_Theta", i, prj.Predicted(i).Predicted_Theta)
    Call Database_SavePropertyWithIndex(rs1, "Predicted_E", i, prj.Predicted(i).Predicted_E)
  Next i
  
  Call Database_SaveProperty(rs1, "Experimental_Count", prj.Experimental_count)
  For i = 1 To prj.Experimental_count
    Call Database_SavePropertyWithIndex(rs1, "Experimental_Theta", i, prj.Experimental(i).Experimental_Theta)
    Call Database_SavePropertyWithIndex(rs1, "Experimental_E", i, prj.Experimental(i).Experimental_E)
  Next i
  
  Call Database_SaveProperty(rs1, "PredictedDispClosed_Count", prj.PredictedDispClosed_count)
  For i = 1 To prj.PredictedDispClosed_count
    Call Database_SavePropertyWithIndex(rs1, "PredictedDispClosed_Theta", i, _
      prj.DispClosed(i).PredictedDispClosed_Theta)
    Call Database_SavePropertyWithIndex(rs1, "PredictedDispClosed_E", i, _
      prj.DispClosed(i).PredictedDispClosed_E)
  Next i
  
  Call Database_SaveProperty(rs1, "PredictedDispOpen_Count", prj.PredictedDispOpen_count)
  For i = 1 To prj.PredictedDispOpen_count
    Call Database_SavePropertyWithIndex(rs1, "PredictedDispOpen_Theta", i, _
      prj.DispOpen(i).PredictedDispOpen_Theta)
    Call Database_SavePropertyWithIndex(rs1, "PredictedDispOpen_E", i, _
      prj.DispOpen(i).PredictedDispOpen_E)
  Next i
  
  Call Database_SaveProperty(rs1, "plottype", prj.plottype)
  
  'CLOSE THE DATABASE FILE.
  Db1.Close

  'COMPACT THE DATABASE FILE.
      'TO DO: USE THE DbEngine.CompactDatabase METHOD
      'TO COMPACT THE DATABASE.  PROBLEM TO CONSIDER:
      'THE DB MUST BE COMPACTED TO A TEMPORARY FILE,
      'WHICH THEN SHOULD OVERWRITE THE ORIGINAL FILE.
  
  'RETURN A "SUCCESS" MESSAGE TO CALLER.
  File_Save_Latest_v1_00 = True
  
End Function


Sub Units1_Database_SaveProperty(rs1 As Recordset, CboX As Control, Desc As String)
Dim OutStr As String
  If (CboX.ListIndex >= 0) Then
    OutStr = CboX.List(CboX.ListIndex)
  Else
    If (CboX.ListCount > 0) Then
      OutStr = CboX.List(0)
    Else
      OutStr = ""     'NOT LIKELY TO GET HERE!
    End If
  End If
  ''''Call ProjectFile_Write(f, OutStr, Desc)
  Call Database_SaveProperty(rs1, Desc, OutStr)
End Sub
Sub Units1_Database_LoadProperty(rs1 As Recordset, CboX As Control)
Dim TxtX As Control
Dim InLine As String
Dim Dummy1 As String
Dim NewUnits As String
Dim H As Integer
  Call Database_LoadProperty(rs1, InLine)
  ''''Call ProjectFile_Read(f, InLine, Dummy1)
  NewUnits = InLine
  H = unitsys_lookup_cbox(CboX)
  Set TxtX = unitsys(H).TxtX
  Call unitsys_set_units(TxtX, NewUnits)
End Sub


Sub Database_LoadProperty( _
    rs1 As Recordset, _
    LoadedData As Variant, _
    Optional Use_memoValue As Boolean = False)
  Select Case VarType(LoadedData)
    Case vbBoolean:
      LoadedData = CBool(Database_Get_Long(rs1, "lngValue"))
    Case vbByte:
      LoadedData = CByte(Database_Get_Long(rs1, "lngValue"))
    Case vbInteger:
      LoadedData = CInt(Database_Get_Long(rs1, "lngValue"))
    Case vbLong:
      LoadedData = CLng(Database_Get_Long(rs1, "lngValue"))
    Case vbString, vbDate:
      If (Use_memoValue) Then
        LoadedData = CStr(Database_Get_String(rs1, "memoValue"))
      Else
        LoadedData = CStr(Database_Get_String(rs1, "strValue"))
      End If
    Case vbDouble:
      LoadedData = CDbl(Database_Get_Double(rs1, "dblValue"))
    Case vbSingle:
      LoadedData = CSng(Database_Get_Double(rs1, "dblValue"))
  End Select
End Sub
Sub Database_SaveProperty( _
    rs1 As Recordset, _
    in_Use_FieldName As String, _
    SavedData As Variant)
Dim Use_memoValue As Boolean
Dim Use_FieldName As String
  'NOTE: IF THE FIRST CHARACTER IS AN ASTERISK ("*"), THEN
  'THE FIELD TYPE USED IS THE MEMO TYPE (memoValue).
  Use_memoValue = False
  If (left$(in_Use_FieldName, 1) = "*") Then
    Use_FieldName = Right$(in_Use_FieldName, Len(in_Use_FieldName) - 1)
    Use_memoValue = True
  Else
    Use_FieldName = in_Use_FieldName
  End If
  rs1.AddNew
  rs1("FieldName") = Use_FieldName
  Select Case VarType(SavedData)
    Case vbBoolean, vbByte, vbInteger, vbLong:
      rs1("lngValue") = CLng(SavedData)
    Case vbString, vbDate:
      If (Use_memoValue) Then
        rs1("memoValue") = CStr(SavedData)
      Else
        rs1("strValue") = CStr(SavedData)
      End If
    Case vbDouble, vbSingle:
      rs1("dblValue") = CDbl(SavedData)
  End Select
  rs1.Update
End Sub
Sub Database_SavePropertyWithIndex( _
    rs1 As Recordset, _
    in_Use_FieldName As String, _
    Use_FieldIndex As Integer, _
    SavedData As Variant)
Dim Use_memoValue As Boolean
Dim Use_FieldName As String
  'NOTE: IF THE FIRST CHARACTER IS AN ASTERISK ("*"), THEN
  'THE FIELD TYPE USED IS THE MEMO TYPE (memoValue).
  Use_memoValue = False
  If (left$(in_Use_FieldName, 1) = "*") Then
    Use_FieldName = Right$(in_Use_FieldName, Len(in_Use_FieldName) - 1)
    Use_memoValue = True
  Else
    Use_FieldName = in_Use_FieldName
  End If
  rs1.AddNew
  rs1("FieldName") = Use_FieldName
  rs1("FieldIndex") = Use_FieldIndex
  Select Case VarType(SavedData)
    Case vbBoolean, vbByte, vbInteger, vbLong:
      rs1("lngValue") = CLng(SavedData)
    Case vbString, vbDate:
      If (Use_memoValue) Then
        rs1("memoValue") = CStr(SavedData)
      Else
        rs1("strValue") = CStr(SavedData)
      End If
    Case vbDouble, vbSingle:
      rs1("dblValue") = CDbl(SavedData)
  End Select
  rs1.Update
End Sub


Sub Database_DeleteTableContents( _
    Db1 As Database, _
    TableName As String)
Dim rs1 As Recordset
  On Error GoTo err_Database_DeleteTableContents
  Set rs1 = Db1.OpenRecordset(TableName)
  rs1.MoveFirst
  Do Until rs1.EOF
    rs1.Delete
    rs1.MoveNext
  Loop
  rs1.Close
  Exit Sub
exit_err_Database_DeleteTableContents:
  Exit Sub
err_Database_DeleteTableContents:
  Resume exit_err_Database_DeleteTableContents
End Sub
Sub Database_CreateMFBTable( _
    Db1 As Database, _
    TableName As String, _
    Include_FieldIndex As Boolean, _
    Include_FieldIndex2 As Boolean)
Dim Td1 As TableDef
Dim Ff As Field
    
  Set Td1 = Db1.CreateTableDef(TableName)
  Set Ff = Td1.CreateField("RecordID", dbLong):
  'TODO: ADD AUTONUMBER SETUP FOR THIS FIELD (NOT TOO NEEDED).
  Td1.Fields.Append Ff
  Set Ff = Td1.CreateField("FieldName", dbText, 250):
  Ff.AllowZeroLength = True
  Td1.Fields.Append Ff
  If (Include_FieldIndex) Then
    Set Ff = Td1.CreateField("FieldIndex", dbLong):
    Td1.Fields.Append Ff
  End If
  If (Include_FieldIndex2) Then
    Set Ff = Td1.CreateField("FieldIndex2", dbLong):
    Td1.Fields.Append Ff
  End If
  Set Ff = Td1.CreateField("strValue", dbText, 250):
  Ff.AllowZeroLength = True
  Td1.Fields.Append Ff
  Set Ff = Td1.CreateField("dblValue", dbDouble):
  Td1.Fields.Append Ff
  Set Ff = Td1.CreateField("lngValue", dbLong):
  Td1.Fields.Append Ff
  Set Ff = Td1.CreateField("memoValue", dbMemo):
  Ff.AllowZeroLength = True
  Td1.Fields.Append Ff
  Set Ff = Td1.CreateField("Comments", dbText, 250):
  Ff.AllowZeroLength = True
  Td1.Fields.Append Ff
  Db1.TableDefs.Append Td1
End Sub
Sub Database_CreateMFBTable_IfNoExist( _
    Db1 As Database, _
    Use_TableName As String, _
    Include_FieldIndex As Boolean)
  If (Database_IsTableExist(Db1, Use_TableName) = False) Then
    Call Database_CreateMFBTable(Db1, Use_TableName, Include_FieldIndex, False)
  End If
End Sub
Sub Database_CreateMFBTable_IfNoExist_TwoIndices( _
    Db1 As Database, _
    Use_TableName As String)
  If (Database_IsTableExist(Db1, Use_TableName) = False) Then
    Call Database_CreateMFBTable(Db1, Use_TableName, True, True)
  End If
End Sub




