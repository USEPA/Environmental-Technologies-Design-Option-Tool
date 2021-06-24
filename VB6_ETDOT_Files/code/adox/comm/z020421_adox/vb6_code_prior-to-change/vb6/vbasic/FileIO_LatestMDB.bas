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
Dim Rs1 As Recordset
Dim Use_FieldIndex As Integer
Dim Use_FieldIndex2 As Integer
Dim ContainsTable_PSDMInRoomData As Boolean
Dim Prj As Project_Type
Dim i As Integer
Dim j As Integer

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
    Set Rs1 = Db1.OpenRecordset("Version")
    If (Database_NoRecordsInRecordset(Rs1)) Then
      'DO NOTHING: USE DEFAULT VALUES.
    Else
      Rs1.MoveFirst
      Do Until Rs1.EOF
        ''''Use_FieldIndex = CInt(Database_Get_Long(Rs1, "FieldIndex"))
        Select Case Trim$(UCase$(Database_Get_String(Rs1, "FieldName")))
          'HEADER BLOCK.
          Case Trim$(UCase$("ContainsTable_PSDMInRoomData")): Call Database_LoadProperty(Rs1, ContainsTable_PSDMInRoomData)
        End Select
        Rs1.MoveNext
      Loop
    End If
    Rs1.Close
  End If
  
  '------ INPUT DATA FROM TABLE "Main". ------------------------------------------------------------------------------------------------------
  If (Database_IsTableExist(Db1, "Main") = False) Then
    'DO NOTHING: USE DEFAULT VALUES.
  Else
    Set Rs1 = Db1.OpenRecordset("Main")
    If (Database_NoRecordsInRecordset(Rs1)) Then
      'DO NOTHING: USE DEFAULT VALUES.
    Else
      '
      ' SET DEFAULT PROJECT DATA IN TEMPORARY VARIABLE.
      '
      Call Project_SetDefaults(Prj)
      '
      ' READ IN THE PROJECT DATA TO TEMPORARY VARIABLE.
      '
      Rs1.MoveFirst
      Do Until Rs1.EOF
        Use_FieldIndex = CInt(Database_Get_Long(Rs1, "FieldIndex"))
        Select Case Trim$(UCase$(Database_Get_String(Rs1, "FieldName")))
          'MAIN BLOCK.
  
           'REACTOR PROPERTIES.
            Case Trim$(UCase$("idreact")): Call Database_LoadProperty(Rs1, Prj.idreact)
            Case Trim$(UCase$("volume")): Call Database_LoadProperty(Rs1, Prj.volume)
            Case Trim$(UCase$("tau")): Call Database_LoadProperty(Rs1, Prj.tau)
            
            'NUMERICAL SIMULATION PARAMETERS.
            Case Trim$(UCase$("ssize")): Call Database_LoadProperty(Rs1, Prj.ssize)
            Case Trim$(UCase$("ttotal")): Call Database_LoadProperty(Rs1, Prj.ttotal)
            Case Trim$(UCase$("opsize")): Call Database_LoadProperty(Rs1, Prj.opsize)
            Case Trim$(UCase$("xntimes")): Call Database_LoadProperty(Rs1, Prj.xntimes)
          
            'WATER QUALITY PROPERTIES.
            Case Trim$(UCase$("ph0")): Call Database_LoadProperty(Rs1, Prj.ph0)
            Case Trim$(UCase$("phosph")): Call Database_LoadProperty(Rs1, Prj.phosph)
            Case Trim$(UCase$("idcarbn")): Call Database_LoadProperty(Rs1, Prj.idcarbn)
            Case Trim$(UCase$("alk")): Call Database_LoadProperty(Rs1, Prj.alk)
            Case Trim$(UCase$("ticarbn")): Call Database_LoadProperty(Rs1, Prj.ticarbn)
            Case Trim$(UCase$("inf_h2o2")): Call Database_LoadProperty(Rs1, Prj.inf_h2o2)
            
            'TARGET COMPOUNDS.
            Case Trim$(UCase$("TargetCompounds_Count")): Call Database_LoadProperty(Rs1, Prj.TargetCompounds_Count)
            If (Prj.TargetCompounds_Count > 0) Then
              ReDim Prj.TargetCompounds(1 To Prj.TargetCompounds_Count)
              Rs1.MoveNext
              For i = 1 To Prj.TargetCompounds_Count
                Do While i = Trim$(UCase$(Database_Get_Integer(Rs1, "FieldIndex")))
                  Select Case Trim$(UCase$(Database_Get_String(Rs1, "FieldName")))
                    Case Trim$(UCase$("comname")): Call Database_LoadProperty(Rs1, Prj.TargetCompounds(i).comname)
                    Case Trim$(UCase$("cas")): Call Database_LoadProperty(Rs1, Prj.TargetCompounds(i).cas)
                    'PROPERTIES OF THIS COMPOUND.
                    Case Trim$(UCase$("concini")): Call Database_LoadProperty(Rs1, Prj.TargetCompounds(i).concini)
                    Case Trim$(UCase$("val")): Call Database_LoadProperty(Rs1, Prj.TargetCompounds(i).val)
                    Case Trim$(UCase$("mw")): Call Database_LoadProperty(Rs1, Prj.TargetCompounds(i).mw)
                    Case Trim$(UCase$("ncarbn")): Call Database_LoadProperty(Rs1, Prj.TargetCompounds(i).ncarbn)
                    Case Trim$(UCase$("nsubstt")): Call Database_LoadProperty(Rs1, Prj.TargetCompounds(i).nsubstt)
                    Case Trim$(UCase$("xk")): Call Database_LoadProperty(Rs1, Prj.TargetCompounds(i).xk)
                    'PROPERTIES OF THE DEPROTONATED COMPOUND (NOT APPLICABLE TO NOM).
                    Case Trim$(UCase$("dep_comname")): Call Database_LoadProperty(Rs1, Prj.TargetCompounds(i).dep_comname)
                    Case Trim$(UCase$("dep_val")): Call Database_LoadProperty(Rs1, Prj.TargetCompounds(i).dep_val)
                    Case Trim$(UCase$("dep_mw")): Call Database_LoadProperty(Rs1, Prj.TargetCompounds(i).dep_mw)
                    Case Trim$(UCase$("dep_xk")): Call Database_LoadProperty(Rs1, Prj.TargetCompounds(i).dep_xk)
                    Case Trim$(UCase$("dep_xke")): Call Database_LoadProperty(Rs1, Prj.TargetCompounds(i).dep_xke)
                    'OTHER RATE CONSTANTS.
                    Case Trim$(UCase$("xk_co3XM")): Call Database_LoadProperty(Rs1, Prj.TargetCompounds(i).xk_co3XM)
                    Case Trim$(UCase$("xk_hpo4XM")): Call Database_LoadProperty(Rs1, Prj.TargetCompounds(i).xk_hpo4XM)
                    Case Trim$(UCase$("xk_o2XM")): Call Database_LoadProperty(Rs1, Prj.TargetCompounds(i).xk_o2XM)
                    Case Trim$(UCase$("xk_ho2X")): Call Database_LoadProperty(Rs1, Prj.TargetCompounds(i).xk_ho2X)
                  End Select
                  Rs1.MoveNext
                  If Trim$(UCase$(Database_Get_Integer(Rs1, "FieldIndex"))) = 0 Then
                    Rs1.MovePrevious
                    Exit Do
                  End If
                  Loop
              Next i
            End If
            
            'PHOTOCHEMICAL PARAMETERS.
            Case Trim$(UCase$("Wavelength_Count")): Call Database_LoadProperty(Rs1, Prj.Wavelength_Count)
            If (Prj.Wavelength_Count > 0) Then
              ReDim Prj.Wavelengths(1 To Prj.Wavelength_Count)
              Rs1.MoveNext
              For i = 1 To Prj.Wavelength_Count
                Do While i = Trim$(UCase$(Database_Get_Integer(Rs1, "FieldIndex"))) _
                  And (Trim$(UCase$(Database_Get_String(Rs1, "FieldName"))) = "LWAVE" _
                  Or Trim$(UCase$(Database_Get_String(Rs1, "FieldName"))) = "UVI")
                  
                  Select Case Trim$(UCase$(Database_Get_String(Rs1, "FieldName")))
                    Case Trim$(UCase$("lwave")): Call Database_LoadProperty(Rs1, Prj.Wavelengths(i).lwave)
                    Case Trim$(UCase$("uvi")): Call Database_LoadProperty(Rs1, Prj.Wavelengths(i).uvi)
                  End Select
                  Rs1.MoveNext
                  Loop
              Next i
            End If
          
            If (Prj.TargetCompounds_Count > 0) And _
               (Prj.Wavelength_Count > 0) Then
              ReDim Prj.extcoef(1 To Prj.TargetCompounds_Count, 1 To Prj.Wavelength_Count)
              For i = 1 To Prj.TargetCompounds_Count
                For j = 1 To Prj.Wavelength_Count
                  Do While i = Trim$(UCase$(Database_Get_Integer(Rs1, "FieldIndex"))) And _
                  j = Trim$(UCase$(Database_Get_Integer(Rs1, "FieldIndex2")))
                    Select Case Trim$(UCase$(Database_Get_String(Rs1, "FieldName")))
                      Case Trim$(UCase$("extcoef")): Call Database_LoadProperty(Rs1, Prj.extcoef(i, j))
                    End Select
                    Rs1.MoveNext
                    Loop
                Next j
              Next i
              ReDim Prj.quatyd(1 To Prj.TargetCompounds_Count, 1 To Prj.Wavelength_Count)
              For i = 1 To Prj.TargetCompounds_Count
                For j = 1 To Prj.Wavelength_Count
                  Do While i = Trim$(UCase$(Database_Get_Integer(Rs1, "FieldIndex"))) And _
                  j = Trim$(UCase$(Database_Get_Integer(Rs1, "FieldIndex2")))
                    Select Case Trim$(UCase$(Database_Get_String(Rs1, "FieldName")))
                      Case Trim$(UCase$("quatyd")): Call Database_LoadProperty(Rs1, Prj.quatyd(i, j))
                    End Select
                    Rs1.MoveNext
                    Loop
                Next j
              Next i
            End If
            If (Prj.Wavelength_Count > 0) Then
              ReDim Prj.extcoef_h2o2(1 To Prj.Wavelength_Count)
              For j = 1 To Prj.Wavelength_Count
                Select Case Trim$(UCase$(Database_Get_String(Rs1, "FieldName")))
                  Case Trim$(UCase$("extcoef_h2o2")): Call Database_LoadProperty(Rs1, Prj.extcoef_h2o2(j))
                End Select
                Rs1.MoveNext
              Next j
              ReDim Prj.quatyd_h2o2(1 To Prj.Wavelength_Count)
              For j = 1 To Prj.Wavelength_Count
                  Select Case Trim$(UCase$(Database_Get_String(Rs1, "FieldName")))
                    Case Trim$(UCase$("quatyd_h2o2")): Call Database_LoadProperty(Rs1, Prj.quatyd_h2o2(j))
                  End Select
                  Rs1.MoveNext
              Next j
            End If
            Rs1.MovePrevious
            Case Trim$(UCase$("uvpathl")): Call Database_LoadProperty(Rs1, Prj.uvpathl)
            Case Trim$(UCase$("lamp_name")): Call Database_LoadProperty(Rs1, Prj.lamp_name)
            Case Trim$(UCase$("lamp_power")): Call Database_LoadProperty(Rs1, Prj.lamp_power)
            Case Trim$(UCase$("iduvi")): Call Database_LoadProperty(Rs1, Prj.iduvi)
            Case Trim$(UCase$("num_tanks")): Call Database_LoadProperty(Rs1, Prj.num_tanks)
            Case Trim$(UCase$("frmMain.cboUnits(0)")): Call Units1_Database_LoadProperty(Rs1, frmMain.cboUnits(0))
            Case Trim$(UCase$("frmMain.cboUnits(1)")): Call Units1_Database_LoadProperty(Rs1, frmMain.cboUnits(1))
            Case Trim$(UCase$("frmMain.cboUnits(2)")): Call Units1_Database_LoadProperty(Rs1, frmMain.cboUnits(2))
            Case Trim$(UCase$("frmMain.cboUnits(3)")): Call Units1_Database_LoadProperty(Rs1, frmMain.cboUnits(3))
            Rs1.MoveNext
            
'''            'DYE STUDY
'''            Dim f As Integer
'''            Dim s As String
'''            Dim fn_outputtxt As String
'''
'''            Case Trim$(UCase$("DyeStudy_Count")): Call Database_LoadProperty(Rs1, Prj.dyestudy_count)
'''            If (Prj.dyestudy_count > 0) Then
'''              ReDim Prj.DyeStudy(1 To Prj.dyestudy_count)
'''              Rs1.MoveNext
'''              For i = 1 To Prj.dyestudy_count
'''                  Do While i = Trim$(UCase$(Database_Get_Integer(Rs1, "FieldIndex")))
'''                    Select Case Trim$(UCase$(Database_Get_String(Rs1, "FieldName")))
'''                      Case Trim$(UCase$("time")): Call Database_LoadProperty(Rs1, Prj.DyeStudy(i).time)
'''                      Case Trim$(UCase$("concentration")): Call Database_LoadProperty(Rs1, Prj.DyeStudy(i).concentration)
'''                    End Select
'''                    Rs1.MoveNext
'''                    Loop
'''                Next i
'''              Rs1.MovePrevious
'''            End If
            
'''            Case Trim$(UCase$("DyeStudy_Output")): Call Database_LoadProperty(Rs1, Prj.dyestudy_output, True)
'''            f = FreeFile
''''            fn_OutputTxt = App.Path + "\exes\output.txt"
'''            fn_outputtxt = App.Path + "\exes\outpt.txt"
'''            Open fn_outputtxt For Output As #f
'''            Print #f, Prj.dyestudy_output
'''            Close #f
'''
'''            Case Trim$(UCase$("DyeStudy_Calcdate")): Call Database_LoadProperty(Rs1, Prj.dyestudy_calcdate)
        
         End Select
        
        Rs1.MoveNext
        
      Loop
      '
      ' TRANSFER PROJECT DATA TO MEMORY.
      '
      NowProj = Prj
    End If
    Rs1.Close
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
Dim Rs1 As Recordset
Dim i As Integer
Dim j As Integer
Dim IsInvalidFormat As Boolean
Dim NeedToCreateNewDatabase As Boolean
Dim Prj As Project_Type

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
    FileCopy MAIN_APP_PATH & "\dbase\template.adx", fn_this
'    Set Db1 = CreateDatabase(fn_this, dbLangGeneral)
    Set Db1 = OpenDatabase(fn_this)
  End If
  'CREATE NEW TABLES WITHIN DATABASE, IF NECESSARY.
  Call Database_CreateMFBTable_IfNoExist_TwoIndices(Db1, "Version")
  Call Database_CreateMFBTable_IfNoExist_TwoIndices(Db1, "Main")
  
  '=========== OUTPUT DATA TO DATABASE TABLES. =================
  
  '------ OUTPUT DATA TO TABLE "Version". ------------------------------------------------------------------------------------------------------
  'START SAVE TO THIS TABLE.
  Call Database_DeleteTableContents(Db1, "Version")
  Set Rs1 = Db1.OpenRecordset("Version")
  Call Database_SaveProperty(Rs1, "DataVersion_Major", CInt(Latest_DataVersion_Major))
  Call Database_SaveProperty(Rs1, "DataVersion_Minor", CInt(Latest_DataVersion_Minor))
  ''''Call Database_SaveProperty(Rs1, "ContainsTable_PSDMInRoomData", True)
  'END SAVE TO THIS TABLE.
  Rs1.Close
  
  '------ OUTPUT DATA TO TABLE "Main". ------------------------------------------------------------------------------------------------------
  'START SAVE TO THIS TABLE.
  Call Database_DeleteTableContents(Db1, "Main")
  Set Rs1 = Db1.OpenRecordset("Main")
  'MAIN BLOCK.
  Prj = NowProj
  'Call Database_SaveProperty(Rs1, "Length", Prj.length)
  'Call Database_SaveProperty(Rs1, "Diameter", Prj.Diameter)
  'Call Database_SaveProperty(Rs1, "Mass", Prj.Mass)
  'Call Database_SaveProperty(Rs1, "FlowRate", Prj.FlowRate)
  
  'REACTOR PROPERTIES.
  Call Database_SaveProperty(Rs1, "idreact", Prj.idreact)
  Call Database_SaveProperty(Rs1, "volume", Prj.volume)
'  For i = 0 To 2
'    Call Database_SavePropertyWithIndex(Rs1, "Iw.UnitsOfDisplay", i, Iw.UnitsOfDisplay(i))
 ' Next i
  Call Database_SaveProperty(Rs1, "tau", Prj.tau)
'  For i = 0 To 2
'    Call Database_SavePropertyWithIndex(Rs1, "Iw.UnitsOfDisplay", i, Iw.UnitsOfDisplay(i))
'  Next i
  
  'NUMERICAL SIMULATION PARAMETERS.
  Call Database_SaveProperty(Rs1, "ssize", Prj.ssize)
  Call Database_SaveProperty(Rs1, "ttotal", Prj.ttotal)
  Call Database_SaveProperty(Rs1, "opsize", Prj.opsize)
  Call Database_SaveProperty(Rs1, "xntimes", Prj.xntimes)
  
  'WATER QUALITY PROPERTIES.
  Call Database_SaveProperty(Rs1, "ph0", Prj.ph0)
  Call Database_SaveProperty(Rs1, "phosph", Prj.phosph)
  Call Database_SaveProperty(Rs1, "idcarbn", Prj.idcarbn)
  Call Database_SaveProperty(Rs1, "alk", Prj.alk)
  Call Database_SaveProperty(Rs1, "ticarbn", Prj.ticarbn)
  Call Database_SaveProperty(Rs1, "inf_h2o2", Prj.inf_h2o2)
  
  'TARGET COMPOUNDS.
  Call Database_SaveProperty(Rs1, "TargetCompounds_Count", Prj.TargetCompounds_Count)
  For i = 1 To Prj.TargetCompounds_Count
    Call Database_SavePropertyWithIndex(Rs1, "comname", i, Prj.TargetCompounds(i).comname)
    Call Database_SavePropertyWithIndex(Rs1, "cas", i, Prj.TargetCompounds(i).cas)
    Call Database_SavePropertyWithIndex(Rs1, "concini", i, Prj.TargetCompounds(i).concini)
    Call Database_SavePropertyWithIndex(Rs1, "val", i, Prj.TargetCompounds(i).val)
    Call Database_SavePropertyWithIndex(Rs1, "mw", i, Prj.TargetCompounds(i).mw)
    Call Database_SavePropertyWithIndex(Rs1, "ncarbn", i, Prj.TargetCompounds(i).ncarbn)
    Call Database_SavePropertyWithIndex(Rs1, "nsubstt", i, Prj.TargetCompounds(i).nsubstt)
    Call Database_SavePropertyWithIndex(Rs1, "xk", i, Prj.TargetCompounds(i).xk)
    Call Database_SavePropertyWithIndex(Rs1, "dep_comname", i, Prj.TargetCompounds(i).dep_comname)
    Call Database_SavePropertyWithIndex(Rs1, "dep_val", i, Prj.TargetCompounds(i).dep_val)
    Call Database_SavePropertyWithIndex(Rs1, "dep_mw", i, Prj.TargetCompounds(i).dep_mw)
    Call Database_SavePropertyWithIndex(Rs1, "dep_xk", i, Prj.TargetCompounds(i).dep_xk)
    Call Database_SavePropertyWithIndex(Rs1, "dep_xke", i, Prj.TargetCompounds(i).dep_xke)
    Call Database_SavePropertyWithIndex(Rs1, "xk_co3XM", i, Prj.TargetCompounds(i).xk_co3XM)
    Call Database_SavePropertyWithIndex(Rs1, "xk_hpo4XM", i, Prj.TargetCompounds(i).xk_hpo4XM)
    Call Database_SavePropertyWithIndex(Rs1, "xk_o2XM", i, Prj.TargetCompounds(i).xk_o2XM)
    Call Database_SavePropertyWithIndex(Rs1, "xk_ho2X", i, Prj.TargetCompounds(i).xk_ho2X)
  Next i
  
  'PHOTOCHEMICAL PARAMETERS.
   Call Database_SaveProperty(Rs1, "Wavelength_Count", Prj.Wavelength_Count)
  For i = 1 To Prj.Wavelength_Count
    Call Database_SavePropertyWithIndex(Rs1, "lwave", i, Prj.Wavelengths(i).lwave)
    Call Database_SavePropertyWithIndex(Rs1, "uvi", i, Prj.Wavelengths(i).uvi)
  Next i
  For i = 1 To Prj.TargetCompounds_Count
    For j = 1 To Prj.Wavelength_Count
      Call Database_SavePropertyWithTwoIndeces(Rs1, _
        "extcoef", i, j, Prj.extcoef(i, j))
    Next j
  Next i
  For i = 1 To Prj.TargetCompounds_Count
    For j = 1 To Prj.Wavelength_Count
      Call Database_SavePropertyWithTwoIndeces(Rs1, _
        "quatyd", i, j, Prj.quatyd(i, j))
    Next j
  Next i
  For j = 1 To Prj.Wavelength_Count
    Call Database_SavePropertyWithIndex(Rs1, "extcoef_h2o2", j, Prj.extcoef_h2o2(j))
  Next j
  For j = 1 To Prj.Wavelength_Count
    Call Database_SavePropertyWithIndex(Rs1, "quatyd_h2o2", j, Prj.quatyd_h2o2(j))
  Next j
  Call Database_SaveProperty(Rs1, "uvpathl", Prj.uvpathl)
  Call Database_SaveProperty(Rs1, "lamp_name", Prj.lamp_name)
  Call Database_SaveProperty(Rs1, "lamp_power", Prj.lamp_power)
  Call Database_SaveProperty(Rs1, "iduvi", Prj.iduvi)
  Call Database_SaveProperty(Rs1, "num_tanks", Prj.num_tanks)
  
  'DYE STUDY PARAMETERS.
'''  Call Database_SaveProperty(Rs1, "DyeStudy_Count", Prj.dyestudy_count)
'''  For i = 1 To Prj.dyestudy_count
'''    Call Database_SavePropertyWithIndex(Rs1, "time", i, Prj.DyeStudy(i).time)
'''    Call Database_SavePropertyWithIndex(Rs1, "concentration", i, Prj.DyeStudy(i).concentration)
'''  Next i
'''
'''  Call Database_SaveProperty(Rs1, "*DyeStudy_Output", Prj.dyestudy_output)
'''  Call Database_SaveProperty(Rs1, "DyeStudy_Calcdate", Prj.dyestudy_calcdate)
  
  'END SAVE TO THIS TABLE.
  Rs1.Close
  
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


Sub Units1_Database_SaveProperty(Rs1 As Recordset, CboX As Control, Desc As String)
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
  Call Database_SaveProperty(Rs1, Desc, OutStr)
End Sub
Sub Units1_Database_LoadProperty(Rs1 As Recordset, CboX As Control)
Dim TxtX As Control
Dim InLine As String
Dim Dummy1 As String
Dim NewUnits As String
Dim H As Integer
  Call Database_LoadProperty(Rs1, InLine)
  NewUnits = InLine
  H = unitsys_lookup_cbox(CboX)
  Set TxtX = unitsys(H).TxtX
  Call unitsys_set_units(TxtX, NewUnits)
End Sub


Sub Database_LoadProperty( _
    Rs1 As Recordset, _
    LoadedData As Variant, _
    Optional Use_memoValue As Boolean = False)
  Select Case VarType(LoadedData)
    Case vbBoolean:
      LoadedData = CBool(Database_Get_Long(Rs1, "lngValue"))
    Case vbByte:
      LoadedData = CByte(Database_Get_Long(Rs1, "lngValue"))
    Case vbInteger:
      LoadedData = CInt(Database_Get_Long(Rs1, "lngValue"))
    Case vbLong:
      LoadedData = CLng(Database_Get_Long(Rs1, "lngValue"))
    Case vbString, vbDate:
      If (Use_memoValue) Then
        LoadedData = CStr(Database_Get_String(Rs1, "memoValue"))
      Else
        LoadedData = CStr(Database_Get_String(Rs1, "strValue"))
      End If
    Case vbDouble:
      LoadedData = CDbl(Database_Get_Double(Rs1, "dblValue"))
    Case vbSingle:
      LoadedData = CSng(Database_Get_Double(Rs1, "dblValue"))
  End Select
End Sub
Sub Database_SaveProperty( _
    Rs1 As Recordset, _
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
  Rs1.AddNew
  Rs1("FieldName") = Use_FieldName
  Select Case VarType(SavedData)
    Case vbBoolean, vbByte, vbInteger, vbLong:
      Rs1("lngValue") = CLng(SavedData)
    Case vbString, vbDate:
      If (Use_memoValue) Then
        Rs1("memoValue") = CStr(SavedData)
      Else
        Rs1("strValue") = CStr(SavedData)
      End If
    Case vbDouble, vbSingle:
      Rs1("dblValue") = CDbl(SavedData)
  End Select
  Rs1.Update
End Sub
Sub Database_SavePropertyWithIndex( _
    Rs1 As Recordset, _
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
  Rs1.AddNew
  Rs1("FieldName") = Use_FieldName
  Rs1("FieldIndex") = Use_FieldIndex
  Select Case VarType(SavedData)
    Case vbBoolean, vbByte, vbInteger, vbLong:
      Rs1("lngValue") = CLng(SavedData)
    Case vbString, vbDate:
      If (Use_memoValue) Then
        Rs1("memoValue") = CStr(SavedData)
      Else
        Rs1("strValue") = CStr(SavedData)
      End If
    Case vbDouble, vbSingle:
      Rs1("dblValue") = CDbl(SavedData)
  End Select
  Rs1.Update
End Sub

Sub Database_SavePropertyWithTwoIndeces( _
    Rs1 As Recordset, _
    in_Use_FieldName As String, _
    Use_FieldIndex As Integer, _
    Use_FieldIndex2 As Integer, _
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
  Rs1.AddNew
  Rs1("FieldName") = Use_FieldName
  Rs1("FieldIndex") = Use_FieldIndex
  Rs1("FieldIndex2") = Use_FieldIndex2
  Select Case VarType(SavedData)
    Case vbBoolean, vbByte, vbInteger, vbLong:
      Rs1("lngValue") = CLng(SavedData)
    Case vbString, vbDate:
      If (Use_memoValue) Then
        Rs1("memoValue") = CStr(SavedData)
      Else
        Rs1("strValue") = CStr(SavedData)
      End If
    Case vbDouble, vbSingle:
      Rs1("dblValue") = CDbl(SavedData)
  End Select
  Rs1.Update
End Sub
Sub Database_DeleteTableContents( _
    Db1 As Database, _
    TableName As String)
Dim Rs1 As Recordset
  On Error GoTo err_Database_DeleteTableContents
  Set Rs1 = Db1.OpenRecordset(TableName)
  Rs1.MoveFirst
  Do Until Rs1.EOF
    Rs1.Delete
    Rs1.MoveNext
  Loop
  Rs1.Close
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



Sub GridFunc_Convert_CommasToLF_OneLine( _
    OldLine As String, _
    NewLine As String _
    )
Dim WorkingStr As String
Dim NextStr As String
Dim NextPos As Integer
Dim ThisIter As Integer
  WorkingStr = OldLine
  ThisIter = 0
  Do While (1 = 1)
    NextPos = InStr(WorkingStr, ",")
    If (NextPos = 0) Then Exit Do
    If (NextPos > 1) Then
      NextStr = left$(WorkingStr, NextPos - 1)
    Else
      NextStr = ""
    End If
    NextStr = NextStr & Chr$(10)     'LF character
    If (NextPos < Len(WorkingStr)) Then
      NextStr = NextStr & Right$(WorkingStr, Len(WorkingStr) - NextPos)
    End If
    WorkingStr = NextStr
    ThisIter = ThisIter + 1
    If (ThisIter > 100) Then Exit Do
  Loop
  NewLine = WorkingStr
End Sub

