Attribute VB_Name = "FileIO_LatestMDB"
Option Explicit





Const FileIO_LatestMDB_declarations_end = True


'RETURNS:
'         TRUE = SUCCEEDED IN LOADING.
'         FALSE = FAILED IN LOADING.
Function File_Open_Latest_v1_00( _
    fn_This As String) As Boolean
Dim Ws1 As Workspace
Dim Db1 As Database
Dim Rs1 As Recordset
Dim Use_FieldIndex As Integer
Dim Use_FieldIndex2 As Integer
Dim Use_FieldIndex3 As Integer
Dim Use_FieldIndex4 As Integer
Dim ContainsTable_PSDMInRoomData As Boolean
Dim Prj As Project_Type
Dim FieldName As String
Dim SearchFor As String
Dim UB As Integer
Dim i As Integer


'>>>>>>>>>>>>>>>>>>>> *TODO* UPDATE THIS ENTIRE FUNCTION <<<<<<<<<<<<<<<<<<<<<<<<<

  '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
  '>>>>>>>>>>>>>>>>>>>>>>>>>>>  INPUT FROM MAIN DATABASE  <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
  '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
  If (Not FileExists(fn_This)) Then
    'ERROR: UNABLE TO FIND THE FILE!
    File_Open_Latest_v1_00 = False
    Exit Function
  End If
  'OPEN DATABASE.
  Set Db1 = OpenDatabase(fn_This)

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
      With Prj
        Do Until Rs1.EOF
          Use_FieldIndex = CInt(Database_Get_Long(Rs1, "FieldIndex"))
          Select Case Trim$(UCase$(Database_Get_String(Rs1, "FieldName")))
            '
            ' MISCELLANEOUS FILE DATA.
            '
            Case Trim$(UCase$("File_Note")): Call Database_LoadProperty(Rs1, .File_Note)
            '
            ' HIERARCHY RELATED DATA.
            '
            ''''UserHierarchy As UserHierarchy_Type
            '
            ' MAIN DATA SET.
            '
            Case Trim$(UCase$("Op_T")): Call Database_LoadProperty(Rs1, .Op_T)
            Case Trim$(UCase$("Op_P")): Call Database_LoadProperty(Rs1, .Op_P)
            Case Trim$(UCase$("Op_T_UnitDisplayed")): Call Database_LoadProperty(Rs1, .Op_T_UnitDisplayed)
            Case Trim$(UCase$("Op_P_UnitDisplayed")): Call Database_LoadProperty(Rs1, .Op_P_UnitDisplayed)
            ''''UserChemicals() As UserChemical_Type
          End Select
          Rs1.MoveNext
        Loop
        '
        ' TRANSFER PROJECT DATA TO MEMORY.
        '
        NowProj = Prj
      End With
    End If
    Rs1.Close
  End If

  '------ INPUT DATA FROM TABLE "UserHierarchy". ------------------------------------------------------------------------------------------------------
  If (Database_IsTableExist(Db1, "UserHierarchy") = False) Then
    'DO NOTHING: USE DEFAULT VALUES.
  Else
    Set Rs1 = Db1.OpenRecordset("UserHierarchy")
    If (Database_NoRecordsInRecordset(Rs1)) Then
      'DO NOTHING: USE DEFAULT VALUES.
    Else
      '
      ' SET DEFAULT PROJECT DATA IN TEMPORARY VARIABLE.
      '
      ReDim NowProj.UserHierarchy.PropertySheetOrder(0 To 0)
      '
      ' READ IN THE PROJECT DATA TO TEMPORARY VARIABLE.
      '
      Rs1.MoveFirst
      With NowProj.UserHierarchy
        Do Until Rs1.EOF
          Use_FieldIndex = CInt(Database_Get_Long(Rs1, "FieldIndex"))
          Use_FieldIndex2 = CInt(Database_Get_Long(Rs1, "FieldIndex2"))
          Use_FieldIndex3 = CInt(Database_Get_Long(Rs1, "FieldIndex3"))
          FieldName = Trim$(UCase$(Database_Get_String(Rs1, "FieldName")))
          SearchFor = "PropertySheetOrder."
          If (Trim$(UCase$(Left$(FieldName, Len(SearchFor)))) = Trim$(UCase$(SearchFor))) Then
            UB = UBound(.PropertySheetOrder)
            If (Use_FieldIndex > UB) Then
              If (UB = 0) Then
                ReDim .PropertySheetOrder(1 To Use_FieldIndex)
              Else
                ReDim Preserve .PropertySheetOrder(1 To Use_FieldIndex)
              End If
              For i = UB + 1 To Use_FieldIndex
                ReDim .PropertySheetOrder(i).PropertyOrder(0 To 0)
              Next i
            End If
          End If
          SearchFor = "PropertySheetOrder.PropertyOrder."
          If (Trim$(UCase$(Left$(FieldName, Len(SearchFor)))) = Trim$(UCase$(SearchFor))) Then
            UB = UBound(.PropertySheetOrder(Use_FieldIndex).PropertyOrder)
            'UBound(.PropertySheetOrder)
            If (Use_FieldIndex2 > UB) Then
              If (UB = 0) Then
                ReDim .PropertySheetOrder(Use_FieldIndex).PropertyOrder(1 To Use_FieldIndex2)
              Else
                ReDim Preserve .PropertySheetOrder(Use_FieldIndex).PropertyOrder(1 To Use_FieldIndex2)
              End If
              For i = UB + 1 To Use_FieldIndex2
                ReDim .PropertySheetOrder(Use_FieldIndex). _
                    PropertyOrder(i).Technique_Code(0 To 0)
              Next i
            End If
          End If
          SearchFor = "PropertySheetOrder.PropertyOrder.Technique_Code"
          If (Trim$(UCase$(Left$(FieldName, Len(SearchFor)))) = Trim$(UCase$(SearchFor))) Then
            UB = UBound(.PropertySheetOrder(Use_FieldIndex).PropertyOrder(Use_FieldIndex2).Technique_Code)
            If (Use_FieldIndex3 > UB) Then
              If (UB = 0) Then
                ReDim .PropertySheetOrder(Use_FieldIndex).PropertyOrder(Use_FieldIndex2).Technique_Code(1 To Use_FieldIndex3)
              Else
                ReDim Preserve .PropertySheetOrder(Use_FieldIndex).PropertyOrder(Use_FieldIndex2).Technique_Code(1 To Use_FieldIndex3)
              End If
            End If
          End If
          Select Case FieldName
            Case Trim$(UCase$("PropertySheetOrder.Name")):
              Call Database_LoadProperty(Rs1, .PropertySheetOrder(Use_FieldIndex).Name)
            Case Trim$(UCase$("PropertySheetOrder.PropertyOrder.Property_Code")):
              Call Database_LoadProperty(Rs1, .PropertySheetOrder(Use_FieldIndex). _
                  PropertyOrder(Use_FieldIndex2).Property_Code)
            Case Trim$(UCase$("PropertySheetOrder.PropertyOrder.Technique_Code")):
              Call Database_LoadProperty(Rs1, .PropertySheetOrder(Use_FieldIndex). _
                  PropertyOrder(Use_FieldIndex2).Technique_Code(Use_FieldIndex3))
          End Select
          Rs1.MoveNext
        Loop
        '''''
        ''''' TRANSFER PROJECT DATA TO MEMORY.
        '''''
        ''''NowProj = Prj
      End With
    End If
    Rs1.Close
  End If

  '------ INPUT DATA FROM TABLE "UserChemicals". ------------------------------------------------------------------------------------------------------
  If (Database_IsTableExist(Db1, "UserChemicals") = False) Then
    'DO NOTHING: USE DEFAULT VALUES.
  Else
    Set Rs1 = Db1.OpenRecordset("UserChemicals")
    If (Database_NoRecordsInRecordset(Rs1)) Then
      'DO NOTHING: USE DEFAULT VALUES.
    Else
      '
      ' SET DEFAULT PROJECT DATA IN TEMPORARY VARIABLE.
      '
      
      
      '
      ' READ IN THE PROJECT DATA TO TEMPORARY VARIABLE.
      '
      Rs1.MoveFirst
      With NowProj
        Do Until Rs1.EOF
          Use_FieldIndex = CInt(Database_Get_Long(Rs1, "FieldIndex"))
          Use_FieldIndex2 = CInt(Database_Get_Long(Rs1, "FieldIndex2"))
          Use_FieldIndex3 = CInt(Database_Get_Long(Rs1, "FieldIndex3"))
          Use_FieldIndex4 = CInt(Database_Get_Long(Rs1, "FieldIndex3"))
          UB = UBound(.UserChemicals)
          If (Use_FieldIndex > UB) Then
            If (UB = 0) Then
              ReDim .UserChemicals(Use_FieldIndex)
            Else
              ReDim Preserve .UserChemicals(Use_FieldIndex)
            End If
            For i = UB + 1 To Use_FieldIndex
              ReDim .UserChemicals(i).PropertyData(0 To 0)
            Next i
          End If
          FieldName = Trim$(UCase$(Database_Get_String(Rs1, "FieldName")))
          SearchFor = "PD."
          If (Trim$(UCase$(Left$(FieldName, Len(SearchFor)))) = Trim$(UCase$(SearchFor))) Then
            UB = UBound(.UserChemicals(Use_FieldIndex).PropertyData)
            If (Use_FieldIndex2 > UB) Then
              If (UB = 0) Then
                ReDim .UserChemicals(Use_FieldIndex).PropertyData(1 To Use_FieldIndex2)
              Else
                ReDim Preserve .UserChemicals(Use_FieldIndex).PropertyData(1 To Use_FieldIndex2)
              End If
              For i = UB + 1 To Use_FieldIndex2
                ReDim .UserChemicals(Use_FieldIndex).PropertyData(i).TechniqueData(0 To 0)
              Next i
            End If
          End If
          SearchFor = "PD.TD."
          If (Trim$(UCase$(Left$(FieldName, Len(SearchFor)))) = Trim$(UCase$(SearchFor))) Then
            UB = UBound(.UserChemicals(Use_FieldIndex).PropertyData(Use_FieldIndex2).TechniqueData)
            If (Use_FieldIndex3 > UB) Then
              If (UB = 0) Then
                ReDim .UserChemicals(Use_FieldIndex). _
                    PropertyData(Use_FieldIndex2).TechniqueData(1 To Use_FieldIndex3)
              Else
                ReDim Preserve .UserChemicals(Use_FieldIndex). _
                    PropertyData(Use_FieldIndex2).TechniqueData(1 To Use_FieldIndex3)
              End If
              For i = UB + 1 To Use_FieldIndex3
                ReDim .UserChemicals(Use_FieldIndex). _
                    PropertyData(Use_FieldIndex2).TechniqueData(i).FofT_Coeffs(0 To 0)
              Next i
            End If
          End If
          SearchFor = "PD.TD.FofT_Coeffs"
          If (Trim$(UCase$(Left$(FieldName, Len(SearchFor)))) = Trim$(UCase$(SearchFor))) Then
            UB = UBound(.UserChemicals(Use_FieldIndex). _
                PropertyData(Use_FieldIndex2).TechniqueData(Use_FieldIndex3).FofT_Coeffs)
            If (Use_FieldIndex4 > UB) Then
              If (UB = 0) Then
                ReDim .UserChemicals(Use_FieldIndex). _
                    PropertyData(Use_FieldIndex2). _
                    TechniqueData(Use_FieldIndex3).FofT_Coeffs(1 To Use_FieldIndex4)
              Else
                ReDim Preserve .UserChemicals(Use_FieldIndex). _
                    PropertyData(Use_FieldIndex2). _
                    TechniqueData(Use_FieldIndex3).FofT_Coeffs(1 To Use_FieldIndex4)
              End If
            End If
          End If
          Select Case FieldName
'            Case Trim$(UCase$("PropertySheetOrder.Name")):
'              Call Database_LoadProperty(Rs1, .PropertySheetOrder(Use_FieldIndex).Name)
'            Case Trim$(UCase$("PropertySheetOrder.PropertyOrder.Property_Code")):
'              Call Database_LoadProperty(Rs1, .PropertySheetOrder(Use_FieldIndex). _
'                  PropertyOrder(Use_FieldIndex2).Property_Code)
'            Case Trim$(UCase$("PropertySheetOrder.PropertyOrder.Technique_Code")):
'              Call Database_LoadProperty(Rs1, .PropertySheetOrder(Use_FieldIndex). _
'                  PropertyOrder(Use_FieldIndex2).Technique_Code(Use_FieldIndex3))
          End Select
          Rs1.MoveNext
        Loop
        '''''
        ''''' TRANSFER PROJECT DATA TO MEMORY.
        '''''
        ''''NowProj = Prj
      End With
    End If
    Rs1.Close
  End If








'  '------ OUTPUT DATA TO TABLE "UserChemicals". ------------------------------------------------------------------------------------------------------
'  'START SAVE TO THIS TABLE.
'  Call Database_DeleteTableContents(Db1, "UserChemicals")
'  Set Rs1 = Db1.OpenRecordset("UserChemicals")
'  'MAIN BLOCK.
'  Prj = NowProj
'  For i = 1 To UBound(Prj.UserChemicals)
'    With Prj.UserChemicals(i)
'      '
'      ' MISCELLANEOUS CHEMICAL DATA.
'      '
'      Call Database_SavePropertyWithIndexes(Rs1, "User_Note", .User_Note, i)
'      '
'      ' BASIC CHEMICAL INFO.
'      '
'      Call Database_SavePropertyWithIndexes(Rs1, "Name", .Name, i)
'      Call Database_SavePropertyWithIndexes(Rs1, "CAS", .CAS, i)
'      Call Database_SavePropertyWithIndexes(Rs1, "SMILES", .SMILES, i)
'      Call Database_SavePropertyWithIndexes(Rs1, "Formula", .Formula, i)
'      Call Database_SavePropertyWithIndexes(Rs1, "Family", .Family, i)
'      Call Database_SavePropertyWithIndexes(Rs1, "Source", .Source, i)
'      '
'      ' CALCULATED RESULTS.
'      '
'      ''''PropertyData() As PropertyData_Type
'      For j = 1 To UBound(.PropertyData)
'        With .PropertyData(j)
'          '
'          ' MISCELLANEOUS PROPERTY DATA.
'          '
'          Call Database_SavePropertyWithIndexes(Rs1, "PD.User_Note", .User_Note, i, j)
'          '
'          ' MAIN DATA SET.
'          '
'          Call Database_SavePropertyWithIndexes(Rs1, "PD.UnitType", .UnitType, i, j)
'          Call Database_SavePropertyWithIndexes(Rs1, "PD.UnitBase", .UnitBase, i, j)
'          Call Database_SavePropertyWithIndexes(Rs1, "PD.UnitDisplayed", .UnitDisplayed, i, j)
'          Call Database_SavePropertyWithIndexes(Rs1, "PD.Property_Code", .Property_Code, i, j)
'          Call Database_SavePropertyWithIndexes(Rs1, "PD.Is_FofT", .Is_FofT, i, j)
'          ''''''Call Database_SavePropertyWithIndexes(Rs1, "PD.TechniqueData", .TechniqueData, i, j)
'          Call Database_SavePropertyWithIndexes(Rs1, "PD.IsAvail", .IsAvail, i, j)
'          Call Database_SavePropertyWithIndexes(Rs1, "PD.idx_Technique_Used", .idx_Technique_Used, i, j)
'          Call Database_SavePropertyWithIndexes(Rs1, "PD.Override_Technique_Code", .Override_Technique_Code, i, j)
'          For k = 1 To UBound(.TechniqueData)
'            With .TechniqueData(k)
'              '
'              ' Important note: The value actually reported by the program
'              ' on the main window is the first technique (ordered by
'              ' NowProj.UserHierarchy) that has .IsAvail=true.
'              '
'              Call Database_SavePropertyWithIndexes(Rs1, "PD.TD.Technique_Code", .Technique_Code, i, j, k)
'              Call Database_SavePropertyWithIndexes(Rs1, "PD.TD.IsAvail", .IsAvail, i, j, k)
'              Call Database_SavePropertyWithIndexes(Rs1, "PD.TD.Error_Code", .Error_Code, i, j, k)
'              Call Database_SavePropertyWithIndexes(Rs1, "PD.TD.Value", .value, i, j, k)
'              Call Database_SavePropertyWithIndexes(Rs1, "PD.TD.IsTagged", .IsTagged, i, j, k)
'              Call Database_SavePropertyWithIndexes(Rs1, "PD.TD.ReferenceText", .ReferenceText, i, j, k)
'              Call Database_SavePropertyWithIndexes(Rs1, "PD.TD.Text_When_Blank", .Text_When_Blank, i, j, k)
'              '
'              ' DIPPR RELATED VALUES.
'              '
'              Call Database_SavePropertyWithIndexes(Rs1, "PD.TD.DIPPR_REF", .DIPPR_REF, i, j, k)
'              Call Database_SavePropertyWithIndexes(Rs1, "PD.TD.DIPPR_REL", .DIPPR_REL, i, j, k)
'              Call Database_SavePropertyWithIndexes(Rs1, "PD.TD.DIPPR_R", .DIPPR_R, i, j, k)
'              Call Database_SavePropertyWithIndexes(Rs1, "PD.TD.DIPPR_Value", .DIPPR_Value, i, j, k)
'              Call Database_SavePropertyWithIndexes(Rs1, "PD.TD.DIPPR_Units", .DIPPR_Units, i, j, k)
'              Call Database_SavePropertyWithIndexes(Rs1, "PD.TD.DIPPR_Pressure", .DIPPR_Pressure, i, j, k)
'              Call Database_SavePropertyWithIndexes(Rs1, "PD.TD.DIPPR_DescMethod", .DIPPR_DescMethod, i, j, k)
'              Call Database_SavePropertyWithIndexes(Rs1, "PD.TD.DIPPR_Comment", .DIPPR_Comment, i, j, k)
'              Call Database_SavePropertyWithIndexes(Rs1, "PD.TD.DIPPR_ArticleNumber", .DIPPR_ArticleNumber, i, j, k)
'              '
'              ' FUNCTION OF TEMPERATURE VALUES.
'              '
'              Call Database_SavePropertyWithIndexes(Rs1, "PD.TD.FofT_EqForm", .FofT_EqForm, i, j, k)
'              For m = 1 To UBound(.FofT_Coeffs)
'                Call Database_SavePropertyWithIndexes(Rs1, "PD.TD.FofT_Coeffs", .FofT_Coeffs(m), i, j, k, m)
'              Next m
'              Call Database_SavePropertyWithIndexes(Rs1, "PD.TD.FofT_Units_F", .FofT_Units_F, i, j, k)
'              Call Database_SavePropertyWithIndexes(Rs1, "PD.TD.FofT_Units_T", .FofT_Units_T, i, j, k)
'              Call Database_SavePropertyWithIndexes(Rs1, "PD.TD.FofT_Minimum_T", .FofT_Minimum_T, i, j, k)
'              Call Database_SavePropertyWithIndexes(Rs1, "PD.TD.FofT_Maximum_T", .FofT_Maximum_T, i, j, k)
'            End With
'          Next k
'        End With
'      Next j
'    End With
'  Next i





  'CLOSE THE DATABASE FILE.
  Db1.Close

  'RETURN A "SUCCESS" MESSAGE TO CALLER.
  File_Open_Latest_v1_00 = True

End Function


'RETURNS:
'         TRUE = SUCCEEDED IN SAVING.
'         FALSE = FAILED IN SAVING.
Function File_Save_Latest_v1_00( _
    fn_This As String) As Boolean
Dim Ws1 As Workspace
Dim Db1 As Database
Dim Rs1 As Recordset
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim m As Integer
Dim IsInvalidFormat As Boolean
Dim NeedToCreateNewDatabase As Boolean
Dim Prj As Project_Type

'>>>>>>>>>>>>>>>>>>>> *TODO* UPDATE THIS ENTIRE FUNCTION <<<<<<<<<<<<<<<<<<<<<<<<<

  '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
  '>>>>>>>>>>>>>>>>>>>>>>>>>>>  SAVE TO MAIN DATABASE  <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
  '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
  
  'IF FILE DOES NOT EXIST, CREATE IT.
  'FOR EACH TABLE, IF IT EXISTS, DELETE IT.
  If (Not FileExists(fn_This)) Then
    'CREATE NEW DATABASE.
    NeedToCreateNewDatabase = True
  Else
    'DETERMINE WHETHER OLD FILE IS AN INVALID VERSION (i.e. A NON-MDB FILE).
    IsInvalidFormat = True
    On Error Resume Next
    Set Db1 = OpenDatabase(fn_This)
    If (Err.Number = 0) Then
      IsInvalidFormat = False
      Db1.Close
    End If
    On Error GoTo 0
    If (IsInvalidFormat) Then
      'DELETE OLD FILE, CREATE NEW DATABASE (SEE BELOW).
      Kill fn_This
      NeedToCreateNewDatabase = True
    Else
      'OPEN DATABASE NORMALLY.
      Set Db1 = OpenDatabase(fn_This)
    End If
  End If
  If (NeedToCreateNewDatabase) Then
    Set Db1 = CreateDatabase(fn_This, dbLangGeneral)
  End If
  'CREATE NEW TABLES WITHIN DATABASE, IF NECESSARY.
  Call Database_CreateMFBTable_IfNoExist_MultipleIndexes(Db1, "Version")
  Call Database_CreateMFBTable_IfNoExist_MultipleIndexes(Db1, "Main")
  Call Database_CreateMFBTable_IfNoExist_MultipleIndexes(Db1, "UserHierarchy")
  Call Database_CreateMFBTable_IfNoExist_MultipleIndexes(Db1, "UserChemicals")
  
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
  With Prj
    '
    ' MISCELLANEOUS FILE DATA.
    '
    Call Database_SaveProperty(Rs1, "File_Note", .File_Note)
    '
    ' HIERARCHY RELATED DATA.
    '
    ''''UserHierarchy As UserHierarchy_Type
    '
    ' MAIN DATA SET.
    '
    Call Database_SaveProperty(Rs1, "Op_T", .Op_T)
    Call Database_SaveProperty(Rs1, "Op_P", .Op_P)
    Call Database_SaveProperty(Rs1, "Op_T_UnitDisplayed", .Op_T_UnitDisplayed)
    Call Database_SaveProperty(Rs1, "Op_P_UnitDisplayed", .Op_P_UnitDisplayed)
    ''''UserChemicals() As UserChemical_Type
  End With

  '------ OUTPUT DATA TO TABLE "UserHierarchy". ------------------------------------------------------------------------------------------------------
  'START SAVE TO THIS TABLE.
  Call Database_DeleteTableContents(Db1, "UserHierarchy")
  Set Rs1 = Db1.OpenRecordset("UserHierarchy")
  'MAIN BLOCK.
  Prj = NowProj
  With Prj.UserHierarchy
    For i = 1 To UBound(.PropertySheetOrder)
      Call Database_SavePropertyWithIndexes(Rs1, "PropertySheetOrder.Name", .PropertySheetOrder(i).Name, i)
      For j = 1 To UBound(.PropertySheetOrder(i).PropertyOrder)
        Call Database_SavePropertyWithIndexes(Rs1, "PropertySheetOrder.PropertyOrder.Property_Code", _
            .PropertySheetOrder(i).PropertyOrder(j).Property_Code, i, j)
        For k = 1 To UBound(.PropertySheetOrder(i).PropertyOrder(j).Technique_Code)
          Call Database_SavePropertyWithIndexes(Rs1, "PropertySheetOrder.PropertyOrder.Technique_Code", _
              .PropertySheetOrder(i).PropertyOrder(j).Technique_Code(k), i, j, k)
        Next k
      Next j
    Next i
  End With
  
  '------ OUTPUT DATA TO TABLE "UserChemicals". ------------------------------------------------------------------------------------------------------
  'START SAVE TO THIS TABLE.
  Call Database_DeleteTableContents(Db1, "UserChemicals")
  Set Rs1 = Db1.OpenRecordset("UserChemicals")
  'MAIN BLOCK.
  Prj = NowProj
  For i = 1 To UBound(Prj.UserChemicals)
    With Prj.UserChemicals(i)
      '
      ' MISCELLANEOUS CHEMICAL DATA.
      '
      Call Database_SavePropertyWithIndexes(Rs1, "User_Note", .User_Note, i)
      '
      ' BASIC CHEMICAL INFO.
      '
      Call Database_SavePropertyWithIndexes(Rs1, "Name", .Name, i)
      Call Database_SavePropertyWithIndexes(Rs1, "CAS", .CAS, i)
      Call Database_SavePropertyWithIndexes(Rs1, "SMILES", .SMILES, i)
      Call Database_SavePropertyWithIndexes(Rs1, "Formula", .Formula, i)
      Call Database_SavePropertyWithIndexes(Rs1, "Family", .Family, i)
      Call Database_SavePropertyWithIndexes(Rs1, "Source", .Source, i)
      '
      ' CALCULATED RESULTS.
      '
      ''''PropertyData() As PropertyData_Type
      For j = 1 To UBound(.PropertyData)
        With .PropertyData(j)
          '
          ' MISCELLANEOUS PROPERTY DATA.
          '
          Call Database_SavePropertyWithIndexes(Rs1, "PD.User_Note", .User_Note, i, j)
          '
          ' MAIN DATA SET.
          '
          Call Database_SavePropertyWithIndexes(Rs1, "PD.UnitType", .UnitType, i, j)
          Call Database_SavePropertyWithIndexes(Rs1, "PD.UnitBase", .UnitBase, i, j)
          Call Database_SavePropertyWithIndexes(Rs1, "PD.UnitDisplayed", .UnitDisplayed, i, j)
          Call Database_SavePropertyWithIndexes(Rs1, "PD.Property_Code", .Property_Code, i, j)
          Call Database_SavePropertyWithIndexes(Rs1, "PD.Is_FofT", .Is_FofT, i, j)
          ''''''Call Database_SavePropertyWithIndexes(Rs1, "PD.TechniqueData", .TechniqueData, i, j)
          Call Database_SavePropertyWithIndexes(Rs1, "PD.IsAvail", .IsAvail, i, j)
          Call Database_SavePropertyWithIndexes(Rs1, "PD.idx_Technique_Used", .idx_Technique_Used, i, j)
          Call Database_SavePropertyWithIndexes(Rs1, "PD.Override_Technique_Code", .Override_Technique_Code, i, j)
          For k = 1 To UBound(.TechniqueData)
            With .TechniqueData(k)
              '
              ' Important note: The value actually reported by the program
              ' on the main window is the first technique (ordered by
              ' NowProj.UserHierarchy) that has .IsAvail=true.
              '
              Call Database_SavePropertyWithIndexes(Rs1, "PD.TD.Technique_Code", .Technique_Code, i, j, k)
              Call Database_SavePropertyWithIndexes(Rs1, "PD.TD.IsAvail", .IsAvail, i, j, k)
              Call Database_SavePropertyWithIndexes(Rs1, "PD.TD.Error_Code", .Error_Code, i, j, k)
              Call Database_SavePropertyWithIndexes(Rs1, "PD.TD.Value", .value, i, j, k)
              Call Database_SavePropertyWithIndexes(Rs1, "PD.TD.IsTagged", .IsTagged, i, j, k)
              Call Database_SavePropertyWithIndexes(Rs1, "PD.TD.ReferenceText", .ReferenceText, i, j, k)
              Call Database_SavePropertyWithIndexes(Rs1, "PD.TD.Text_When_Blank", .Text_When_Blank, i, j, k)
              '
              ' DIPPR RELATED VALUES.
              '
              Call Database_SavePropertyWithIndexes(Rs1, "PD.TD.DIPPR_REF", .DIPPR_REF, i, j, k)
              Call Database_SavePropertyWithIndexes(Rs1, "PD.TD.DIPPR_REL", .DIPPR_REL, i, j, k)
              Call Database_SavePropertyWithIndexes(Rs1, "PD.TD.DIPPR_R", .DIPPR_R, i, j, k)
              Call Database_SavePropertyWithIndexes(Rs1, "PD.TD.DIPPR_Value", .DIPPR_Value, i, j, k)
              Call Database_SavePropertyWithIndexes(Rs1, "PD.TD.DIPPR_Units", .DIPPR_Units, i, j, k)
              Call Database_SavePropertyWithIndexes(Rs1, "PD.TD.DIPPR_Pressure", .DIPPR_Pressure, i, j, k)
              Call Database_SavePropertyWithIndexes(Rs1, "PD.TD.DIPPR_DescMethod", .DIPPR_DescMethod, i, j, k)
              Call Database_SavePropertyWithIndexes(Rs1, "PD.TD.DIPPR_Comment", .DIPPR_Comment, i, j, k)
              Call Database_SavePropertyWithIndexes(Rs1, "PD.TD.DIPPR_ArticleNumber", .DIPPR_ArticleNumber, i, j, k)
              '
              ' FUNCTION OF TEMPERATURE VALUES.
              '
              Call Database_SavePropertyWithIndexes(Rs1, "PD.TD.FofT_EqForm", .FofT_EqForm, i, j, k)
              For m = 1 To UBound(.FofT_Coeffs)
                Call Database_SavePropertyWithIndexes(Rs1, "PD.TD.FofT_Coeffs", .FofT_Coeffs(m), i, j, k, m)
              Next m
              Call Database_SavePropertyWithIndexes(Rs1, "PD.TD.FofT_Units_F", .FofT_Units_F, i, j, k)
              Call Database_SavePropertyWithIndexes(Rs1, "PD.TD.FofT_Units_T", .FofT_Units_T, i, j, k)
              Call Database_SavePropertyWithIndexes(Rs1, "PD.TD.FofT_Minimum_T", .FofT_Minimum_T, i, j, k)
              Call Database_SavePropertyWithIndexes(Rs1, "PD.TD.FofT_Maximum_T", .FofT_Maximum_T, i, j, k)
            End With
          Next k
        End With
      Next j
    End With
  Next i

  '
  ' END SAVE TO THIS TABLE.
  '
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
  ''''Call ProjectFile_Read(f, InLine, Dummy1)
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
      'If (Use_memoValue) Then
        LoadedData = CStr(Database_Get_String(Rs1, "memoValue"))
      'Else
      '  LoadedData = CStr(Database_Get_String(Rs1, "strValue"))
      'End If
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
  If (Left$(in_Use_FieldName, 1) = "*") Then
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
      'If (Use_memoValue) Then
        Rs1("memoValue") = CStr(SavedData)
      'Else
      '  Rs1("strValue") = CStr(SavedData)
      'End If
    Case vbDouble, vbSingle:
      Rs1("dblValue") = CDbl(SavedData)
  End Select
  Rs1.Update
End Sub
Sub Database_SavePropertyWithIndexes( _
    Rs1 As Recordset, _
    in_Use_FieldName As String, _
    SavedData As Variant, _
    Use_FieldIndex As Integer, _
    Optional Use_FieldIndex2 As Integer = -1, _
    Optional Use_FieldIndex3 As Integer = -1, _
    Optional Use_FieldIndex4 As Integer = -1)
Dim Use_memoValue As Boolean
Dim Use_FieldName As String
  'NOTE: IF THE FIRST CHARACTER IS AN ASTERISK ("*"), THEN
  'THE FIELD TYPE USED IS THE MEMO TYPE (memoValue).
  Use_memoValue = False
  If (Left$(in_Use_FieldName, 1) = "*") Then
    Use_FieldName = Right$(in_Use_FieldName, Len(in_Use_FieldName) - 1)
    Use_memoValue = True
  Else
    Use_FieldName = in_Use_FieldName
  End If
  Rs1.AddNew
  Rs1("FieldName") = Use_FieldName
  Rs1("FieldIndex") = Use_FieldIndex
  If (Use_FieldIndex2 <> -1) Then
    Rs1("FieldIndex2") = Use_FieldIndex2
  End If
  If (Use_FieldIndex3 <> -1) Then
    Rs1("FieldIndex3") = Use_FieldIndex3
  End If
  Select Case VarType(SavedData)
    Case vbBoolean, vbByte, vbInteger, vbLong:
      Rs1("lngValue") = CLng(SavedData)
    Case vbString, vbDate:
      'If (Use_memoValue) Then
        Rs1("memoValue") = CStr(SavedData)
      'Else
      '  Rs1("strValue") = CStr(SavedData)
      'End If
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
    Set Ff = Td1.CreateField("FieldIndex3", dbLong):
    Td1.Fields.Append Ff
    Set Ff = Td1.CreateField("FieldIndex4", dbLong):
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
Sub Database_CreateMFBTable_IfNoExist_MultipleIndexes( _
    Db1 As Database, _
    Use_TableName As String)
  If (Database_IsTableExist(Db1, Use_TableName) = False) Then
    Call Database_CreateMFBTable(Db1, Use_TableName, True, True)
  End If
End Sub




