Attribute VB_Name = "StructsDo"
Option Explicit



Const StructsDo_declarations_end = True


Sub Project_SetDefaults(Prj As Project_Type)
''''  '
''''  ' TEMPORARY, SOON-TO-BE-DELETED DATA.
''''  '
''''  Prj.length = 1#
''''  Prj.Diameter = 1#
''''  Prj.Mass = 1#
''''  Prj.FlowRate = 1#
  '
  ' MISCELLANEOUS FILE DATA.
  '
  Prj.File_Note = ""
  '
  ' HIERARCHY RELATED DATA.
  '
  Call Project_UserHierarchy_SetDefaults(Prj)
  '
  ' MAIN DATA SET.
  '
  Prj.Op_T = 298.15         'K
  Prj.Op_P = 101325#        'Pa
  Prj.Op_T_UnitDisplayed = "C"
  Prj.Op_P_UnitDisplayed = "Pa"
  ReDim Prj.UserChemicals(0 To 0)
End Sub


Sub UserChemical_SetDefaults( _
    CH As UserChemical_Type, _
    NewName As String)
  With CH
    '
    ' MISCELLANEOUS CHEMICAL DATA.
    '
    .User_Note = ""
    '
    ' BASIC CHEMICAL INFO.
    '
    .Name = NewName
    .CAS = ""
    .SMILES = ""
    .Formula = ""
    .Family = ""
    .Source = ""
    '
    ' CALCULATED RESULTS.
    '
    ReDim .PropertyData(0 To 0)
  End With
End Sub
Function UserChemical_GetNewNameDefault( _
    ) As String
Dim RecName_New As String
Dim i As Integer
  'DETERMINE DEFAULT NAME OF NEW RECORD.
  i = 1
  Do While (1 = 1)
    RecName_New = "New Chemical " & Trim$(Str$(i))
    If (Not UserChemical_IsKeyExist(RecName_New)) Then Exit Do
    i = i + 1
  Loop
  'RETURN THIS DEFAULT NAME.
  UserChemical_GetNewNameDefault = RecName_New
End Function
Function UserChemical_GetIndex( _
    RecName As String) As Integer
Dim Found As Integer
Dim i As Integer
  Found = False
  For i = 1 To UBound(NowProj.UserChemicals)
    If (Trim$(UCase$(NowProj.UserChemicals(i).Name)) = _
        Trim$(UCase$(RecName))) Then
      Found = True
      Exit For
    End If
  Next i
  If (Found) Then
    UserChemical_GetIndex = i
  Else
    UserChemical_GetIndex = 0
  End If
End Function
Function UserChemical_IsKeyExist( _
    RecName As String) As Boolean
Dim RetVal As Integer
  RetVal = UserChemical_GetIndex(RecName)
  If (RetVal = 0) Then
    UserChemical_IsKeyExist = False
  Else
    UserChemical_IsKeyExist = True
  End If
End Function


Function PropertySheetOrder_GetIndex( _
    in_Name As String) As Integer
Dim Found As Integer
Dim i As Integer
  Found = False
  For i = 1 To UBound(NowProj.UserHierarchy.PropertySheetOrder)
    If (Trim$(UCase$(NowProj.UserHierarchy.PropertySheetOrder(i).Name)) = _
        Trim$(UCase$(in_Name))) Then
      Found = True
      Exit For
    End If
  Next i
  If (Found) Then
    PropertySheetOrder_GetIndex = i
  Else
    PropertySheetOrder_GetIndex = -1
  End If
End Function
Function PropertySheetOrder_IsKeyExist( _
    in_Name As String) As Boolean
Dim RetVal As Integer
  RetVal = PropertySheetOrder_GetIndex(in_Name)
  If (RetVal = -1) Then
    PropertySheetOrder_IsKeyExist = False
  Else
    PropertySheetOrder_IsKeyExist = True
  End If
End Function
Function PropertySheetOrder_GetNewNameDefault( _
    ) As String
Dim RecName_New As String
Dim i As Integer
  'DETERMINE DEFAULT NAME OF NEW RECORD.
  i = 1
  Do While (1 = 1)
    RecName_New = "New Property Sheet " & Trim$(Str$(i))
    If (Not PropertySheetOrder_IsKeyExist(RecName_New)) Then Exit Do
    i = i + 1
  Loop
  'RETURN THIS DEFAULT NAME.
  PropertySheetOrder_GetNewNameDefault = RecName_New
End Function


Function PropertyOrder_Technique_Code_GetIndexes( _
    in_Property_Code As Long, _
    in_Technique_Code As Long, _
    out_idx_PropertySheetOrder() As Integer, _
    out_idx_PropertyOrder() As Integer, _
    out_idx_Technique_Code() As Integer, _
    out_Size As Integer) As Boolean
Dim UB_Hier1 As Integer
Dim UB_Hier2 As Integer
Dim UB_Hier3 As Integer
Dim i1 As Integer
Dim i2 As Integer
Dim i3 As Integer
  out_Size = 0
  ReDim out_idx_PropertySheetOrder(0 To 0)
  ReDim out_idx_PropertyOrder(0 To 0)
  UB_Hier1 = UBound(NowProj.UserHierarchy. _
      PropertySheetOrder)
  For i1 = 1 To UB_Hier1
    UB_Hier2 = UBound(NowProj.UserHierarchy. _
        PropertySheetOrder(i1).PropertyOrder)
    For i2 = 1 To UB_Hier2
      If (in_Property_Code = NowProj.UserHierarchy. _
          PropertySheetOrder(i1).PropertyOrder(i2).Property_Code) Then
        UB_Hier3 = UBound(NowProj.UserHierarchy. _
            PropertySheetOrder(i1).PropertyOrder(i2).Technique_Code)
        For i3 = 1 To UB_Hier3
          If (in_Technique_Code = NowProj.UserHierarchy. _
              PropertySheetOrder(i1).PropertyOrder(i2).Technique_Code(i3)) Then
            out_Size = out_Size + 1
            If (out_Size = 1) Then
              ReDim out_idx_PropertySheetOrder(1 To out_Size)
              ReDim out_idx_PropertyOrder(1 To out_Size)
              ReDim out_idx_Technique_Code(1 To out_Size)
            Else
              ReDim Preserve out_idx_PropertySheetOrder(1 To out_Size)
              ReDim Preserve out_idx_PropertyOrder(1 To out_Size)
              ReDim Preserve out_idx_Technique_Code(1 To out_Size)
            End If
            out_idx_PropertySheetOrder(out_Size) = i1
            out_idx_PropertyOrder(out_Size) = i2
            out_idx_Technique_Code(out_Size) = i3
          End If
        Next i3
      End If
    Next i2
  Next i1
  PropertyOrder_Technique_Code_GetIndexes = True
End Function



Function PropertyOrder_Update_Technique_Code( _
    in_Technique_Code() As Long, _
    in_PropCode As Long) _
    As Boolean
On Error GoTo err_ThisFunc
Dim out_idx_PropertySheetOrder() As Integer
Dim out_idx_PropertyOrder() As Integer
Dim out_Size As Integer
Dim i As Integer
Dim idx1 As Integer
Dim idx2 As Integer
  If (False = PropertyOrder_Property_Code_GetIndexes( _
      in_PropCode, _
      out_idx_PropertySheetOrder(), _
      out_idx_PropertyOrder(), _
      out_Size)) Then
    GoTo exit_err_ThisFunc
  End If
  For i = 1 To out_Size
    idx1 = out_idx_PropertySheetOrder(i)
    idx2 = out_idx_PropertyOrder(i)
    NowProj.UserHierarchy. _
        PropertySheetOrder(idx1). _
        PropertyOrder(idx2). _
        Technique_Code = in_Technique_Code
  Next i
exit_normally_ThisFunc:
  PropertyOrder_Update_Technique_Code = True
  Exit Function
exit_err_ThisFunc:
  PropertyOrder_Update_Technique_Code = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("PropertyOrder_Update_Technique_Code")
  Resume exit_err_ThisFunc
End Function
Function PropertyOrder_Get_List_of_Unique_Property_Codes( _
    out_List_PropCodes() As Long, _
    out_List_idxPropSheet_First_Occurrences() As Integer, _
    out_List_idxPropOrd_First_Occurrences() As Integer) _
    As Boolean
On Error GoTo err_ThisFunc
Dim UB_Output As Integer
Dim UB1 As Integer
Dim UB2 As Integer
Dim i As Integer
Dim j As Integer
Dim This_PropCode As Long
Dim out_idx_Elem As Integer
  ReDim out_List_PropCodes(0 To 0)
  ReDim out_List_idxPropSheet_First_Occurrences(0 To 0)
  ReDim out_List_idxPropOrd_First_Occurrences(0 To 0)
  UB_Output = 0
  UB1 = UBound(NowProj.UserHierarchy.PropertySheetOrder)
  For i = 1 To UB1
    UB2 = UBound(NowProj.UserHierarchy.PropertySheetOrder(i).PropertyOrder)
    For j = 1 To UB2
      This_PropCode = NowProj.UserHierarchy. _
          PropertySheetOrder(i).PropertyOrder(j).Property_Code
      If (False = sc_ElemFind( _
          out_List_PropCodes, _
          This_PropCode, _
          out_idx_Elem)) Then
        UB_Output = UB_Output + 1
        If (UB_Output = 1) Then
          ReDim out_List_PropCodes(1 To 1)
          ReDim out_List_idxPropSheet_First_Occurrences(1 To 1)
          ReDim out_List_idxPropOrd_First_Occurrences(1 To 1)
        Else
          ReDim Preserve out_List_PropCodes(1 To UB_Output)
          ReDim Preserve out_List_idxPropSheet_First_Occurrences(1 To UB_Output)
          ReDim Preserve out_List_idxPropOrd_First_Occurrences(1 To UB_Output)
        End If
        out_List_PropCodes(UB_Output) = This_PropCode
        out_List_idxPropSheet_First_Occurrences(UB_Output) = i
        out_List_idxPropOrd_First_Occurrences(UB_Output) = j
      End If
    Next j
  Next i
  '
  ' EXIT OUTTA HERE.
  '
exit_normally_ThisFunc:
  PropertyOrder_Get_List_of_Unique_Property_Codes = True
  Exit Function
exit_err_ThisFunc:
  PropertyOrder_Get_List_of_Unique_Property_Codes = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("PropertyOrder_Get_List_of_Unique_Property_Codes")
  Resume exit_err_ThisFunc
End Function
Function PropertyOrder_Property_Code_GetIndexes( _
    in_Property_Code As Long, _
    out_idx_PropertySheetOrder() As Integer, _
    out_idx_PropertyOrder() As Integer, _
    out_Size As Integer) As Boolean
Dim UB_Hier1 As Integer
Dim UB_Hier2 As Integer
Dim i1 As Integer
Dim i2 As Integer
  out_Size = 0
  ReDim out_idx_PropertySheetOrder(0 To 0)
  ReDim out_idx_PropertyOrder(0 To 0)
  UB_Hier1 = UBound(NowProj.UserHierarchy. _
      PropertySheetOrder)
  For i1 = 1 To UB_Hier1
    UB_Hier2 = UBound(NowProj.UserHierarchy. _
        PropertySheetOrder(i1).PropertyOrder)
    For i2 = 1 To UB_Hier2
      If (in_Property_Code = NowProj.UserHierarchy. _
          PropertySheetOrder(i1).PropertyOrder(i2).Property_Code) Then
        out_Size = out_Size + 1
        If (out_Size = 1) Then
          ReDim out_idx_PropertySheetOrder(1 To out_Size)
          ReDim out_idx_PropertyOrder(1 To out_Size)
        Else
          ReDim Preserve out_idx_PropertySheetOrder(1 To out_Size)
          ReDim Preserve out_idx_PropertyOrder(1 To out_Size)
        End If
        out_idx_PropertySheetOrder(out_Size) = i1
        out_idx_PropertyOrder(out_Size) = i2
      End If
    Next i2
  Next i1
  PropertyOrder_Property_Code_GetIndexes = True
End Function


Function TechValue_Put( _
    in_idx_Chem As Integer, _
    in_Property_Code As Long, _
    in_Technique_Code As Long, _
    in_TechValue As Double, _
    in_TechValue_IsAvail As Boolean) _
    As Boolean
On Error GoTo err_ThisFunc
Dim out_idx_PropertyData As Integer
Dim out_idx_TechniqueData As Integer
Dim Use_TechValue_IsAvail As Boolean
  Use_TechValue_IsAvail = in_TechValue_IsAvail
  If (in_TechValue = 0#) Then
    Use_TechValue_IsAvail = False
  End If
  If (False = TechniqueData_GetIndex( _
      in_idx_Chem, _
      in_Property_Code, _
      in_Technique_Code, _
      out_idx_PropertyData, _
      out_idx_TechniqueData)) Then
    GoTo exit_err_ThisFunc
  End If
  '
  ' SET THE VALUE AND WHETHER IT IS AVAILABLE.
  '
  With NowProj.UserChemicals(in_idx_Chem). _
      PropertyData(out_idx_PropertyData). _
      TechniqueData(out_idx_TechniqueData)
    .IsAvail = Use_TechValue_IsAvail
    .Error_Code = ""
    .value = in_TechValue
    .ReferenceText = ""     'IMPORTANT TO SET TO NULL HERE;
              'THE ROUTINE BELOW ONLY _APPENDS_ TO EXISTING REFERENCE TEXT.
  End With
  '
  ' SET THE REFERENCE TEXT.
  '
  If (False = Calc_Mod_GetRefText( _
      NowProj.UserChemicals(in_idx_Chem). _
      PropertyData(out_idx_PropertyData). _
      TechniqueData(out_idx_TechniqueData))) Then
    GoTo exit_err_ThisFunc
  End If
exit_normally_ThisFunc:
  TechValue_Put = True
  Exit Function
exit_err_ThisFunc:
  TechValue_Put = False
  Exit Function
err_ThisFunc:
  ''''Call Show_Trapped_Error("TechValue_Put")
  Resume exit_err_ThisFunc
End Function
Function TechValue_Get( _
    in_idx_Chem As Integer, _
    in_Property_Code As Long, _
    out_TechValue As Double) _
    As Boolean
On Error GoTo err_ThisFunc
Dim This_idx_Technique_Used As Integer
Dim This_IsAvail As Boolean
Dim This_Value As Double
Dim idx_PropertyData As Integer
  idx_PropertyData = PropertyData_GetIndex( _
      in_idx_Chem, _
      in_Property_Code)
  On Error GoTo 0
  If (-1 = idx_PropertyData) Then
    GoTo exit_err_ThisFunc
  End If
  With NowProj.UserChemicals(in_idx_Chem). _
      PropertyData(idx_PropertyData)
    This_idx_Technique_Used = .idx_Technique_Used
    If (This_idx_Technique_Used < LBound(.TechniqueData)) Or _
        (This_idx_Technique_Used > UBound(.TechniqueData)) Then
      GoTo exit_err_ThisFunc
    End If
    This_IsAvail = .TechniqueData(This_idx_Technique_Used).IsAvail
    If (This_IsAvail = True) Then
      This_Value = .TechniqueData(This_idx_Technique_Used).value
      ''''This_Text_When_Blank = .TechniqueData(This_idx_Technique_Used).Text_When_Blank
    Else
      GoTo exit_err_ThisFunc
    End If
    ''''This_UnitType = .UnitType
    ''''This_UnitBase = .UnitBase
    ''''This_UnitDisplayed = .UnitDisplayed
  End With
  out_TechValue = This_Value
exit_normally_ThisFunc:
  TechValue_Get = True
  Exit Function
exit_err_ThisFunc:
  TechValue_Get = False
  Exit Function
err_ThisFunc:
  ''''Call Show_Trapped_Error("TechValue_Get")
  GoTo exit_err_ThisFunc
End Function
Function TechValue_IsAvail( _
    in_idx_Chem As Integer, _
    in_Property_Code As Long, _
    out_TechValue_IsAvail As Double) _
    As Boolean
On Error GoTo err_ThisFunc
Dim This_idx_Technique_Used As Integer
Dim This_IsAvail As Boolean
Dim This_Value As Double
Dim idx_PropertyData As Integer
  idx_PropertyData = PropertyData_GetIndex( _
      in_idx_Chem, _
      in_Property_Code)
  On Error GoTo 0
  If (-1 = idx_PropertyData) Then
    GoTo exit_err_ThisFunc
  End If
  With NowProj.UserChemicals(in_idx_Chem). _
      PropertyData(idx_PropertyData)
    This_idx_Technique_Used = .idx_Technique_Used
    If (This_idx_Technique_Used < LBound(.TechniqueData)) Or _
        (This_idx_Technique_Used > UBound(.TechniqueData)) Then
      GoTo exit_err_ThisFunc
    End If
    This_IsAvail = .TechniqueData(This_idx_Technique_Used).IsAvail
    out_TechValue_IsAvail = This_IsAvail
    'If (This_IsAvail = True) Then
    '  This_Value = .TechniqueData(This_idx_Technique_Used).Value
    '  ''''This_Text_When_Blank = .TechniqueData(This_idx_Technique_Used).Text_When_Blank
    'Else
    '  GoTo exit_err_ThisFunc
    'End If
    ''''This_UnitType = .UnitType
    ''''This_UnitBase = .UnitBase
    ''''This_UnitDisplayed = .UnitDisplayed
  End With
  ''''out_TechValue = This_Value
exit_normally_ThisFunc:
  TechValue_IsAvail = True
  Exit Function
exit_err_ThisFunc:
  TechValue_IsAvail = False
  Exit Function
err_ThisFunc:
  ''''Call Show_Trapped_Error("TechValue_IsAvail")
  GoTo exit_err_ThisFunc
End Function


Function TechniqueData_GetIndex( _
    in_idx_Chem As Integer, _
    in_Property_Code As Long, _
    in_Technique_Code As Long, _
    out_idx_PropertyData As Integer, _
    out_idx_TechniqueData As Integer) As Boolean
Dim Found As Integer
Dim i As Integer
Dim UB As Integer
  Found = False
  out_idx_PropertyData = PropertyData_GetIndex(in_idx_Chem, in_Property_Code)
  If (out_idx_PropertyData = -1) Then
    out_idx_TechniqueData = -1
    TechniqueData_GetIndex = False
    Exit Function
  End If
  UB = UBound(NowProj.UserChemicals(in_idx_Chem). _
      PropertyData(out_idx_PropertyData).TechniqueData)
  For i = 1 To UB
    If (NowProj.UserChemicals(in_idx_Chem). _
        PropertyData(out_idx_PropertyData). _
        TechniqueData(i).Technique_Code = _
        in_Technique_Code) Then
      Found = True
      Exit For
    End If
  Next i
  If (Found) Then
    out_idx_TechniqueData = i
    TechniqueData_GetIndex = True
  Else
    out_idx_TechniqueData = -1
    TechniqueData_GetIndex = False
  End If
End Function
Function TechniqueData_SetDefaults( _
    TD As TechniqueData_Type) _
    As Boolean
On Error GoTo err_ThisFunc
  With TD
    '
    ' Important note: The value actually reported by the program
    ' on the main window is the first technique (ordered by
    ' NowProj.UserHierarchy) that has .IsAvail=true.
    '
    .Technique_Code = -1
    .IsAvail = False
    .Error_Code = TECH_ERRORCODE_NEVER_INITED
    .value = 0#
    .IsTagged = False
    .ReferenceText = ""
    .Text_When_Blank = ""
    '
    ' DIPPR RELATED VALUES.
    '
    .DIPPR_REF = ""
    .DIPPR_REL = ""
    .DIPPR_R = -1
    .DIPPR_Value = 0#
    .DIPPR_Units = ""
    .DIPPR_Pressure = ""
    .DIPPR_DescMethod = ""
    .DIPPR_Comment = ""
    .DIPPR_ArticleNumber = -1
    '
    ' FUNCTION OF TEMPERATURE VALUES.
    '
    .FofT_EqForm = 0
    ReDim .FofT_Coeffs(1 To 5)
    .FofT_Units_F = ""
    .FofT_Units_T = ""
    .FofT_Minimum_T = 0#
    .FofT_Maximum_T = 0#
  End With
exit_normally_ThisFunc:
  TechniqueData_SetDefaults = True
  Exit Function
exit_err_ThisFunc:
  TechniqueData_SetDefaults = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("TechniqueData_SetDefaults")
  Resume exit_err_ThisFunc
End Function


Function PropertyData_GetIndex( _
    in_idx_Chem As Integer, _
    in_Property_Code As Long) As Integer
Dim Found As Integer
Dim i As Integer
  Found = False
  For i = 1 To UBound(NowProj. _
      UserChemicals(in_idx_Chem).PropertyData)
    If (NowProj.UserChemicals(in_idx_Chem). _
        PropertyData(i).Property_Code = _
        in_Property_Code) Then
      Found = True
      Exit For
    End If
  Next i
  If (Found) Then
    PropertyData_GetIndex = i
  Else
    PropertyData_GetIndex = -1
  End If
End Function
Function PropertyData_IsKeyExist( _
    in_idx_Chem As Integer, _
    in_Property_Code As Long) As Boolean
Dim RetVal As Integer
  RetVal = PropertyData_GetIndex(in_idx_Chem, in_Property_Code)
  If (RetVal = -1) Then
    PropertyData_IsKeyExist = False
  Else
    PropertyData_IsKeyExist = True
  End If
End Function
Function PropertyData_SetDefaults( _
    PD As PropertyData_Type) _
    As Boolean
On Error GoTo err_ThisFunc
  With PD
    '
    ' MISCELLANEOUS PROPERTY DATA.
    '
    .User_Note = ""
    '
    ' MAIN DATA SET.
    '
    .UnitType = ""
    .UnitBase = ""
    .UnitDisplayed = ""
    .Property_Code = -1
    .Is_FofT = False
    ReDim .TechniqueData(0 To 0)
    .IsAvail = False
    .idx_Technique_Used = -1
    .Override_Technique_Code = -1
  End With
exit_normally_ThisFunc:
  PropertyData_SetDefaults = True
  Exit Function
exit_err_ThisFunc:
  PropertyData_SetDefaults = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("PropertyData_SetDefaults")
  Resume exit_err_ThisFunc
End Function
Function PropertyData_InitializeAll_AllChemicals( _
    ) _
    As Boolean
On Error GoTo err_ThisFunc
Dim i As Integer
  For i = 1 To UBound(NowProj.UserChemicals)
    Call PropertyData_InitializeAll_OneChemical(i)
  Next i
exit_normally_ThisFunc:
  PropertyData_InitializeAll_AllChemicals = True
  Exit Function
exit_err_ThisFunc:
  PropertyData_InitializeAll_AllChemicals = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("PropertyData_InitializeAll_AllChemicals")
  GoTo exit_err_ThisFunc
End Function
Function PropertyData_InitializeAll_OneChemical( _
    in_idx_Chem As Integer) _
    As Boolean
On Error GoTo err_ThisFunc
Dim i1 As Integer
Dim i2 As Integer
Dim i3 As Integer
Dim UB_Hier1 As Integer
Dim UB_Hier2 As Integer
Dim UB_Hier3 As Integer
Dim Ctr_User1 As Integer
Dim Ctr_User2 As Integer
Dim AlreadyInitialized As Boolean
Dim This_Property_Code As Long
Dim in_Property_Code As Long
Dim in_Technique_Code As Long
Dim out_idx_PropertyData As Integer
Dim out_idx_TechniqueData As Integer
Dim This_UnitType As String
Dim This_UnitBase As String
Dim out_Is_FofT As Boolean
  '
  ' ERASE ALL OF THE EXISTING PROPERTY DATA; NOTE, THIS DESTROYS
  ' ALL OF THE USER DATA (IF ANY) THAT HAS BEEN INPUT.
  '
  ReDim NowProj.UserChemicals(in_idx_Chem).PropertyData(0 To 0)
  Ctr_User1 = UBound(NowProj.UserChemicals(in_idx_Chem).PropertyData)
  '
  ' ITERATE THROUGH THE PROPERTY SHEETS.
  '
  UB_Hier1 = UBound(NowProj.UserHierarchy.PropertySheetOrder)
  For i1 = 1 To UB_Hier1
    '
    ' ITERATE THROUGH THE PROPERTIES ON THE CURRENT PROPERTY SHEET.
    '
    UB_Hier2 = UBound(NowProj.UserHierarchy.PropertySheetOrder(i1).PropertyOrder)
    For i2 = 1 To UB_Hier2
      '
      ' IS THIS PROPERTY ALREADY IN THE CURRENT CHEMICAL?
      '
      This_Property_Code = NowProj.UserHierarchy. _
          PropertySheetOrder(i1). _
          PropertyOrder(i2).Property_Code
      AlreadyInitialized = False
      If (True = PropertyData_IsKeyExist(in_idx_Chem, This_Property_Code)) Then
        AlreadyInitialized = True
      End If
      If (AlreadyInitialized = False) Then
        '
        ' THIS PROPERTY HAS NOT BEEN INITIALIZED FOR THE CURRENT CHEMICAL.
        ' ADD THE PROPERTY RECORD AND INITIALIZE IT.
        '
        Ctr_User1 = Ctr_User1 + 1
        If (Ctr_User1 = 1) Then
          ReDim NowProj.UserChemicals(in_idx_Chem).PropertyData(1 To Ctr_User1)
        Else
          ReDim Preserve NowProj.UserChemicals(in_idx_Chem). _
              PropertyData(1 To Ctr_User1)
        End If
        Call PropertyData_SetDefaults(NowProj. _
            UserChemicals(in_idx_Chem). _
            PropertyData(Ctr_User1))
        NowProj.UserChemicals(in_idx_Chem). _
            PropertyData(Ctr_User1).Property_Code = _
            This_Property_Code
        '
        ' SET THE TYPE OF UNITS, THE BASE UNITS, AND THE
        ' DEFAULT UNITS OF DISPLAY.
        '
        Call Given_PropCode_Get_UnitType_and_UnitBase( _
            This_Property_Code, _
            This_UnitType, _
            This_UnitBase)
        NowProj.UserChemicals(in_idx_Chem). _
            PropertyData(Ctr_User1).UnitType = _
            This_UnitType
        NowProj.UserChemicals(in_idx_Chem). _
            PropertyData(Ctr_User1).UnitBase = _
            This_UnitBase
        NowProj.UserChemicals(in_idx_Chem). _
            PropertyData(Ctr_User1).UnitDisplayed = _
            This_UnitBase
        '
        ' SET THE f(T) NATURE OF THIS PROPERTY.
        '
        Call Given_PropCode_Get_Is_FofT(This_Property_Code, out_Is_FofT)
        NowProj.UserChemicals(in_idx_Chem). _
            PropertyData(Ctr_User1).Is_FofT = _
            out_Is_FofT
      End If
      '
      ' ITERATE THROUGH THE TECHNIQUES FOR THE CURRENT PROPERTY.
      '
      UB_Hier3 = UBound(NowProj.UserHierarchy. _
          PropertySheetOrder(i1). _
          PropertyOrder(i2).Technique_Code)
      Ctr_User2 = UBound(NowProj.UserChemicals(in_idx_Chem). _
          PropertyData(Ctr_User1). _
          TechniqueData)
      'ReDim NowProj.UserChemicals(in_idx_Chem). _
          PropertyData(Ctr_User1). _
          TechniqueData(1 To UB_Hier3)
      For i3 = 1 To UB_Hier3
        in_Property_Code = _
            This_Property_Code
        in_Technique_Code = _
            NowProj.UserHierarchy. _
            PropertySheetOrder(i1). _
            PropertyOrder(i2).Technique_Code(i3)
        '
        ' IS THIS TECHNIQUE ALREADY IN THIS PROPERTY OF THE CURRENT CHEMICAL?
        '
        Call TechniqueData_GetIndex( _
            in_idx_Chem, _
            in_Property_Code, _
            in_Technique_Code, _
            out_idx_PropertyData, _
            out_idx_TechniqueData)
        AlreadyInitialized = False
        If (out_idx_TechniqueData <> -1) Then
          AlreadyInitialized = True
        End If
        If (AlreadyInitialized = False) And (in_Property_Code <> -1) Then
          '
          ' THIS TECHNIQUE HAS NOT BEEN INITIALIZED FOR THIS PROPERTY OF THE CURRENT CHEMICAL.
          ' ADD THE TECHNIQUE RECORD AND INITIALIZE IT.
          '
          Ctr_User2 = Ctr_User2 + 1
          If (Ctr_User2 = 1) Then
            ReDim NowProj.UserChemicals(in_idx_Chem). _
                PropertyData(out_idx_PropertyData). _
                TechniqueData(1 To Ctr_User2)
          Else
            ReDim Preserve NowProj.UserChemicals(in_idx_Chem). _
                PropertyData(out_idx_PropertyData). _
                TechniqueData(1 To Ctr_User2)
          End If
          Call TechniqueData_SetDefaults(NowProj. _
              UserChemicals(in_idx_Chem). _
              PropertyData(out_idx_PropertyData). _
              TechniqueData(Ctr_User2))
          NowProj.UserChemicals(in_idx_Chem). _
              PropertyData(out_idx_PropertyData). _
              TechniqueData(Ctr_User2).Technique_Code = _
              in_Technique_Code
          If (in_Technique_Code = TECHCODE_ANY_000u_USER_INPUT) Then
            '
            ' BY DEFAULT, THERE IS _NO_ USER INPUT (UNTIL THE USER ENTERS IT!).
            '
            NowProj.UserChemicals(in_idx_Chem). _
                PropertyData(out_idx_PropertyData). _
                TechniqueData(Ctr_User2).IsAvail = False
          End If
        End If
      Next i3
    Next i2
  Next i1
exit_normally_ThisFunc:
  PropertyData_InitializeAll_OneChemical = True
  Exit Function
exit_err_ThisFunc:
  PropertyData_InitializeAll_OneChemical = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("PropertyData_InitializeAll_OneChemical")
  Resume exit_err_ThisFunc
End Function





