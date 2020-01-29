Attribute VB_Name = "Refresh"
Option Explicit




Const Refresh_declarations_end = True



Function frmTechniques_Format_txtReference( _
    in_ReferenceText As String) _
    As String
'OnError GoTo err_ThisFunc
  If (Trim$(in_ReferenceText) = "") Then
    frmTechniques_Format_txtReference = "(No reference available.)"
  Else
    frmTechniques_Format_txtReference = in_ReferenceText
  End If
exit_normally_ThisFunc:
  ''''frmTechniques_Format_txtReference = True
  Exit Function
'exit_err_ThisFunc:
'  frmTechniques_Format_txtReference = False
'  Exit Function
'err_ThisFunc:
'  Call Show_Trapped_Error("frmTechniques_Format_txtReference")
'  Resume exit_err_ThisFunc
End Function
Function frmTechniques_Populate_cbo_ALL_FOFT( _
    in_UnitType As String, _
    in_This_TechniqueData As TechniqueData_Type) _
    As Boolean
On Error GoTo err_ThisFunc
Dim Frm As Form
Set Frm = frmTechniques
Dim Ctl As Control
  '
  ' SET THE EQUATION FORM NUMBERS.
  Set Ctl = Frm.cboEqForm
  Ctl.Clear
  Ctl.AddItem "#100": Ctl.ItemData(Ctl.NewIndex) = 100
  Ctl.AddItem "#101": Ctl.ItemData(Ctl.NewIndex) = 101
  Ctl.AddItem "#102": Ctl.ItemData(Ctl.NewIndex) = 102
  Ctl.AddItem "#105": Ctl.ItemData(Ctl.NewIndex) = 105
  Ctl.AddItem "#106": Ctl.ItemData(Ctl.NewIndex) = 106
  Ctl.AddItem "#107": Ctl.ItemData(Ctl.NewIndex) = 107
  Ctl.AddItem "#114": Ctl.ItemData(Ctl.NewIndex) = 114
  Ctl.AddItem "#115": Ctl.ItemData(Ctl.NewIndex) = 115
  Ctl.AddItem "#116": Ctl.ItemData(Ctl.NewIndex) = 116
  Ctl.AddItem "#200": Ctl.ItemData(Ctl.NewIndex) = 200
  Ctl.AddItem "#201": Ctl.ItemData(Ctl.NewIndex) = 201
  Ctl.AddItem "#202": Ctl.ItemData(Ctl.NewIndex) = 202
  '
  ' SET THE TEMPERATURE UNITS.
  Set Ctl = Frm.cboUnits_T
  Call unitsys_populate_units0( _
      Ctl, _
      "temperature", _
      in_This_TechniqueData.FofT_Units_T)
  '
  ' SET THE PROPERTY UNITS.
  Set Ctl = Frm.cboUnits_f
  Call unitsys_populate_units0( _
      Ctl, _
      in_UnitType, _
      in_This_TechniqueData.FofT_Units_F)
  '
  ' EXIT OUT OF HERE.
  '
exit_normally_ThisFunc:
  frmTechniques_Populate_cbo_ALL_FOFT = True
  Exit Function
exit_err_ThisFunc:
  frmTechniques_Populate_cbo_ALL_FOFT = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("frmTechniques_Populate_cbo_ALL_FOFT")
  Resume exit_err_ThisFunc
End Function
Function frmTechniques_lvMain_Extract_Key_Info( _
    in_Key As String, _
    out_idx_Technique_Code As Integer, _
    out_idx_TechniqueData As Integer) _
    As Boolean
On Error GoTo err_ThisFunc
Dim sTemp1 As String
Dim sTemp2 As String
Dim NumArgs As Integer
  NumArgs = Parser_GetNumArgs("-", in_Key)
  If (NumArgs <> 3) Then GoTo exit_err_ThisFunc
  Call Parser_GetArg("-", in_Key, 2, sTemp1)
  Call Parser_GetArg("-", in_Key, 3, sTemp2)
  out_idx_Technique_Code = CInt(Val(sTemp1))
  out_idx_TechniqueData = CInt(Val(sTemp2))
exit_normally_ThisFunc:
  frmTechniques_lvMain_Extract_Key_Info = True
  Exit Function
exit_err_ThisFunc:
  frmTechniques_lvMain_Extract_Key_Info = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("frmTechniques_lvMain_Extract_Key_Info")
  GoTo exit_err_ThisFunc
End Function
Function frmTechniques_Populate_DataDetails( _
    ) _
    As Boolean
On Error GoTo err_ThisFunc
Dim Frm As Form
Set Frm = frmTechniques
Dim Old_Key As String
Dim out_idx_Technique_Code As Integer
Dim out_idx_TechniqueData As Integer
Dim This_Technique_Code As Long
Dim Enabled_FofT As Boolean
Dim NotBlanked_FofT As Boolean
Dim i As Integer
Dim This_TechniqueData As TechniqueData_Type
Dim This_PropertyData As PropertyData_Type
  Old_Key = "(n/a)"
  On Error Resume Next
  Old_Key = Frm.lvMain.SelectedItem.Key
  On Error GoTo err_ThisFunc
'Call debug_output("frmTechniques_Populate_DataDetails: START")
  If (Old_Key = "(n/a)") Then
    '
    ' DISPLAY BLANKS IN ALL TEXTBOXES.
    '
    Frm.txtError = ""
            '
            ' more to come ...........
            '
    GoTo exit_normally_ThisFunc
  End If
  Call frmTechniques_lvMain_Extract_Key_Info( _
      Old_Key, _
      out_idx_Technique_Code, _
      out_idx_TechniqueData)
  This_TechniqueData = _
      NowProj.UserChemicals(Frm.Window_idx_Chemical). _
      PropertyData(Frm.Window_idx_PropertyData). _
      TechniqueData(out_idx_TechniqueData)
  This_PropertyData = _
      NowProj.UserChemicals(Frm.Window_idx_Chemical). _
      PropertyData(Frm.Window_idx_PropertyData)
  With This_TechniqueData
    This_Technique_Code = .Technique_Code
    If (.Error_Code = "") Then
      Frm.txtError = "( No errors )" & _
          vbCrLf & " (testing: Technique_Code = " & _
          Trim$(Str$(.Technique_Code)) & ")"
    Else
      Frm.txtError = "Error: " & .Error_Code & _
          vbCrLf & " (testing: Technique_Code = " & _
          Trim$(Str$(.Technique_Code)) & ")"
    End If
    Frm.txtReference = _
        frmTechniques_Format_txtReference(.ReferenceText)
  End With
  With This_PropertyData
    NotBlanked_FofT = .Is_FofT
    Enabled_FofT = False
    If (NotBlanked_FofT = True) Then
      If (This_Technique_Code = TECHCODE_ANY_000u_USER_INPUT) Then
        '
        ' F-OF-T CONTROLS ARE ONLY ENABLED FOR USER INPUT WHEN
        ' THE CURRENTLY SELECTED TECHNIQUE IS USER INPUT.
        '
        Enabled_FofT = True
      End If
    End If
    If (This_TechniqueData.IsAvail = False) Then
      Enabled_FofT = False
      NotBlanked_FofT = False
    End If
  End With
  '
  ' ENABLE/DISABLE f(T) STUFF.
  '
  For i = 1 To 7
    Frm.txtData(i).Enabled = Enabled_FofT
    Frm.txtData(i).BackColor = QBColor(IIf(Enabled_FofT, 15, 7))
    Frm.lblData(i).Enabled = Enabled_FofT
  Next i
  Frm.cboEqForm.Enabled = True
  Frm.cboUnits_T.Enabled = True
  Frm.cboUnits_f.Enabled = True
  Frm.cboEqForm.Locked = Not Enabled_FofT
  Frm.cboUnits_T.Locked = Not Enabled_FofT
  Frm.cboUnits_f.Locked = Not Enabled_FofT
  Frm.cboEqForm.BackColor = QBColor(IIf(Enabled_FofT, 15, 7))
  Frm.cboUnits_T.BackColor = QBColor(IIf(Enabled_FofT, 15, 7))
  Frm.cboUnits_f.BackColor = QBColor(IIf(Enabled_FofT, 15, 7))
  Frm.lbl_cboEqForm.Enabled = Enabled_FofT
  Frm.lbl_cboUnits_T.Enabled = Enabled_FofT
  Frm.lbl_cboUnits_f.Enabled = Enabled_FofT
  If (NotBlanked_FofT = False) Then
    '
    ' THE F-OF-T CONTROLS ARE BLANKED OUT.
    '
'Call debug_output("frmTechniques_Populate_DataDetails: (1)")
    Frm.ssfTechTemp.Enabled = True
    For i = 1 To 7
      Frm.txtData(i).Text = ""
    Next i
    Frm.cboEqForm.Clear
    Frm.cboUnits_T.Clear
    Frm.cboUnits_f.Clear
    Frm.ssfTechTemp.Enabled = False
  Else
    '
    ' THE F-OF-T CONTROLS ARE _NOT_ BLANKED OUT.
    '
'Call debug_output("frmTechniques_Populate_DataDetails: (2)")
    Frm.ssfTechTemp.Enabled = True
    With This_TechniqueData
      Call frmTechniques_Populate_cbo_ALL_FOFT( _
          This_PropertyData.UnitType, _
          This_TechniqueData)
      Call unitsys_set_units0(Frm.cboUnits_T, .FofT_Units_T)
      Call unitsys_set_units0(Frm.cboUnits_f, .FofT_Units_F)
Dim New_Tag As Integer
Dim Ctl As Control
      '
      ' LOOK UP APPROPRIATE VALUE FOR cboEqForm SCROLLBOX.
      '
      Set Ctl = Frm.cboEqForm
      Frm.HALT_Controls = True
      New_Tag = 0
      For i = 0 To Ctl.ListCount - 1
        If (Ctl.ItemData(i) = .FofT_EqForm) Then
          New_Tag = i
          Exit For
        End If
      Next i
      'OnError GoTo 0
      Ctl.Tag = Trim$(Str$(New_Tag))
      Ctl.ListIndex = New_Tag
      Frm.HALT_Controls = False
      '
      ' SET UNITS.
      '
      Call unitsys_set_number_in_base_units( _
          Frm.txtData(2), .FofT_Maximum_T)
      Call unitsys_set_number_in_base_units( _
          Frm.txtData(1), .FofT_Minimum_T)
      For i = 1 To 5
        Call unitsys_set_number_in_base_units( _
            Frm.txtData(i + 2), .FofT_Coeffs(i))
      Next i
    End With
  End If
  '
  ' EXIT OUT OF HERE.
  '
exit_normally_ThisFunc:
'Call debug_output("frmTechniques_Populate_DataDetails: exit_normally_ThisFunc")
  frmTechniques_Populate_DataDetails = True
  Exit Function
exit_err_ThisFunc:
'Call debug_output("frmTechniques_Populate_DataDetails: exit_err_ThisFunc")
  frmTechniques_Populate_DataDetails = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("frmTechniques_Populate_DataDetails")
  GoTo exit_err_ThisFunc
End Function
Function frmTechniques_Populate_CurrentDataTab( _
    inout_First_Display As Boolean) _
    As Boolean
On Error GoTo err_ThisFunc
Dim Frm As Form
Set Frm = frmTechniques
Dim Ctl1 As Control
Set Ctl1 = Frm.lvMain
Dim Ctl2 As Control
Set Ctl2 = Frm.sspDipprData
Dim Ctl3 As Control
Set Ctl3 = Frm.sspPropertyNote
Dim i As Integer
Dim Old_ItemData_lstPropSheets As Integer
'Dim Old_ItemData_lstUser As Integer
Dim Old_ItemData As Integer
Dim This_ListIndex As Integer
Dim New_ListIndex As Integer
Dim ItmX As ListItem
'Dim idx_PropertySheetOrder As Integer
'Dim Name_PropertySheet As String
Dim Old_Key As String
Dim This_Key As String
'Dim This_PropertyName As String
Dim This_TechniqueName As String
Dim This_Property_Code As Long
'Dim WasSpecialCase As Boolean
Dim This_idx_PropertyData As Integer
Dim This_idx_Technique_Used As Integer
Dim This_Override_Technique_Code As Long
Dim This_IsAvail As Boolean
Dim This_Value_BaseUnits As Double
Dim This_Value_DisplayedUnits As Double
''''Dim This_Value As Double
Dim This_Value_As_String As String
Dim idx_lstPropSheets As Integer
Dim This_Technique_Code As Long
Dim out_TechCategory As Integer
Dim sTechCatName As String
Dim out_idx_PropertyData As Integer
Dim out_idx_TechniqueData As Integer
Dim This_Error_Code As String
Dim This_UnitType As String
Dim This_UnitBase As String
Dim This_UnitDisplayed As String
Dim out_Found As Integer
  Frm.HALT_Controls = True
  Ctl1.Visible = False
  Ctl2.Visible = False
  Ctl3.Visible = False
  '
  ' DETERMINE WHETHER sspTech/sspDipprData/
  ' sspPropertyNote/etc ARE VISIBLE OR INVISIBLE.
  '
  Old_ItemData_lstPropSheets = -1
  If (Frm.lstPropSheets.ListIndex >= 0) Then
    Old_ItemData_lstPropSheets = _
        Frm.lstPropSheets.ItemData(Frm.lstPropSheets.ListIndex)
  End If
  If (Old_ItemData_lstPropSheets < 0) Then
    ' KEEP CONTROLS INVISIBLE AND EXIT.
    Frm.HALT_Controls = True
    GoTo exit_normally_ThisFunc
  End If
  '
  ' MAIN CODE.
  '
  idx_lstPropSheets = _
      Frm.lstPropSheets.ItemData(Frm.lstPropSheets.ListIndex)
  Select Case idx_lstPropSheets
    Case TECHNIQUE_TAB_01a_LIST:
      '
      ' DISPLAY THE TECHNIQUE LIST (AND OTHER INFO).
      '
      Old_Key = "(n/a)"
      On Error Resume Next
      Old_Key = Ctl1.SelectedItem.Key
      On Error GoTo err_ThisFunc
      ''''OnError GoTo 0
      Ctl1.ListItems.Clear
      For i = 1 To UBound( _
          NowProj. _
          UserHierarchy. _
          PropertySheetOrder(Frm.Window_idx_PropertySheetOrder_FIRST). _
          PropertyOrder(Frm.Window_idx_PropertyOrder_FIRST). _
          Technique_Code)
        This_Technique_Code = NowProj. _
            UserHierarchy. _
            PropertySheetOrder(Frm.Window_idx_PropertySheetOrder_FIRST). _
            PropertyOrder(Frm.Window_idx_PropertyOrder_FIRST). _
            Technique_Code(i)
        Call TechniqueData_GetIndex( _
            Frm.Window_idx_Chemical, _
            Frm.Window_Property_Code, _
            This_Technique_Code, _
            out_idx_PropertyData, _
            out_idx_TechniqueData)
        This_Key = "x-" & _
            Trim$(Str$(i)) & _
            "-" & Trim$(Str$(out_idx_TechniqueData))
        Set ItmX = Ctl1.ListItems.Add(, This_Key, " ")
        Call Given_TechCode_Get_Name( _
            This_Technique_Code, _
            This_TechniqueName)
        Call Given_TechCode_Get_TechCategory( _
            This_Technique_Code, _
            out_TechCategory)
        Select Case out_TechCategory
          Case TECHCATEGORY_USER: sTechCatName = "User"
          Case TECHCATEGORY_DATA: sTechCatName = "Data"
          Case TECHCATEGORY_ESTIMATE: sTechCatName = "Est"
          Case Else: sTechCatName = "(Error)"
        End Select
        ItmX.SubItems(2) = sTechCatName
        ItmX.SubItems(3) = This_TechniqueName
        '
        ' LOOK UP THE CALCULATED VALUE FOR THIS TECHNIQUE.
        '
        With NowProj.UserChemicals(Frm.Window_idx_Chemical). _
            PropertyData(out_idx_PropertyData). _
            TechniqueData(out_idx_TechniqueData)
          This_IsAvail = .IsAvail
          This_Error_Code = .Error_Code
          This_Value_BaseUnits = .value
        End With
        With NowProj.UserChemicals(Frm.Window_idx_Chemical). _
            PropertyData(out_idx_PropertyData)
          This_idx_Technique_Used = .idx_Technique_Used
          This_Override_Technique_Code = .Override_Technique_Code
          This_UnitType = .UnitType
          This_UnitBase = .UnitBase
          This_UnitDisplayed = .UnitDisplayed
        End With
        Call unitsys_convert0( _
            This_UnitType, _
            This_UnitBase, _
            This_UnitDisplayed, _
            This_Value_BaseUnits, _
            This_Value_DisplayedUnits, _
            out_Found)
        If (out_Found = True) Then
          If (This_IsAvail = True) Then
            This_Value_As_String = Format_Numerical_Value(This_Value_DisplayedUnits)
          Else
            This_Value_As_String = "Not Available"
          End If
        Else
          This_Value_As_String = "Unit Conversion Error!"
        End If
        ItmX.SubItems(4) = This_Value_As_String
        If (i = This_idx_Technique_Used) Then
          If (This_Override_Technique_Code <> -1) Then
            ItmX.SubItems(1) = "X(ov!)"
          Else
            ItmX.SubItems(1) = "X"
          End If
        Else
          ItmX.SubItems(1) = ""
        End If
        ItmX.SubItems(5) = This_UnitDisplayed
''''Call debug_output("frmTechniques_Populate_CurrentDataTab: " & _
    "Frm.Window_idx_Chemical = " & Trim$(Str$(Frm.Window_idx_Chemical)) & ", " & _
    "Frm.Window_Property_Code = " & Trim$(Str$(Frm.Window_Property_Code)) & ", " & _
    "out_idx_PropertyData = " & Trim$(Str$(out_idx_PropertyData)) & ".")
''''Call debug_output("frmTechniques_Populate_CurrentDataTab: " & _
    "This_UnitDisplayed = `" & This_UnitDisplayed & "`")

'        This_idx_PropertyData = PropertyData_GetIndex( _
'            Old_ItemData_lstUser, _
'            This_Property_Code)
'        If (This_idx_PropertyData = -1) Then
'          This_Value_As_String = "( Error! )"
'        Else
'          With NowProj.UserChemicals(Old_ItemData_lstUser). _
'              PropertyData(This_idx_PropertyData)
'            This_idx_Technique_Used = .idx_Technique_Used
'            If (.IsAvail = True) Then
'              This_IsAvail = .TechniqueData(This_idx_Technique_Used).IsAvail
'              This_Value = .TechniqueData(This_idx_Technique_Used).Value
'            Else
'              This_IsAvail = False
'            End If
'          End With
'          If (This_IsAvail = True) Then
'            ''''This_Value_As_String = Trim$(Str$(This_Value))
'            This_Value_As_String = Format_Numerical_Value(This_Value)
'          Else
'            This_Value_As_String = "Not Available"
'          End If
'        End If
'        ItmX.SubItems(2) = This_Value_As_String
'        ItmX.SubItems(3) = "Testing2"
'        ItmX.SubItems(4) = " "
        
        If (This_IsAvail = True) Then
          ItmX.Icon = 1: ItmX.SmallIcon = 1
        Else
          ItmX.Icon = 2: ItmX.SmallIcon = 2
        End If
        If (inout_First_Display = True) Then
          If (i = This_idx_Technique_Used) Then
            '
            ' WHEN WINDOW IS FIRST DISPLAYED, SELECT THE CURRENTLY "USED" TECHNIQUE.
            inout_First_Display = False
            ItmX.Selected = True
          End If
        Else
          If (Old_Key = This_Key) Then
            '
            ' RESTORE PREVIOUSLY SELECTED ROW.
            ItmX.Selected = True
          End If
        End If
      Next i
      Ctl1.Visible = True
    Case TECHNIQUE_TAB_02a_DIPPR801, TECHNIQUE_TAB_02b_DIPPR911:
      '
      ' DISPLAY THE DIPPR801 OR DIPPR911 DATA.
      '
Dim TechDat As TechniqueData_Type
Dim Found_DIPPR_Data As Boolean
Dim sDipprName As String
      Select Case idx_lstPropSheets
        Case TECHNIQUE_TAB_02a_DIPPR801:
          This_Technique_Code = TECHCODE_ANY_992d_DB801
          sDipprName = "DIPPR801"
        Case TECHNIQUE_TAB_02b_DIPPR911:
          This_Technique_Code = TECHCODE_ANY_991d_DB911
          sDipprName = "DIPPR911"
      End Select
      Found_DIPPR_Data = False
      If (True = TechniqueData_GetIndex( _
          Frm.Window_idx_Chemical, _
          Frm.Window_Property_Code, _
          This_Technique_Code, _
          out_idx_PropertyData, _
          out_idx_TechniqueData)) Then
        TechDat = NowProj. _
            UserChemicals(Frm.Window_idx_Chemical). _
            PropertyData(out_idx_PropertyData). _
            TechniqueData(out_idx_TechniqueData)
        If (TechDat.IsAvail = True) Then
          Found_DIPPR_Data = True
        End If
      End If
      Frm.ssfDipprData.Caption = sDipprName & " Data:"
      If (Found_DIPPR_Data = False) Then
        '
        ' THE DIPPR DATA IS UNAVAILABLE; DISPLAY ERROR MESSAGE AND
        ' BLANK ALL OF THE TEXTBOXES.
        '
        Frm.lblNotAvailable.Visible = True
        Frm.lblNotAvailable.Caption = "The " & sDipprName & _
            " data is not available for this property."
        Frm.txtDataStr(0).Text = ""
        Frm.txtDataStr(1).Text = ""
        Frm.txtDataStr(2).Text = ""
        Frm.txtDataStr(3).Text = ""
        Frm.txtDataStr(4).Text = ""
        Frm.txtComment.Text = ""
        Frm.txtCitation.Text = ""
      Else
        '
        ' THE DIPPR DATA IS AVAILABLE; DISPLAY THE DATA.
        '
        Frm.lblNotAvailable.Visible = False
        With TechDat
          Frm.txtDataStr(0).Text = NowProj. _
              UserChemicals(Frm.Window_idx_Chemical).CAS
          If (.DIPPR_R = -1) Then
            Frm.txtDataStr(1).Text = ""
          Else
            Frm.txtDataStr(1).Text = Trim$(Str$(.DIPPR_R))
          End If
          Frm.txtDataStr(2).Text = .DIPPR_REL
          Frm.txtDataStr(3).Text = .DIPPR_DescMethod
          Frm.txtDataStr(4).Text = .DIPPR_Pressure
          Frm.txtComment.Text = .DIPPR_Comment
          Frm.txtCitation.Text = _
              frmTechniques_Format_txtReference(.ReferenceText)
        End With
      End If
      Ctl2.Visible = True
    Case TECHNIQUE_TAB_03a_NOTE:
      '
      ' DISPLAY THE PROPERTY NOTE.
      '
      Ctl3.Visible = True
  End Select
  Frm.HALT_Controls = False
exit_normally_ThisFunc:
  frmTechniques_Populate_CurrentDataTab = True
  Exit Function
exit_err_ThisFunc:
  frmTechniques_Populate_CurrentDataTab = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("frmTechniques_Populate_CurrentDataTab")
  Resume exit_err_ThisFunc
End Function
Function frmTechniques_Refresh() _
    As Boolean
On Error GoTo err_ThisFunc
Dim Frm As Form
Set Frm = frmTechniques
  '
  ' RESIZE FONT ON VARIOUS CONTROLS.
  '
  With PrefEnvironment
    Frm.lvMain.Font.Size = .FontSize_Lists
    Frm.lstPropSheets.Font.Size = .FontSize_Lists
    Frm.txtError.Font.Size = .FontSize_Lists
    Frm.txtReference.Font.Size = .FontSize_Lists
    ''''Frm..Font.Size = .FontSize_Lists
  End With
exit_normally_ThisFunc:
  frmTechniques_Refresh = True
  Exit Function
exit_err_ThisFunc:
  frmTechniques_Refresh = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("frmTechniques_Refresh")
  Resume exit_err_ThisFunc
End Function
Function frmTechniques_Populate_lstPropSheets() _
    As Boolean
On Error GoTo err_ThisFunc
Dim Frm As Form
Set Frm = frmTechniques
Dim Ctl As Control
Set Ctl = Frm.lstPropSheets
Dim i As Integer
Dim Old_ItemData As Integer
Dim This_ListIndex As Integer
Dim New_ListIndex As Integer
  Frm.HALT_Controls = True
  Old_ItemData = -1
  If (Ctl.ListIndex >= 0) Then
    Old_ItemData = Ctl.ItemData(Ctl.ListIndex)
  End If
  New_ListIndex = -1
  Ctl.Clear
  Ctl.AddItem "List of Techniques"
  Ctl.ItemData(Ctl.NewIndex) = TECHNIQUE_TAB_01a_LIST
  Ctl.AddItem "DIPPR801 Data"
  Ctl.ItemData(Ctl.NewIndex) = TECHNIQUE_TAB_02a_DIPPR801
  Ctl.AddItem "DIPPR911 Data"
  Ctl.ItemData(Ctl.NewIndex) = TECHNIQUE_TAB_02b_DIPPR911
  Ctl.AddItem "Property Note"
  Ctl.ItemData(Ctl.NewIndex) = TECHNIQUE_TAB_03a_NOTE
  For i = 0 To Ctl.ListCount - 1
    If (Ctl.ItemData(i) = Old_ItemData) Then New_ListIndex = i
  Next i
  If (New_ListIndex <> -1) Then
    Ctl.ListIndex = New_ListIndex
  End If
  Ctl.ListIndex = 0
  Frm.HALT_Controls = False
exit_normally_ThisFunc:
  frmTechniques_Populate_lstPropSheets = True
  Exit Function
exit_err_ThisFunc:
  frmTechniques_Populate_lstPropSheets = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("frmTechniques_Populate_lstPropSheets")
  Resume exit_err_ThisFunc
End Function


Function frmMain_Populate_lstUser() _
    As Boolean
On Error GoTo err_ThisFunc
Dim Frm As Form
Set Frm = frmMain
Dim Ctl As Control
Set Ctl = Frm.lstUser
Dim i As Integer
Dim Old_ItemData As Integer
Dim This_ListIndex As Integer
Dim New_ListIndex As Integer
  Frm.HALT_Controls = True
  Old_ItemData = -1
  If (Ctl.ListIndex >= 0) Then
    Old_ItemData = Ctl.ItemData(Ctl.ListIndex)
  End If
  New_ListIndex = -1
  Ctl.Clear
  For i = 1 To UBound(NowProj.UserChemicals)
    Ctl.AddItem NowProj.UserChemicals(i).Name
    This_ListIndex = Ctl.NewIndex
    Ctl.ItemData(This_ListIndex) = i
    If (i = Old_ItemData) Then New_ListIndex = This_ListIndex
  Next i
  If (New_ListIndex <> -1) Then
    Ctl.ListIndex = New_ListIndex
  End If
  Frm.HALT_Controls = False
exit_normally_ThisFunc:
  frmMain_Populate_lstUser = True
  Exit Function
exit_err_ThisFunc:
  frmMain_Populate_lstUser = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("frmMain_Populate_lstUser")
  Resume exit_err_ThisFunc
End Function
Function frmMain_Populate_lstPropSheets() _
    As Boolean
On Error GoTo err_ThisFunc
Dim Frm As Form
Set Frm = frmMain
Dim Ctl As Control
Set Ctl = Frm.lstPropSheets
Dim i As Integer
Dim Old_ItemData As Integer
Dim This_ListIndex As Integer
Dim New_ListIndex As Integer
  Frm.HALT_Controls = True
  Old_ItemData = -1
  If (Ctl.ListIndex >= 0) Then
    Old_ItemData = Ctl.ItemData(Ctl.ListIndex)
  End If
  New_ListIndex = -1
  Ctl.Clear
  For i = 1 To UBound(NowProj.UserHierarchy.PropertySheetOrder)
    Ctl.AddItem NowProj.UserHierarchy.PropertySheetOrder(i).Name
    This_ListIndex = Ctl.NewIndex
    Ctl.ItemData(This_ListIndex) = i
    If (i = Old_ItemData) Then New_ListIndex = This_ListIndex
  Next i
  If (New_ListIndex <> -1) Then
    Ctl.ListIndex = New_ListIndex
  End If
  Frm.HALT_Controls = False
exit_normally_ThisFunc:
  frmMain_Populate_lstPropSheets = True
  Exit Function
exit_err_ThisFunc:
  frmMain_Populate_lstPropSheets = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("frmMain_Populate_lstPropSheets")
  Resume exit_err_ThisFunc
End Function
Function frmMain_lvMain_Extract_Key_Info( _
    in_Key As String, _
    out_idx_PropertySheetOrder As Integer, _
    out_idx_PropertyOrder As Integer) _
    As Boolean
On Error GoTo err_ThisFunc
Dim sTemp1 As String
Dim sTemp2 As String
Dim NumArgs As Integer
  NumArgs = Parser_GetNumArgs("-", in_Key)
  If (NumArgs <> 3) Then GoTo exit_err_ThisFunc
  Call Parser_GetArg("-", in_Key, 2, sTemp1)
  Call Parser_GetArg("-", in_Key, 3, sTemp2)
  out_idx_PropertySheetOrder = CInt(Val(sTemp1))
  out_idx_PropertyOrder = CInt(Val(sTemp2))
exit_normally_ThisFunc:
  frmMain_lvMain_Extract_Key_Info = True
  Exit Function
exit_err_ThisFunc:
  frmMain_lvMain_Extract_Key_Info = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("frmMain_lvMain_Extract_Key_Info")
  GoTo exit_err_ThisFunc
End Function
Function frmMain_lstUser_GetItemData() _
    As Integer
On Error GoTo err_ThisFunc
Dim Frm As Form
Set Frm = frmMain
  If (Frm.lstUser.ListIndex < 0) Then
    GoTo exit_err_ThisFunc
  End If
  frmMain_lstUser_GetItemData = _
      Frm.lstUser.ItemData(Frm.lstUser.ListIndex)
exit_normally_ThisFunc:
  frmMain_lstUser_GetItemData = frmMain_lstUser_GetItemData
  Exit Function
exit_err_ThisFunc:
  frmMain_lstUser_GetItemData = -1
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("frmMain_lstUser_GetItemData")
  Resume exit_err_ThisFunc
End Function
Function frmMain_Populate_lvMain() _
    As Boolean
On Error GoTo err_ThisFunc
Dim Frm As Form
Set Frm = frmMain
Dim Ctl1 As Control
Set Ctl1 = Frm.lvMain
Dim Ctl2 As Control
Set Ctl2 = Frm.txtChemNote
Dim Ctl3 As Control
Set Ctl3 = Frm.sspBasic
Dim i As Integer
Dim Old_ItemData_lstPropSheets As Integer
Dim Old_ItemData_lstUser As Integer
Dim Old_ItemData As Integer
Dim This_ListIndex As Integer
Dim New_ListIndex As Integer
Dim ItmX As ListItem
Dim idx_PropertySheetOrder As Integer
Dim Name_PropertySheet As String
Dim Old_Key As String
Dim This_Key As String
Dim This_PropertyName As String
Dim This_Property_Code As Long
Dim WasSpecialCase As Boolean
Dim This_idx_PropertyData As Integer
Dim This_idx_Technique_Used As Integer
Dim This_IsAvail As Boolean
Dim This_Text_When_Blank As String
''''Dim This_Value As Double
Dim This_Value_As_String As String
Dim This_UnitType As String
Dim This_UnitBase As String
Dim This_UnitDisplayed As String
Dim out_Found As Integer
Dim This_Value_BaseUnits As Double
Dim This_Value_DisplayedUnits As Double
  Frm.HALT_Controls = True
  Ctl1.Visible = False
  Ctl2.Visible = False
  Ctl3.Visible = False
  '
  ' DETERMINE WHETHER lvMain/txtChemNote/etc ARE VISIBLE OR INVISIBLE.
  '
  Old_ItemData_lstPropSheets = -1
  If (Frm.lstPropSheets.ListIndex >= 0) Then
    Old_ItemData_lstPropSheets = _
        Frm.lstPropSheets.ItemData(Frm.lstPropSheets.ListIndex)
  End If
  Old_ItemData_lstUser = -1
  If (Frm.lstUser.ListIndex >= 0) Then
    Old_ItemData_lstUser = _
        Frm.lstUser.ItemData(Frm.lstUser.ListIndex)
  End If
  If (Old_ItemData_lstPropSheets < 0) Or (Old_ItemData_lstUser < 0) Then
    ' KEEP CONTROLS INVISIBLE AND EXIT.
    Frm.HALT_Controls = True
    GoTo exit_normally_ThisFunc
  End If
  '
  ' MAIN CODE.
  '
  idx_PropertySheetOrder = _
      Frm.lstPropSheets.ItemData(Frm.lstPropSheets.ListIndex)
  Name_PropertySheet = _
      Frm.lstPropSheets.List(Frm.lstPropSheets.ListIndex)
  WasSpecialCase = False
  If (UCase$(Name_PropertySheet) = _
      UCase$(PROPERTYSHEETNAME_BASIC_CHEMICAL_INFO)) Then
    '
    ' DISPLAY THE BASIC CHEMICAL INFO FRAME.
    '
    WasSpecialCase = True
    Frm.ssfMain.Caption = "Basic Chemical Information:"
    Ctl3.Visible = True
  
  
  
  'MsgBox "More code needed here ... !"
  
  
  
  
  End If
  If (UCase$(Name_PropertySheet) = _
      UCase$(PROPERTYSHEETNAME_CHEMICAL_NOTE)) Then
    '
    ' DISPLAY THE CHEMICAL NOTE.
    '
    WasSpecialCase = True
    Frm.ssfMain.Caption = "Chemical Note:"
    Ctl2.Visible = True
  End If
  If (WasSpecialCase = False) Then
    '
    ' DISPLAY PROPERTIES FOR SELECTED PROPERTY SHEET.
    '
    Old_Key = "(n/a)"
    On Error Resume Next
    Old_Key = Ctl1.SelectedItem.Key
    On Error GoTo err_ThisFunc
    'OnError GoTo 0
    Ctl1.ListItems.Clear
    For i = 1 To UBound( _
        NowProj. _
        UserHierarchy. _
        PropertySheetOrder(idx_PropertySheetOrder). _
        PropertyOrder)
      This_Key = "x-" & _
          Trim$(Str$(idx_PropertySheetOrder)) & "-" & _
          Trim$(Str$(i))
      Set ItmX = Ctl1.ListItems.Add(, This_Key, " ")
      This_Property_Code = _
          NowProj. _
          UserHierarchy. _
          PropertySheetOrder(idx_PropertySheetOrder). _
          PropertyOrder(i).Property_Code
      Call Given_PropCode_Get_Name(This_Property_Code, This_PropertyName)
      ItmX.SubItems(1) = This_PropertyName
      '
      ' LOOK UP THE CALCULATED VALUE FOR THIS PROPERTY.
      '
      This_idx_PropertyData = PropertyData_GetIndex( _
          Old_ItemData_lstUser, _
          This_Property_Code)
      If (This_idx_PropertyData = -1) Then
        This_Value_As_String = "( Error! )"
      Else
        With NowProj.UserChemicals(Old_ItemData_lstUser). _
            PropertyData(This_idx_PropertyData)
          This_idx_Technique_Used = .idx_Technique_Used
          If (.IsAvail = True) Then
            This_IsAvail = .TechniqueData(This_idx_Technique_Used).IsAvail
            ''''This_Value = .TechniqueData(This_idx_Technique_Used).Value
            This_Text_When_Blank = .TechniqueData(This_idx_Technique_Used).Text_When_Blank
          Else
            This_IsAvail = False
          End If
          This_UnitType = .UnitType
          This_UnitBase = .UnitBase
          This_UnitDisplayed = .UnitDisplayed
        End With
        If (False = TechValue_Get( _
            Old_ItemData_lstUser, _
            This_Property_Code, _
            This_Value_BaseUnits)) Then
          This_IsAvail = False
        End If
        Call unitsys_convert0( _
            This_UnitType, _
            This_UnitBase, _
            This_UnitDisplayed, _
            This_Value_BaseUnits, _
            This_Value_DisplayedUnits, _
            out_Found)
        If (out_Found = True) Then
          If (This_IsAvail = True) Then
            This_Value_As_String = Format_Numerical_Value(This_Value_DisplayedUnits)
          Else
            This_Value_As_String = "Not Available"
          End If
        Else
          This_Value_As_String = "Unit Conversion Error!"
        End If
        ''''If (This_IsAvail = True) Then
        ''''  ''''This_Value_As_String = Trim$(Str$(This_Value))
        ''''  This_Value_As_String = Format_Numerical_Value(This_Value)
        ''''Else
        ''''  This_Value_As_String = "Not Available"
        ''''End If
        If (This_Value_DisplayedUnits = 0#) And _
            (This_Text_When_Blank <> "") Then
          This_Value_As_String = This_Text_When_Blank
        End If
      End If
      ItmX.SubItems(2) = This_Value_As_String
      ItmX.SubItems(3) = This_UnitDisplayed   '"Testing2"
      ItmX.SubItems(4) = " "
      If (This_IsAvail = True) Then
        ItmX.Icon = 1: ItmX.SmallIcon = 1
      Else
        ItmX.Icon = 2: ItmX.SmallIcon = 2
      End If
      If (Old_Key = This_Key) Then
        ItmX.Selected = True
      End If
    Next i
    Frm.ssfMain.Caption = Name_PropertySheet & ":"
    Ctl1.Visible = True
  End If
  Frm.HALT_Controls = False
exit_normally_ThisFunc:
  frmMain_Populate_lvMain = True
  Exit Function
exit_err_ThisFunc:
  frmMain_Populate_lvMain = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("frmMain_Populate_lvMain")
  Resume exit_err_ThisFunc
End Function
Sub frmMain_Repopulate_Values()
Dim Frm As Form
Set Frm = frmMain
  '
  ' NUMERICAL VALUES.
  '
  Call unitsys_set_number_in_base_units(frmMain.txtData(0), NowProj.Op_T)
  Call unitsys_set_number_in_base_units(frmMain.txtData(1), NowProj.Op_P)
''''  Call unitsys_set_number_in_base_units(frmMain.txtData(2), NowProj.Mass)
''''  Call unitsys_set_number_in_base_units(frmMain.txtData(3), NowProj.FlowRate)
  '
  ' STRINGS.
  '
Dim Old_ItemData_lstUser As Integer
Dim i As Integer
  Old_ItemData_lstUser = -1
  If (Frm.lstUser.ListIndex >= 0) Then
    Old_ItemData_lstUser = _
        Frm.lstUser.ItemData(Frm.lstUser.ListIndex)
  End If
  If (Old_ItemData_lstUser = -1) Then
    For i = 0 To 5
      Frm.txtDataStr(i).Text = ""
    Next i
  Else
    With NowProj.UserChemicals(Old_ItemData_lstUser)
      Frm.txtDataStr(0).Text = .Name
      Frm.txtDataStr(1).Text = .CAS
      Frm.txtDataStr(2).Text = .SMILES
      Frm.txtDataStr(3).Text = .Formula
      Frm.txtDataStr(4).Text = .Family
      Frm.txtDataStr(5).Text = .Source
    End With
  End If
End Sub
Sub frmMain_Refresh()
Dim Frm As Form
Set Frm = frmMain
  Call frmMain_Repopulate_Values
  '
  ' RESIZE FONT ON VARIOUS CONTROLS.
  '
  With PrefEnvironment
    Frm.lvMain.Font.Size = .FontSize_Lists
    Frm.lstPropSheets.Font.Size = .FontSize_Lists
    Frm.dblstMaster.Font.Size = .FontSize_Lists
    Frm.lstUser.Font.Size = .FontSize_Lists
    Frm.txtChemNote.Font.Size = .FontSize_Lists
    ''''Frm..Font.Size = .FontSize_Lists
  End With
End Sub


Sub frmUnitsAndOrValue_Repopulate_Values()
Dim Frm As Form
Set Frm = frmUnitsAndOrValue
Dim ValueInBaseUnits As Double
Dim ValueInDisplayedUnits As Double
Dim out_Found As Integer
  ValueInBaseUnits = Frm.ValueInBaseUnits
  Call unitsys_convert0( _
      Frm.UnitType, _
      Frm.UnitBase, _
      Frm.UnitDisplayed, _
      ValueInBaseUnits, _
      ValueInDisplayedUnits, _
      out_Found)
  Call unitsys_set_number_in_base_units(Frm.txtData(0), ValueInDisplayedUnits)
''''  Call unitsys_set_number_in_base_units(frmMain.txtData(1), NowProj.Diameter)
''''  Call unitsys_set_number_in_base_units(frmMain.txtData(2), NowProj.Mass)
''''  Call unitsys_set_number_in_base_units(frmMain.txtData(3), NowProj.FlowRate)
End Sub
Sub frmUnitsAndOrValue_Refresh()
Dim Frm As Form
Set Frm = frmUnitsAndOrValue
  '
  ' REFRESH VALUES.
  '
  Call frmUnitsAndOrValue_Repopulate_Values
  '
  ' SELECT APPROPRIATE UNIT.
  '
  Frm.HALT_ALL_CONTROLS = True
Dim i As Integer
Dim Ctl As Control
Dim New_Index As Integer
  Set Ctl = Frm.lstUnits
  If (Ctl.ListCount > 0) Then
    New_Index = 0
    For i = 0 To Ctl.ListCount - 1
      If (Trim$(UCase$(Ctl.List(i))) = _
          Trim$(UCase$(Frm.UnitDisplayed))) Then
        New_Index = i
        Exit For
      End If
    Next i
    Ctl.ListIndex = New_Index
  End If
  Frm.HALT_ALL_CONTROLS = False
  ''''Debug.Print Now & " - frmUnitsAndOrValue_Refresh() - " & Frm.UnitDisplayed
End Sub


Function frmPrefEnvironment_Repopulate_Values() _
    As Boolean
On Error GoTo err_ThisFunc
Dim Frm As Form
Set Frm = frmPrefEnvironment




exit_normally_ThisFunc:
  frmPrefEnvironment_Repopulate_Values = True
  Exit Function
exit_err_ThisFunc:
  frmPrefEnvironment_Repopulate_Values = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("frmPrefEnvironment_Repopulate_Values")
  Resume exit_err_ThisFunc
End Function
Function frmPrefEnvironment_Refresh() _
    As Boolean
On Error GoTo err_ThisFunc
Dim Frm As Form
Set Frm = frmPrefEnvironment
Dim Ctl As Control
Dim i As Integer
Dim j As Integer
Dim New_Index As Integer
Dim This_ItemData As Integer
  '
  ' REPOPULATE MISC VALUES.
  '
  Call frmPrefEnvironment_Repopulate_Values
  '
  ' NUMERICAL DISPLAY FORMAT SCROLLBOXES.
  '
  Frm.HALT_ALL_CONTROLS = True
  For j = 0 To 2
    Set Ctl = Frm.cboSigFig(j)
    With PrefEnvironment
      Select Case j
        Case 0: This_ItemData = .NumFormat_Greater1000
        Case 1: This_ItemData = .NumFormat_Less0_001
        Case 2: This_ItemData = .NumFormat_Other
      End Select
    End With
    If (Ctl.ListCount > 0) Then
      New_Index = 0
      For i = 0 To Ctl.ListCount - 1
        If (Trim$(UCase$(Ctl.ItemData(i))) = _
            Trim$(UCase$(This_ItemData))) Then
          New_Index = i
          Exit For
        End If
      Next i
      Ctl.ListIndex = New_Index
      Ctl.Tag = Trim$(Str$(New_Index))
    End If
  Next j
  Frm.HALT_ALL_CONTROLS = False
  '
  ' FONT SIZES IN LISTS.
  '
  Frm.HALT_ALL_CONTROLS = True
  Set Ctl = Frm.cboFontSize
  With PrefEnvironment
    This_ItemData = .FontSize_Lists
  End With
  If (Ctl.ListCount > 0) Then
    New_Index = 0
    For i = 0 To Ctl.ListCount - 1
      If (Trim$(UCase$(Ctl.ItemData(i))) = _
          Trim$(UCase$(This_ItemData))) Then
        New_Index = i
        Exit For
      End If
    Next i
    Ctl.ListIndex = New_Index
    Ctl.Tag = Trim$(Str$(New_Index))
  End If
  Frm.HALT_ALL_CONTROLS = False
  '
  ' EXIT OUT OF HERE.
  '
exit_normally_ThisFunc:
  frmPrefEnvironment_Refresh = True
  Exit Function
exit_err_ThisFunc:
  frmPrefEnvironment_Refresh = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("frmPrefEnvironment_Refresh")
  Resume exit_err_ThisFunc
End Function


Function frmCustomProperties_Repopulate_Values() _
    As Boolean
On Error GoTo err_ThisFunc
Dim Frm As Form
Set Frm = frmCustomProperties




exit_normally_ThisFunc:
  frmCustomProperties_Repopulate_Values = True
  Exit Function
exit_err_ThisFunc:
  frmCustomProperties_Repopulate_Values = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("frmCustomProperties_Repopulate_Values")
  Resume exit_err_ThisFunc
End Function
Function frmCustomProperties_Refresh() _
    As Boolean
On Error GoTo err_ThisFunc
Dim Frm As Form
Set Frm = frmCustomProperties
Dim Ctl As Control
Dim i As Integer
Dim j As Integer
Dim New_Index As Integer
Dim Old_Index As Integer
Dim This_ItemData As Integer
Dim Old_Tag As String
Dim UB As Integer
Dim lstTop_PropSheets_Name As String
Dim lstTop_PropSheets_Index As Integer
Dim out_Name As String
Dim List_PropCodes() As Long
Dim List_PropCodes_IsSelected() As Boolean
Dim out_idx_Elem As Integer
  '
  ' REPOPULATE MISC VALUES.
  '
  Call frmCustomProperties_Repopulate_Values
  '
  ' TURN OFF CONTROLS ON THE FORM.
  '
  Frm.HALT_ALL_CONTROLS = True
  '
  '/////////////////////////////////////////////////////////////////////////////////////////////////////////
  '/////////////////////////////////////////////////////////////////////////////////////////////////////////
  '/////////////////////////////////////////////////////////////////////////////////////////////////////////
  ' REPOPULATE lstTop_PropSheets.
  '/////////////////////////////////////////////////////////////////////////////////////////////////////////
  '/////////////////////////////////////////////////////////////////////////////////////////////////////////
  '/////////////////////////////////////////////////////////////////////////////////////////////////////////
  '
  Set Ctl = Frm.lstTop_PropSheets
  Ctl.Visible = False
  Old_Tag = ""
  lstTop_PropSheets_Name = ""
  lstTop_PropSheets_Index = 0
  New_Index = 0
  If (Ctl.ListIndex >= 0) Then
    Old_Tag = Ctl.List(Ctl.ListIndex)
  End If
  UB = UBound(NowProj.UserHierarchy.PropertySheetOrder)
  Ctl.Clear
  For i = 1 To UB
    With NowProj.UserHierarchy.PropertySheetOrder(i)
      Ctl.AddItem .Name
      Ctl.ItemData(Ctl.NewIndex) = i
      If (.Name = Old_Tag) Then New_Index = Ctl.NewIndex
    End With
  Next i
  If (New_Index <= Ctl.ListCount - 1) Then
    Ctl.ListIndex = New_Index
    lstTop_PropSheets_Name = Ctl.List(Ctl.ListIndex)
    lstTop_PropSheets_Index = Ctl.ItemData(Ctl.ListIndex)
  End If
  Ctl.Visible = True
  '
  ' REPOPULATE lstTop_Props AND lstTop_PropsAll.
  '
  If ((lstTop_PropSheets_Name = "") Or _
      (lstTop_PropSheets_Name = PROPERTYSHEETNAME_BASIC_CHEMICAL_INFO) Or _
      (lstTop_PropSheets_Name = PROPERTYSHEETNAME_CHEMICAL_NOTE)) Then
    Frm.ssfTopRight.Caption = ""
    Frm.lstTop_Props.Clear
    Frm.lstTop_PropsAll.Clear
    Frm.cmdTopLeftCmds(1).Enabled = False
    Frm.cmdTopLeftCmds(2).Enabled = False
    For i = 0 To 5
      Frm.cmdTopRight1Cmds(i).Enabled = False
    Next i
  Else
    Frm.ssfTopRight.Caption = "Properties for '" & lstTop_PropSheets_Name & "':"
    Frm.cmdTopLeftCmds(1).Enabled = True
    Frm.cmdTopLeftCmds(2).Enabled = True
    For i = 0 To 5
      Frm.cmdTopRight1Cmds(i).Enabled = True
    Next i
    '
    ' ALLOW USER TO SPECIFY PROPERTIES FOR SELECTED PROPERTY SHEET.
    '
    ' POPULATE LIST OF ALL POSSIBLE PROPERTIES.
    '
    Call Get_Complete_List_of_PropCodes(List_PropCodes)
    ReDim List_PropCodes_IsSelected(1 To UBound(List_PropCodes))
    For i = 1 To UBound(List_PropCodes)
      List_PropCodes_IsSelected(i) = False
    Next i
    '
    ' REPOPULATE lstTop_Props.
    '
    Set Ctl = Frm.lstTop_Props
    Ctl.Visible = False
    Old_Tag = ""
    Old_Index = 0
    New_Index = -1
    If (Ctl.ListIndex >= 0) Then
      Old_Tag = Ctl.List(Ctl.ListIndex)
      Old_Index = Ctl.ListIndex
    End If
    UB = UBound(NowProj.UserHierarchy. _
        PropertySheetOrder(lstTop_PropSheets_Index).PropertyOrder)
    Ctl.Clear
    For i = 1 To UB
      With NowProj.UserHierarchy. _
          PropertySheetOrder(lstTop_PropSheets_Index).PropertyOrder(i)
        Call Given_PropCode_Get_Name( _
            .Property_Code, _
            out_Name)
        Ctl.AddItem out_Name
        Ctl.ItemData(Ctl.NewIndex) = .Property_Code
        If (out_Name = Old_Tag) Then New_Index = Ctl.NewIndex
        If (True = sc_ElemFind( _
            List_PropCodes, _
            .Property_Code, _
            out_idx_Elem)) Then
          List_PropCodes_IsSelected(out_idx_Elem) = True
        End If
      End With
    Next i
    If (New_Index <= Ctl.ListCount - 1) And (New_Index <> -1) Then
      Ctl.ListIndex = New_Index
      Ctl.Selected(Ctl.ListIndex) = True
    End If
    Ctl.Visible = True
    '
    ' REPOPULATE lstTop_PropsAll.
    '
    Set Ctl = Frm.lstTop_PropsAll
    Ctl.Visible = False
    Old_Tag = ""
    Old_Index = 0
    New_Index = -1
    If (Ctl.ListIndex >= 0) Then
      Old_Tag = Ctl.List(Ctl.ListIndex)
      Old_Index = Ctl.ListIndex
    End If
    UB = UBound(List_PropCodes)
    Ctl.Clear
    For i = 1 To UB
      If (List_PropCodes_IsSelected(i) = False) Then
        Call Given_PropCode_Get_Name( _
            List_PropCodes(i), _
            out_Name)
        Ctl.AddItem out_Name
        Ctl.ItemData(Ctl.NewIndex) = List_PropCodes(i)
        If (out_Name = Old_Tag) Then New_Index = Ctl.NewIndex
      End If
    Next i
    If (New_Index <= Ctl.ListCount - 1) And (New_Index <> -1) Then
      Ctl.ListIndex = New_Index
    End If
    Ctl.Visible = True
  End If
  '
  '/////////////////////////////////////////////////////////////////////////////////////////////////////////
  '/////////////////////////////////////////////////////////////////////////////////////////////////////////
  '/////////////////////////////////////////////////////////////////////////////////////////////////////////
  ' REPOPULATE lstBottom_Props.
  '/////////////////////////////////////////////////////////////////////////////////////////////////////////
  '/////////////////////////////////////////////////////////////////////////////////////////////////////////
  '/////////////////////////////////////////////////////////////////////////////////////////////////////////
  '
Dim out_List_PropCodes() As Long
Dim out_List_idxPropSheet_First_Occurrences() As Integer
Dim out_List_idxPropOrd_First_Occurrences() As Integer
Dim lstBottom_Props_Name As String
Dim lstBottom_Props_Index As Long
  Set Ctl = Frm.lstBottom_Props
  Ctl.Visible = False
  Old_Tag = ""
  lstBottom_Props_Name = ""
  lstBottom_Props_Index = 0
  New_Index = 0
  If (Ctl.ListIndex >= 0) Then
    Old_Tag = Ctl.List(Ctl.ListIndex)
  End If
  Call PropertyOrder_Get_List_of_Unique_Property_Codes( _
      out_List_PropCodes, _
      out_List_idxPropSheet_First_Occurrences, _
      out_List_idxPropOrd_First_Occurrences)
  UB = UBound(out_List_PropCodes)
  Ctl.Clear
  For i = 1 To UB
    Call Given_PropCode_Get_Name( _
        out_List_PropCodes(i), _
        out_Name)
    Ctl.AddItem out_Name
    Ctl.ItemData(Ctl.NewIndex) = out_List_PropCodes(i)
  Next i
  For i = 0 To Ctl.ListCount - 1
    If (Ctl.List(i) = Old_Tag) Then
      New_Index = i
      Ctl.ListIndex = i
      lstBottom_Props_Name = Ctl.List(Ctl.ListIndex)
      lstBottom_Props_Index = Ctl.ItemData(Ctl.ListIndex)
    End If
  Next i
  Ctl.Visible = True
  '
  ' REPOPULATE lstTechniques AND lstTechniquesAll.
  '
  If ((lstBottom_Props_Name = "")) Then
    Frm.ssfBottomRight.Caption = ""
    Frm.lstTechniques.Clear
    Frm.lstTechniquesAll.Clear
    For i = 0 To 5
      Frm.cmdBottomRight1Cmds(i).Enabled = False
    Next i
  Else
    Frm.ssfBottomRight.Caption = "Techniques for '" & lstBottom_Props_Name & "':"
    For i = 0 To 5
      Frm.cmdBottomRight1Cmds(i).Enabled = True
    Next i
    '
    ' ALLOW USER TO SPECIFY TECHNIQUES FOR SELECTED PROPERTY.
    '
    ' POPULATE LIST OF ALL POSSIBLE TECHNIQUES.
    '
Dim List_TechCodes() As Long
Dim List_TechCodes_IsSelected() As Boolean
    Call Get_Complete_List_of_TechCodes( _
        lstBottom_Props_Index, _
        List_TechCodes)
    ReDim List_TechCodes_IsSelected(1 To UBound(List_TechCodes))
    For i = 1 To UBound(List_TechCodes)
      List_TechCodes_IsSelected(i) = False
    Next i
    '
    ' REPOPULATE lstTechniques.
    '
    Set Ctl = Frm.lstTechniques
    Ctl.Visible = False
    Old_Tag = ""
    Old_Index = 0
    New_Index = -1
    If (Ctl.ListIndex >= 0) Then
      Old_Tag = Ctl.List(Ctl.ListIndex)
      Old_Index = Ctl.ListIndex
    End If
    If (False = sc_ElemFind( _
        out_List_PropCodes, _
        lstBottom_Props_Index, _
        out_idx_Elem)) Then
      GoTo exit_err_ThisFunc
    End If
Dim out_idx_Elem_ThisProp As Integer
    out_idx_Elem_ThisProp = out_idx_Elem
    UB = UBound(NowProj.UserHierarchy. _
        PropertySheetOrder( _
            out_List_idxPropSheet_First_Occurrences(out_idx_Elem_ThisProp)). _
        PropertyOrder( _
            out_List_idxPropOrd_First_Occurrences(out_idx_Elem_ThisProp)). _
        Technique_Code)
    Ctl.Clear
    For i = 1 To UB
      With NowProj.UserHierarchy. _
          PropertySheetOrder( _
              out_List_idxPropSheet_First_Occurrences(out_idx_Elem_ThisProp)). _
          PropertyOrder( _
              out_List_idxPropOrd_First_Occurrences(out_idx_Elem_ThisProp))
        Call Given_TechCode_Get_Name( _
            .Technique_Code(i), _
            out_Name)
        Ctl.AddItem out_Name
        Ctl.ItemData(Ctl.NewIndex) = .Technique_Code(i)
        If (out_Name = Old_Tag) Then New_Index = Ctl.NewIndex
        If (True = sc_ElemFind( _
            List_TechCodes, _
            .Technique_Code(i), _
            out_idx_Elem)) Then
          List_TechCodes_IsSelected(out_idx_Elem) = True
        End If
      End With
    Next i
    If (New_Index <= Ctl.ListCount - 1) And (New_Index <> -1) Then
      Ctl.ListIndex = New_Index
      Ctl.Selected(Ctl.ListIndex) = True
    End If
    Ctl.Visible = True
    '
    ' REPOPULATE lstTechniquesAll.
    '
    Set Ctl = Frm.lstTechniquesAll
    Old_Tag = ""
    Old_Index = 0
    New_Index = -1
    If (Ctl.ListIndex >= 0) Then
      Old_Tag = Ctl.List(Ctl.ListIndex)
      Old_Index = Ctl.ListIndex
    End If
    UB = UBound(List_TechCodes)
    Ctl.Clear
    For i = 1 To UB
      If (List_TechCodes_IsSelected(i) = False) Then
        Call Given_TechCode_Get_Name( _
            List_TechCodes(i), _
            out_Name)
        Ctl.AddItem out_Name
        Ctl.ItemData(Ctl.NewIndex) = List_TechCodes(i)
        If (out_Name = Old_Tag) Then New_Index = Ctl.NewIndex
      End If
    Next i
    If (New_Index <= Ctl.ListCount - 1) And (New_Index <> -1) Then
      Ctl.ListIndex = New_Index
    End If
  End If
  '
  ' TURN ON CONTROLS ON THE FORM.
  '
  Frm.HALT_ALL_CONTROLS = False
  '
  ' EXIT OUT OF HERE.
  '
exit_normally_ThisFunc:
  frmCustomProperties_Refresh = True
  Exit Function
exit_err_ThisFunc:
  frmCustomProperties_Refresh = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("frmCustomProperties_Refresh")
  Resume exit_err_ThisFunc
End Function


