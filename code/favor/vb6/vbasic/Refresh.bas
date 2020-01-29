Attribute VB_Name = "Refresh"
Option Explicit

'
' SHARED BETWEEN frmMain AND Refresh.*().
'
Global frmMain_OriginalButtonPos(1 To 7) As Long
'
' SHARED BETWEEN frmD5_AerationBasin AND frmD5*.
'
Global frmD5_AerationBasin_Temp_Plant As TYPE_PlantDiagram
Global Const INPUT_UseWhichStructure_D5 = 1
'
' SHARED BETWEEN frmD5A_CSTR AND frmD5B_Biomass.
'
Global frmD5A_CSTR_Temp_Plant As TYPE_PlantDiagram
Global Const INPUT_UseWhichStructure_D5A = 2





Const Refresh_declarations_end = True


Sub frmMain_Repopulate_Values()
Dim Frm As Form
Set Frm = frmMain
  'Call unitsys_set_number_in_base_units(Frm.txtData(0), NowProj.length)
  'Call unitsys_set_number_in_base_units(Frm.txtData(1), NowProj.Diameter)
  'Call unitsys_set_number_in_base_units(Frm.txtData(2), NowProj.Mass)
  'Call unitsys_set_number_in_base_units(Frm.txtData(3), NowProj.FlowRate)
  '
  ' MAIN BLOCK.
  '
  Call unitsys_set_number_in_base_units(Frm.txtData(4), NowProj.Plant.Flow)
  Call unitsys_set_number_in_base_units(Frm.txtData(5), NowProj.Plant.SolidsConc)
  '
  ' OUTPUT UNITS.
  '
  Call unitsys_set_number_in_base_units(Frm.txtOutput(1), NowProj.OutputRec.TotalInfluent)
  Call unitsys_set_number_in_base_units(Frm.txtOutput(2), NowProj.OutputRec.TotalEffluent)
  Call unitsys_set_number_in_base_units(Frm.txtOutput(3), NowProj.OutputRec.pr_TotalRemoved)
End Sub
Sub frmMain_Refresh()
Dim Frm As Form
Set Frm = frmMain
Dim i As Integer
Dim NowPos As Integer
Dim ValidButton As Boolean
Dim ValidButtonArray(1 To 7) As Boolean
Dim idx_PlantMenu As Integer
Dim idx_ViewMenu As Integer
Dim Ctl As Control
  Call frmMain_Repopulate_Values
  '
  ' REFRESH MISCELLANEOUS OUTPUT DATA.
  '
  With NowProj
    Frm.txtOutput(0).Text = Trim$(.Plant.ChemicalData.ContaminantName)
    'Call AssignTextAndTag(Frm.txtOutput(1), .OutputRec.TotalInfluent)
    'Call AssignTextAndTag(Frm.txtOutput(2), .OutputRec.TotalEffluent)
    'Call AssignTextAndTag(Frm.txtOutput(3), .OutputRec.pr_TotalRemoved)
  End With
  '
  ' UPDATE BUTTON POSITIONS.
  '
  NowPos = 1
  For i = 1 To 7
    ValidButton = True
    idx_PlantMenu = -1
    Select Case i
      Case 1:     'INFLUENT WEIR.
        ValidButton = NowProj.Plant.en_InfluentWeir
        idx_PlantMenu = 10
        idx_ViewMenu = 30
      Case 2:     'AERATED GRIT CHAMBER.
        ValidButton = NowProj.Plant.en_GritChamber
        idx_PlantMenu = 20
        idx_ViewMenu = 40
      Case 4:     'PRIMARY CLARIFIER WEIR.
        ValidButton = NowProj.Plant.en_PrimaryWeir
        idx_PlantMenu = 30
        idx_ViewMenu = 60
      Case 7:     'SECONDARY CLARIFIER WEIR.
        ValidButton = NowProj.Plant.en_SecondaryWeir
        idx_PlantMenu = 40
        idx_ViewMenu = 90
    End Select
    If (idx_PlantMenu <> -1) Then
      Frm.mnuPlantItem(idx_PlantMenu).Checked = ValidButton
      Frm.mnuViewItem(idx_ViewMenu).Enabled = ValidButton
    End If
    Frm.cmdMainButton(i).Visible = ValidButton
    Frm.lblButton(i).Visible = ValidButton
    ValidButtonArray(i) = ValidButton
    If (ValidButton = True) Then
      Frm.cmdMainButton(i).Left = frmMain_OriginalButtonPos(NowPos)
      Frm.lblButton(i).Left = frmMain_OriginalButtonPos(NowPos) - 67
      NowPos = NowPos + 1
    End If
  Next i
  '
  ' HIDE ALL COLUMNS ON GRID THAT SHOULD BE HIDDEN.
  '
  Set Ctl = Frm.F1Book1
Dim StandardWidth As Long
Dim ThisWidth As Long
Dim IsVisible As Boolean
  StandardWidth = Ctl.ColWidth(5)   'PRIMARY CLARIFIER.
  For i = 1 To 7
    IsVisible = ValidButtonArray(i)
    ThisWidth = IIf(IsVisible, StandardWidth, 0)
    Ctl.ColWidth(i + 2) = ThisWidth
  Next i
  '
  ' REFRESH MORE MISCELLANEOUS OUTPUT DATA TO GRID.
  '
  Set Ctl = Frm.F1Book1
Dim r As Integer
Dim c As Integer
Dim C_SORBED(0 To 6) As Double
Dim C_SOLUBLE(0 To 6) As Double
Dim dblUnitFac As Double
Dim strUnitDisplayed As String
Dim dblVals(0 To 7) As Double
  '
  ' UNIT-RELATED STUFF.
  Select Case NowProj.UnitType
    Case UnitType___ENGLISH:      'for lb/d
      dblUnitFac = 2.20462262185
      strUnitDisplayed = " ¹ lb/day"
    Case UnitType___SI:           'for kg/d
      dblUnitFac = 1#
      strUnitDisplayed = " ¹ kg/day"
  End Select
  Ctl.EntryRC(4, 2) = strUnitDisplayed
  With NowProj
    r = 5
    c = 3
    '
    ' DISSOLVED EFFLUENT LIQUID CONCENTRATION
    C_SOLUBLE(0) = .OutputRec.InfluentWeir.EffluentConc
    C_SOLUBLE(1) = .OutputRec.GritChamber.EffluentConc
    C_SOLUBLE(2) = .OutputRec.PrimaryClarifier.EffluentConc
    C_SOLUBLE(3) = .OutputRec.PrimaryWeir.EffluentConc
    C_SOLUBLE(4) = .OutputRec.AerationBasin.EffluentConc
    C_SOLUBLE(5) = .OutputRec.SecondaryClarifier.EffluentConc
    C_SOLUBLE(6) = .OutputRec.SecondaryWeir.EffluentConc
    For i = 0 To 6
      Call frmMain_WriteNum_NU(Ctl, r + 0, c + i, C_SOLUBLE(i))
    Next i
    '
    ' EFFLUENT SOLIDS CONCENTRATION
    For i = 0 To 6
      C_SORBED(i) = NowProj.KP1_OUT * NowProj.XVALS_OUT(i + 1) * C_SOLUBLE(i)
      Call frmMain_WriteNum_NU(Ctl, r + 2, c + i, NowProj.XVALS_OUT(i + 1))
    Next i
    '
    ' SORBED EFFLUENT LIQUID CONCENTRATION
    For i = 0 To 6
      Call frmMain_WriteNum_NU(Ctl, r + 1, c + i, C_SORBED(i))
    Next i
    '
    ' STRIPPING
    dblVals(0) = .OutputRec.InfluentWeir.Stripping + 0#
    dblVals(1) = .OutputRec.GritChamber.Stripping + .OutputRec.GritChamber.Volatilization
    dblVals(2) = 0# + .OutputRec.PrimaryClarifier.Volatilization
    dblVals(3) = .OutputRec.PrimaryWeir.Stripping + 0#
    dblVals(4) = .OutputRec.AerationBasin.Stripping + .OutputRec.AerationBasin.Volatilization
    dblVals(5) = 0# + .OutputRec.SecondaryClarifier.Volatilization
    dblVals(6) = .OutputRec.SecondaryWeir.Stripping + 0#
    dblVals(7) = .OutputRec.TotalAmount.Stripping + .OutputRec.TotalAmount.Volatilization
    For i = 0 To 7
      Call frmMain_WriteNum_NU(Ctl, r + 3, c + i, dblVals(i) * dblUnitFac)
    Next i
    '
    ' STRIPPING % OF TOTAL
    Call frmMain_WriteNum_NU(Ctl, r + 4, c + 0, .OutputRec.InfluentWeir.pr_Stripping + 0#)
    Call frmMain_WriteNum_NU(Ctl, r + 4, c + 1, .OutputRec.GritChamber.pr_Stripping + .OutputRec.GritChamber.pr_Volatilization)
    Call frmMain_WriteNum_NU(Ctl, r + 4, c + 2, 0# + .OutputRec.PrimaryClarifier.pr_Volatilization)
    Call frmMain_WriteNum_NU(Ctl, r + 4, c + 3, .OutputRec.PrimaryWeir.pr_Stripping + 0#)
    Call frmMain_WriteNum_NU(Ctl, r + 4, c + 4, .OutputRec.AerationBasin.pr_Stripping + .OutputRec.AerationBasin.pr_Volatilization)
    Call frmMain_WriteNum_NU(Ctl, r + 4, c + 5, 0# + .OutputRec.SecondaryClarifier.pr_Volatilization)
    Call frmMain_WriteNum_NU(Ctl, r + 4, c + 6, .OutputRec.SecondaryWeir.pr_Stripping + 0#)
    Call frmMain_WriteNum_NU(Ctl, r + 4, c + 7, .OutputRec.TotalAmount.pr_Stripping + .OutputRec.TotalAmount.pr_Volatilization)
''    '
''    ' VOLATILIZATION
    Call frmMain_WriteNum_NU(Ctl, r + 2, c + 0, .OutputRec.InfluentWeir.Volatilization)
    Call frmMain_WriteNum_NU(Ctl, r + 2, c + 1, .OutputRec.GritChamber.Volatilization)
    Call frmMain_WriteNum_NU(Ctl, r + 2, c + 2, .OutputRec.PrimaryClarifier.Volatilization)
    Call frmMain_WriteNum_NU(Ctl, r + 2, c + 3, .OutputRec.PrimaryWeir.Volatilization)
    Call frmMain_WriteNum_NU(Ctl, r + 2, c + 4, .OutputRec.AerationBasin.Volatilization)
    Call frmMain_WriteNum_NU(Ctl, r + 2, c + 5, .OutputRec.SecondaryClarifier.Volatilization)
    Call frmMain_WriteNum_NU(Ctl, r + 2, c + 6, .OutputRec.SecondaryWeir.Volatilization)
    Call frmMain_WriteNum_NU(Ctl, r + 2, c + 7, .OutputRec.TotalAmount.Volatilization)
    '
''    ' VOLATILIZATION % OF TOTAL
''    Call frmMain_WriteNum_NU(Ctl, r + 4, C + 1, .OutputRec.GritChamber.pr_Volatilization)
''    Call frmMain_WriteNum_NU(Ctl, r + 4, C + 2, .OutputRec.PrimaryClarifier.pr_Volatilization)
''    Call frmMain_WriteNum_NU(Ctl, r + 4, C + 4, .OutputRec.AerationBasin.pr_Volatilization)
''    Call frmMain_WriteNum_NU(Ctl, r + 4, C + 5, .OutputRec.SecondaryClarifier.pr_Volatilization)
''    Call frmMain_WriteNum_NU(Ctl, r + 4, C + 7, .OutputRec.TotalAmount.pr_Volatilization)
    '
    ' WASTAGE (SOLID WASTE + LIQUID WASTE)
    Call frmMain_WriteNum_NU(Ctl, r + 5, c + 2, (.OutputRec.PrimaryClarifier.SolidWaste + .OutputRec.PrimaryClarifier.LiquidWaste) * dblUnitFac)
    Call frmMain_WriteNum_NU(Ctl, r + 5, c + 5, (.OutputRec.SecondaryClarifier.SolidWaste + .OutputRec.SecondaryClarifier.LiquidWaste) * dblUnitFac)
    Call frmMain_WriteNum_NU(Ctl, r + 5, c + 7, (.OutputRec.TotalAmount.SolidWaste + .OutputRec.TotalAmount.LiquidWaste) * dblUnitFac)
    '
    ' WASTAGE (SOLID WASTE + LIQUID WASTE) % OF TOTAL
    Call frmMain_WriteNum_NU(Ctl, r + 6, c + 2, .OutputRec.PrimaryClarifier.pr_SolidWaste + .OutputRec.PrimaryClarifier.pr_LiquidWaste)
    Call frmMain_WriteNum_NU(Ctl, r + 6, c + 5, .OutputRec.SecondaryClarifier.pr_SolidWaste + .OutputRec.SecondaryClarifier.pr_LiquidWaste)
    Call frmMain_WriteNum_NU(Ctl, r + 6, c + 7, .OutputRec.TotalAmount.pr_SolidWaste + .OutputRec.TotalAmount.pr_LiquidWaste)
    '
    ' BIODEGRADATION
    Call frmMain_WriteNum_NU(Ctl, r + 7, c + 4, (.OutputRec.AerationBasin.Biodegredation) * dblUnitFac)
    Call frmMain_WriteNum_NU(Ctl, r + 7, c + 7, (.OutputRec.TotalAmount.Biodegredation) * dblUnitFac)
    '
    ' BIODEGRADATION % OF TOTAL
    Call frmMain_WriteNum_NU(Ctl, r + 8, c + 4, .OutputRec.AerationBasin.pr_Biodegredation)
    Call frmMain_WriteNum_NU(Ctl, r + 8, c + 7, .OutputRec.TotalAmount.pr_Biodegredation)
  End With
''''  '
''''  ' REFRESH MORE MISCELLANEOUS OUTPUT DATA TO GRID.
''''  '
''''Dim r As Integer
''''Dim C As Integer
''''  With NowProj
''''    r = 5
''''    C = 3
''''    Set Ctl = Frm.F1Book1
''''    ' EFFLUENT LIQUID CONCENTRATION
''''    Call frmMain_WriteNum_NU(Ctl, r + 0, C + 0, .OutputRec.InfluentWeir.EffluentConc)
''''    Call frmMain_WriteNum_NU(Ctl, r + 0, C + 1, .OutputRec.GritChamber.EffluentConc)
''''    Call frmMain_WriteNum_NU(Ctl, r + 0, C + 2, .OutputRec.PrimaryClarifier.EffluentConc)
''''    Call frmMain_WriteNum_NU(Ctl, r + 0, C + 3, .OutputRec.PrimaryWeir.EffluentConc)
''''    Call frmMain_WriteNum_NU(Ctl, r + 0, C + 4, .OutputRec.AerationBasin.EffluentConc)
''''    Call frmMain_WriteNum_NU(Ctl, r + 0, C + 5, .OutputRec.SecondaryClarifier.EffluentConc)
''''    Call frmMain_WriteNum_NU(Ctl, r + 0, C + 6, .OutputRec.SecondaryWeir.EffluentConc)
''''    '
''''    ' STRIPPING
''''    Call frmMain_WriteNum_NU(Ctl, r + 1, C + 0, .OutputRec.InfluentWeir.Stripping)
''''    Call frmMain_WriteNum_NU(Ctl, r + 1, C + 1, .OutputRec.GritChamber.Stripping)
''''    Call frmMain_WriteNum_NU(Ctl, r + 1, C + 3, .OutputRec.PrimaryWeir.Stripping)
''''    Call frmMain_WriteNum_NU(Ctl, r + 1, C + 4, .OutputRec.AerationBasin.Stripping)
''''    Call frmMain_WriteNum_NU(Ctl, r + 1, C + 6, .OutputRec.SecondaryWeir.Stripping)
''''    Call frmMain_WriteNum_NU(Ctl, r + 1, C + 7, .OutputRec.TotalAmount.Stripping)
''''    '
''''    ' STRIPPING % OF TOTAL
''''    Call frmMain_WriteNum_NU(Ctl, r + 2, C + 0, .OutputRec.InfluentWeir.pr_Stripping)
''''    Call frmMain_WriteNum_NU(Ctl, r + 2, C + 1, .OutputRec.GritChamber.pr_Stripping)
''''    Call frmMain_WriteNum_NU(Ctl, r + 2, C + 3, .OutputRec.PrimaryWeir.pr_Stripping)
''''    Call frmMain_WriteNum_NU(Ctl, r + 2, C + 4, .OutputRec.AerationBasin.pr_Stripping)
''''    Call frmMain_WriteNum_NU(Ctl, r + 2, C + 6, .OutputRec.SecondaryWeir.pr_Stripping)
''''    Call frmMain_WriteNum_NU(Ctl, r + 2, C + 7, .OutputRec.TotalAmount.pr_Stripping)
''''    '
''''    ' VOLATILIZATION
''''    Call frmMain_WriteNum_NU(Ctl, r + 3, C + 1, .OutputRec.GritChamber.Volatilization)
''''    Call frmMain_WriteNum_NU(Ctl, r + 3, C + 2, .OutputRec.PrimaryClarifier.Volatilization)
''''    Call frmMain_WriteNum_NU(Ctl, r + 3, C + 4, .OutputRec.AerationBasin.Volatilization)
''''    Call frmMain_WriteNum_NU(Ctl, r + 3, C + 5, .OutputRec.SecondaryClarifier.Volatilization)
''''    Call frmMain_WriteNum_NU(Ctl, r + 3, C + 7, .OutputRec.TotalAmount.Volatilization)
''''    '
''''    ' VOLATILIZATION % OF TOTAL
''''    Call frmMain_WriteNum_NU(Ctl, r + 4, C + 1, .OutputRec.GritChamber.pr_Volatilization)
''''    Call frmMain_WriteNum_NU(Ctl, r + 4, C + 2, .OutputRec.PrimaryClarifier.pr_Volatilization)
''''    Call frmMain_WriteNum_NU(Ctl, r + 4, C + 4, .OutputRec.AerationBasin.pr_Volatilization)
''''    Call frmMain_WriteNum_NU(Ctl, r + 4, C + 5, .OutputRec.SecondaryClarifier.pr_Volatilization)
''''    Call frmMain_WriteNum_NU(Ctl, r + 4, C + 7, .OutputRec.TotalAmount.pr_Volatilization)
''''    '
''''    ' WASTAGE (SOLID WASTE + LIQUID WASTE)
''''    Call frmMain_WriteNum_NU(Ctl, r + 5, C + 2, .OutputRec.PrimaryClarifier.SolidWaste + .OutputRec.PrimaryClarifier.LiquidWaste)
''''    Call frmMain_WriteNum_NU(Ctl, r + 5, C + 5, .OutputRec.SecondaryClarifier.SolidWaste + .OutputRec.SecondaryClarifier.LiquidWaste)
''''    Call frmMain_WriteNum_NU(Ctl, r + 5, C + 7, .OutputRec.TotalAmount.SolidWaste + .OutputRec.TotalAmount.LiquidWaste)
''''    '
''''    ' WASTAGE (SOLID WASTE + LIQUID WASTE) % OF TOTAL
''''    Call frmMain_WriteNum_NU(Ctl, r + 6, C + 2, .OutputRec.PrimaryClarifier.pr_SolidWaste + .OutputRec.PrimaryClarifier.pr_LiquidWaste)
''''    Call frmMain_WriteNum_NU(Ctl, r + 6, C + 5, .OutputRec.SecondaryClarifier.pr_SolidWaste + .OutputRec.SecondaryClarifier.pr_LiquidWaste)
''''    Call frmMain_WriteNum_NU(Ctl, r + 6, C + 7, .OutputRec.TotalAmount.pr_SolidWaste + .OutputRec.TotalAmount.pr_LiquidWaste)
''''    '
''''    ' BIODEGRADATION
''''    Call frmMain_WriteNum_NU(Ctl, r + 7, C + 4, .OutputRec.AerationBasin.Biodegredation)
''''    Call frmMain_WriteNum_NU(Ctl, r + 7, C + 7, .OutputRec.TotalAmount.Biodegredation)
''''    '
''''    ' BIODEGRADATION % OF TOTAL
''''    Call frmMain_WriteNum_NU(Ctl, r + 8, C + 4, .OutputRec.AerationBasin.pr_Biodegredation)
''''    Call frmMain_WriteNum_NU(Ctl, r + 8, C + 7, .OutputRec.TotalAmount.pr_Biodegredation)
''''  End With
End Sub


Function frmMain_NumFormat(Value As Variant) As String
  frmMain_NumFormat = GetDoubleFormat(Value)
End Function
Sub frmMain_WriteNum_NU( _
    in_tCtl As Control, _
    in_Row As Integer, _
    in_Col As Integer, _
    in_Num As Double _
    )
Dim UseNum As Double
  UseNum = in_Num
  in_tCtl.SelEndCol = in_Col
  in_tCtl.SelStartCol = in_Col
  in_tCtl.SelEndRow = in_Row
  in_tCtl.SelStartRow = in_Row
  in_tCtl.NumberFormat = frmMain_NumFormat(UseNum)
  in_tCtl.NumberRC(in_Row, in_Col) = UseNum
End Sub


Sub frmD0_Props_Repopulate_Values(Temp_Plant As TYPE_PlantDiagram)
Dim Frm As Form
Set Frm = frmD0_Props
Dim ThisVal As Double
Dim Is_Unavailable As Boolean
Dim IsLocked As Boolean
Dim Is_NonEditable As Boolean
Dim i As Integer
Dim Ctl As Control
  '
  ' STRINGS (NAME OF CONTAMINANT).
  '
  With Temp_Plant.ChemicalData
    Call AssignTextAndTag( _
        Frm.txtDataStr(0), Trim$(.ContaminantName))
  End With
  '
  ' MAIN BLOCK.
  '
  For i = 0 To 18
    Set Ctl = Frm.txtData(i)
    With Temp_Plant.ChemicalData.DataSources(i)
      Select Case .SourceType
        Case DATASOURCETYPE_USERINPUT: ThisVal = .Val_UserInput
        Case DATASOURCETYPE_STEPP: ThisVal = .Val_StEPP
        Case DATASOURCETYPE_CORR: ThisVal = .Val_Corr
      End Select
      Is_NonEditable = IIf(.SourceType = DATASOURCETYPE_USERINPUT, False, True)
      Is_Unavailable = IIf(ThisVal <= -1E+20, True, False)
      If (Is_Unavailable = True) Then
        Ctl.Text = "Unavailable"
      Else
        Call unitsys_set_number_in_base_units( _
            Ctl, _
            ThisVal)
      End If
      IsLocked = Is_Unavailable
      If (Is_NonEditable) Then IsLocked = True
      Ctl.Locked = IsLocked
      Ctl.BackColor = _
          IIf(Is_Unavailable, QBColor(4), _
              IIf(Is_NonEditable, QBColor(7), QBColor(15)))
      Ctl.ForeColor = _
          IIf(Is_Unavailable, QBColor(14), _
              IIf(Is_NonEditable, QBColor(8), QBColor(0)))
    End With
  Next i
End Sub
Sub frmD0_Props_Refresh(Temp_Plant As TYPE_PlantDiagram)
Dim Frm As Form
Set Frm = frmD0_Props
Dim Ctl As Control
Dim i As Integer
Dim j As Integer
Dim New_Tag As Integer
  '
  ' SET .Val_Corr FOR THE WATER AND AIR CORRELATIONS.
  '
  Call Corr_SetWaterAndAirAndOxygen(Temp_Plant)
  '
  ' REPOPULATE CURRENT VALUES TO TEXTBOX CONTROLS.
  '
  Call frmD0_Props_Repopulate_Values(Temp_Plant)
  '
  ' SELECT THE APPROPRIATE cboSource SETTINGS.
  '
  Frm.HALT_cboSource = True
  New_Tag = 0
  For i = 0 To 18
    Set Ctl = Frm.cboSource(i)
    With Temp_Plant.ChemicalData.DataSources(i)
      For j = 0 To Ctl.ListCount - 1
        If (Ctl.ItemData(j) = .SourceType) Then
          New_Tag = j
          Exit For
        End If
      Next j
    End With
    Ctl.Tag = Trim$(Str$(New_Tag))
    Ctl.ListIndex = New_Tag       'ctl.index
  Next i
  Frm.HALT_cboSource = False
End Sub


Sub frmD1_InfluentWeir_Repopulate_Values(Temp_Plant As TYPE_PlantDiagram)
Dim Frm As Form
Set Frm = frmD1_InfluentWeir
  '
  ' MAIN BLOCK.
  '
  With Temp_Plant.InfluentWeir
    Call unitsys_set_number_in_base_units(Frm.txtData(0), .Width)
    Call unitsys_set_number_in_base_units(Frm.txtData(1), .WaterLevelDiff)
    Call unitsys_set_number_in_base_units(Frm.txtData(2), .GasFlow)
  End With
End Sub
Sub frmD1_InfluentWeir_Refresh(Temp_Plant As TYPE_PlantDiagram)
Dim Frm As Form
Set Frm = frmD1_InfluentWeir
Dim New_Tag As Integer
Dim i As Integer
Dim Ctl As Control
  Call frmD1_InfluentWeir_Repopulate_Values(Temp_Plant)
  '
  ' LOOK UP APPROPRIATE VALUE FOR cbo_Model_Type SCROLLBOX.
  '
  Set Ctl = Frm.cbo_Model_Type
  Frm.HALT_cbo_Model_Type = True
  New_Tag = 0
  For i = 0 To Ctl.ListCount - 1
    If (Ctl.ItemData(i) = Temp_Plant.InfluentWeir.ModelingMechanism) Then
      New_Tag = i
      Exit For
    End If
  Next i
  Ctl.Tag = Trim$(Str$(New_Tag))
  Ctl.ListIndex = New_Tag
  Frm.HALT_cbo_Model_Type = False
End Sub


Sub frmD2_GritChamber_Repopulate_Values(Temp_Plant As TYPE_PlantDiagram)
Dim Frm As Form
Set Frm = frmD2_GritChamber
  '
  ' MAIN BLOCK.
  '
  With Temp_Plant.GritChamber
    Call unitsys_set_number_in_base_units(Frm.txtData(0), CDbl(.Count))
    Call unitsys_set_number_in_base_units(Frm.txtData(1), .VentilationRate)
    Call unitsys_set_number_in_base_units(Frm.txtData(2), .Depth)
    Call unitsys_set_number_in_base_units(Frm.txtData(3), .Volume)
    Call unitsys_set_number_in_base_units(Frm.txtData(4), .GasFlow)
    Call unitsys_set_number_in_base_units(Frm.txtData(5), .SOTR)
  End With
End Sub
Sub frmD2_GritChamber_Refresh(Temp_Plant As TYPE_PlantDiagram)
Dim Frm As Form
Set Frm = frmD2_GritChamber
Dim NewSetting As Boolean
Dim NewTag As Integer
  Call frmD2_GritChamber_Repopulate_Values(Temp_Plant)
  '
  ' SELECT APPROPRIATE opt_IsCovered SETTING.
  '
  With Temp_Plant.GritChamber
    NewSetting = .IsCovered
  End With
  NewTag = IIf(NewSetting, 1, 0)
  Frm.HALT_opt_IsCovered = True
  Frm.opt_IsCovered(NewTag).Value = True
  Frm.opt_IsCovered(1 - NewTag).Value = False
  Frm.opt_IsCovered(0).Enabled = True
  Frm.opt_IsCovered(1).Enabled = True
  Frm.opt_IsCovered(0).Tag = Trim$(Str$(NewTag))
  'Frm.cmdSpecifyTimeVariableInfluentConc.Enabled = NewSetting
  Frm.lblData(1).Enabled = NewSetting
  Frm.txtData(1).Enabled = NewSetting
  Frm.cboUnits(1).Enabled = NewSetting
  Frm.HALT_opt_IsCovered = False
End Sub


Sub frmD3_PrimaryClarifier_Repopulate_Values(Temp_Plant As TYPE_PlantDiagram)
Dim Frm As Form
Set Frm = frmD3_PrimaryClarifier
  '
  ' MAIN BLOCK.
  '
  With Temp_Plant.PrimaryClarifier
    Call unitsys_set_number_in_base_units(Frm.txtData(0), CDbl(.Count))
    Call unitsys_set_number_in_base_units(Frm.txtData(1), .VentilationRate)
    Call unitsys_set_number_in_base_units(Frm.txtData(2), .Depth)
    Call unitsys_set_number_in_base_units(Frm.txtData(3), .Volume)
    Call unitsys_set_number_in_base_units(Frm.txtData(4), .WastageFlow)
    Call unitsys_set_number_in_base_units(Frm.txtData(5), .PercentageRemoval)
  End With
End Sub
Sub frmD3_PrimaryClarifier_Refresh(Temp_Plant As TYPE_PlantDiagram)
Dim Frm As Form
Set Frm = frmD3_PrimaryClarifier
Dim NewSetting As Boolean
Dim NewTag As Integer
Dim i As Integer
Dim j As Integer
Dim LookingFor As Integer
Dim Ctl As Control
Dim New_Tag As Integer
  Call frmD3_PrimaryClarifier_Repopulate_Values(Temp_Plant)
  '
  ' SELECT APPROPRIATE opt_IsCovered SETTING.
  '
  With Temp_Plant.PrimaryClarifier
    NewSetting = .IsCovered
  End With
  NewTag = IIf(NewSetting, 1, 0)
  Frm.HALT_opt_IsCovered = True
  Frm.opt_IsCovered(NewTag).Value = True
  Frm.opt_IsCovered(1 - NewTag).Value = False
  Frm.opt_IsCovered(0).Enabled = True
  Frm.opt_IsCovered(1).Enabled = True
  Frm.opt_IsCovered(0).Tag = Trim$(Str$(NewTag))
  'Frm.cmdSpecifyTimeVariableInfluentConc.Enabled = NewSetting
  Frm.lblData(1).Enabled = NewSetting
  Frm.txtData(1).Enabled = NewSetting
  Frm.cboUnits(1).Enabled = NewSetting
  Frm.HALT_opt_IsCovered = False
  '
  ' LOOK UP APPROPRIATE VALUE FOR cbo_RemovalMechanism SCROLLBOXES.
  '
  For j = 0 To 1
    With Temp_Plant.PrimaryClarifier
      Select Case j
        Case 0:
          Set Ctl = Frm.cbo_RemovalMechanism(0)
          LookingFor = .SorptionRemovalMethod
        Case 1:
          Set Ctl = Frm.cbo_RemovalMechanism(1)
          LookingFor = .VolatilizationRemovalMechanism
      End Select
    End With
    Frm.HALT_cbo_RemovalMechanism = True
    New_Tag = 0
    For i = 0 To Ctl.ListCount - 1
      If (Ctl.ItemData(i) = LookingFor) Then
        New_Tag = i
        Exit For
      End If
    Next i
    Ctl.Tag = Trim$(Str$(New_Tag))
    Ctl.ListIndex = New_Tag
    Frm.HALT_cbo_RemovalMechanism = False
  Next j
End Sub


Sub frmD4_PrimaryWeir_Repopulate_Values(Temp_Plant As TYPE_PlantDiagram)
Dim Frm As Form
Set Frm = frmD4_PrimaryWeir
  '
  ' MAIN BLOCK.
  '
  With Temp_Plant.PrimaryWeir
    Call unitsys_set_number_in_base_units(Frm.txtData(0), .Width)
    Call unitsys_set_number_in_base_units(Frm.txtData(1), .WaterLevelDiff)
    Call unitsys_set_number_in_base_units(Frm.txtData(2), .GasFlow)
  End With
End Sub
Sub frmD4_PrimaryWeir_Refresh(Temp_Plant As TYPE_PlantDiagram)
Dim Frm As Form
Set Frm = frmD4_PrimaryWeir
Dim New_Tag As Integer
Dim i As Integer
Dim Ctl As Control
  Call frmD4_PrimaryWeir_Repopulate_Values(Temp_Plant)
  '
  ' LOOK UP APPROPRIATE VALUE FOR cbo_Model_Type SCROLLBOX.
  '
  Set Ctl = Frm.cbo_Model_Type
  Frm.HALT_cbo_Model_Type = True
  New_Tag = 0
  With Temp_Plant.PrimaryWeir
    For i = 0 To Ctl.ListCount - 1
      If (Ctl.ItemData(i) = .ModelingMechanism) Then
        New_Tag = i
        Exit For
      End If
    Next i
  End With
  Ctl.Tag = Trim$(Str$(New_Tag))
  Ctl.ListIndex = New_Tag
  Frm.HALT_cbo_Model_Type = False
End Sub


Sub frmD7_SecondaryWeir_Repopulate_Values(Temp_Plant As TYPE_PlantDiagram)
Dim Frm As Form
Set Frm = frmD7_SecondaryWeir
  '
  ' MAIN BLOCK.
  '
  With Temp_Plant.SecondaryWeir
    Call unitsys_set_number_in_base_units(Frm.txtData(0), .Width)
    Call unitsys_set_number_in_base_units(Frm.txtData(1), .WaterLevelDiff)
    Call unitsys_set_number_in_base_units(Frm.txtData(2), .GasFlow)
  End With
End Sub
Sub frmD7_SecondaryWeir_Refresh(Temp_Plant As TYPE_PlantDiagram)
Dim Frm As Form
Set Frm = frmD7_SecondaryWeir
Dim New_Tag As Integer
Dim i As Integer
Dim Ctl As Control
  Call frmD7_SecondaryWeir_Repopulate_Values(Temp_Plant)
  '
  ' LOOK UP APPROPRIATE VALUE FOR cbo_Model_Type SCROLLBOX.
  '
  Set Ctl = Frm.cbo_Model_Type
  Frm.HALT_cbo_Model_Type = True
  New_Tag = 0
  With Temp_Plant.SecondaryWeir
    For i = 0 To Ctl.ListCount - 1
      If (Ctl.ItemData(i) = .ModelingMechanism) Then
        New_Tag = i
        Exit For
      End If
    Next i
  End With
  Ctl.Tag = Trim$(Str$(New_Tag))
  Ctl.ListIndex = New_Tag
  Frm.HALT_cbo_Model_Type = False
End Sub


Sub frmD5_AerationBasin_Repopulate_Values(Temp_Plant As TYPE_PlantDiagram)
Dim Frm As Form
Set Frm = frmD5_AerationBasin
  '
  ' MAIN BLOCK.
  '
  With Temp_Plant.AerationBasin
    Call unitsys_set_number_in_base_units(Frm.txtData(0), CDbl(.Count))
    Call unitsys_set_number_in_base_units(Frm.txtData(1), .VentilationRate)
    Call unitsys_set_number_in_base_units(Frm.txtData(2), .Depth)
    Call unitsys_set_number_in_base_units(Frm.txtData(3), .WastageFlow)
    Call unitsys_set_number_in_base_units(Frm.txtData(4), .RecycleFlow)
    Call unitsys_set_number_in_base_units(Frm.txtData(5), .SOTR)
    Call unitsys_set_number_in_base_units(Frm.txtData(6), .Volume)
    Call unitsys_set_number_in_base_units(Frm.txtData(7), .GasFlow)
    Call unitsys_set_number_in_base_units(Frm.txtData(8), .BioMass)
    Call unitsys_set_number_in_base_units(Frm.txtData(9), CDbl(.CSTR.Count))
  End With
  With Temp_Plant.SecondaryClarifier
    Call unitsys_set_number_in_base_units(Frm.txtData(10), .EffluentSolidsConc)
  End With
End Sub
Sub frmD5_AerationBasin_Refresh(Temp_Plant As TYPE_PlantDiagram)
Dim Frm As Form
Set Frm = frmD5_AerationBasin
Dim NewSetting As Boolean
Dim NewTag As Integer
  Call frmD5_AerationBasin_Repopulate_Values(Temp_Plant)
  '
  ' SELECT APPROPRIATE opt_IsCovered SETTING.
  '
  With Temp_Plant.AerationBasin
    NewSetting = .IsCovered
  End With
  NewTag = IIf(NewSetting, 1, 0)
  Frm.HALT_opt_IsCovered = True
  Frm.opt_IsCovered(NewTag).Value = True
  Frm.opt_IsCovered(1 - NewTag).Value = False
  Frm.opt_IsCovered(0).Enabled = True
  Frm.opt_IsCovered(1).Enabled = True
  Frm.opt_IsCovered(0).Tag = Trim$(Str$(NewTag))
  'Frm.cmdSpecifyTimeVariableInfluentConc.Enabled = NewSetting
  Frm.lblData(1).Enabled = NewSetting
  Frm.txtData(1).Enabled = NewSetting
  Frm.cboUnits(1).Enabled = NewSetting
  Frm.HALT_opt_IsCovered = False
  '
  ' LOOK UP APPROPRIATE VALUE FOR cbo_Model_Type SCROLLBOX.
  '
Dim Ctl As Control
Dim New_Tag As Integer
Dim i As Integer
  Set Ctl = Frm.cbo_Model_Type
  Frm.HALT_cbo_Model_Type = True
  New_Tag = 0
  For i = 0 To Ctl.ListCount - 1
    With Temp_Plant.AerationBasin
      If (Ctl.ItemData(i) = .ModelingMechanism) Then
        New_Tag = i
        Exit For
      End If
    End With
  Next i
  Ctl.Tag = Trim$(Str$(New_Tag))
  Ctl.ListIndex = New_Tag
  Frm.HALT_cbo_Model_Type = False
  '
  ' ENABLE/DISABLE cmdNonuniformCSTRs BUTTON.
  '
Dim IsEnabled As Boolean
  With Temp_Plant.AerationBasin
    IsEnabled = (.CSTR.Count <> 1)
  End With
  Frm.cmdNonuniformCSTRs.Enabled = IsEnabled
  '
  ' ENABLE/DISABLE  TEXT BOXES.
  '
Dim IsLocked As Boolean
  With Temp_Plant.AerationBasin
    IsLocked = (.CSTR.Count <> 1)
  End With
  For i = 6 To 8
    Frm.txtData(i).BackColor = IIf(IsLocked, QBColor(7), QBColor(15))
    Frm.txtData(i).ForeColor = IIf(IsLocked, QBColor(8), QBColor(0))
    Frm.txtData(i).Locked = IsLocked
    Frm.lblData(i).Enabled = Not IsLocked
  Next i
End Sub


Sub frmD5A_CSTR_Repopulate_Values(Temp_Plant As TYPE_PlantDiagram)
Dim Frm As Form
Set Frm = frmD5A_CSTR
Dim i As Integer
  '
  ' MAIN BLOCK.
  '
  With Temp_Plant.AerationBasin
    Call unitsys_set_number_in_base_units(Frm.txtData(0), CDbl(.CSTR.Count))
    For i = 0 To 9
      If (i <> 9) Then
        Call unitsys_set_number_in_base_units(Frm.txtFeed(i), .CSTR.Feed(i))
        Call unitsys_set_number_in_base_units(Frm.txtVolume(i), .CSTR.Volume(i))
        Call unitsys_set_number_in_base_units(Frm.txtGasFlow(i), .CSTR.GasFlow(i))
        Call unitsys_set_number_in_base_units(Frm.txtBioMass(i), .CSTR.BioMass(i))
      Else
        Call unitsys_set_number_in_base_units(Frm.txtVolume(i), .Volume)
        Call unitsys_set_number_in_base_units(Frm.txtGasFlow(i), .GasFlow)
        Call unitsys_set_number_in_base_units(Frm.txtBioMass(i), .BioMass)
      End If
    Next i
  End With
End Sub
Sub frmD5A_CSTR_Refresh(Temp_Plant As TYPE_PlantDiagram)
Dim Frm As Form
Set Frm = frmD5A_CSTR
Dim LastRow As Integer
Dim i As Integer
Dim IsVisible_ShouldBe As Boolean
Dim IsVisible_Currently As Boolean
Dim NewSetting As Boolean
Dim Ctl As Control
Dim Ctl9 As Control
Dim IsEnabled As Boolean
Dim IsLocked As Boolean
Dim j As Integer
Dim StepVal As Double
Dim ThisSum As Double
  '
  ' SET LAST ROW.
  '
  LastRow = Temp_Plant.AerationBasin.CSTR.Count - 1
  '
  ' FOR FEED FRACTIONS:
  '     FOR STEP-FEED:
  '         FOR UNIFORM:
  '             SET CSTR(i)=1/N WHERE N IS THE NUMBER OF CSTRS.
  '         FOR NON-UNIFORM:
  '             ALLOW USER TO ENTER CSTR(1..N-1).
  '             CSTR(N) IS FORCED TO 1 - SUM(CSTR(1..N-1)).
  '     FOR NON-STEP-FEED:
  '         SET CSTR(1)=1 AND CSTR(i)=0 WHERE i ARE THE OTHER CSTRS.
  '
  With Temp_Plant.AerationBasin
    If (.CSTR.UseStepFeed) Then
      If (.CSTR.UniformFeed) Then
        StepVal = 1# / CDbl(.CSTR.Count)
        For i = 0 To 8
          .CSTR.Feed(i) = IIf(i <= LastRow, StepVal, 0#)
        Next i
      Else
        ThisSum = 0#
        For i = 0 To LastRow - 1
          ThisSum = ThisSum + .CSTR.Feed(i)
        Next i
        .CSTR.Feed(LastRow) = 1# - ThisSum
      End If
    Else
      .CSTR.Feed(0) = 1#
      For i = 1 To LastRow
        .CSTR.Feed(i) = 0#
      Next i
    End If
  End With
  '
  ' FOR VOLUME: SET UNIFORM VALUES FROM TOTAL, OR
  ' SET TOTAL FROM NON-UNIFORM VALUES.
  '
  With Temp_Plant.AerationBasin
    If (.CSTR.UniformVolume) Then
      StepVal = .Volume / CDbl(.CSTR.Count)
      For i = 0 To 8
        .CSTR.Volume(i) = IIf(i <= LastRow, StepVal, 0#)
      Next i
    Else
      .Volume = 0#
      For i = 0 To LastRow
        .Volume = .Volume + .CSTR.Volume(i)
      Next i
    End If
  End With
  '
  ' FOR GAS FLOW: SET UNIFORM VALUES FROM TOTAL, OR
  ' SET TOTAL FROM NON-UNIFORM VALUES.
  '
  With Temp_Plant.AerationBasin
    If (.CSTR.UniformGasFlow) Then
      StepVal = .GasFlow / CDbl(.CSTR.Count)
      For i = 0 To 8
        .CSTR.GasFlow(i) = IIf(i <= LastRow, StepVal, 0#)
      Next i
    Else
      .GasFlow = 0#
      For i = 0 To LastRow
        .GasFlow = .GasFlow + .CSTR.GasFlow(i)
      Next i
    End If
  End With
  '
  ' FOR BIOMASS CONCENTRATION: SET UNIFORM VALUES FROM AVERAGE, OR
  ' SET AVERAGE FROM NON-UNIFORM VALUES.
  '
  With Temp_Plant.AerationBasin
    If (.CSTR.UniformBioMass) Then
      'StepVal = .BioMass / CDbl(.CSTR.Count)
      'For i = 0 To 8
      '  .CSTR.BioMass(i) = IIf(i <= LastRow, StepVal, 0#)
      'Next i
      For i = 0 To 8
        .CSTR.BioMass(i) = IIf(i <= LastRow, .BioMass, 0#)
      Next i
    Else
      '.BioMass = 0#
      'For i = 0 To LastRow
      '  .BioMass = .BioMass + .CSTR.BioMass(i)
      'Next i
      .BioMass = 0#
      For i = 0 To LastRow
        .BioMass = .BioMass + _
            .CSTR.BioMass(i) * _
            .CSTR.Volume(i) / .Volume
      Next i
    End If
  End With
  '
  ' REPOPULATE VALUES INTO TEXT CONTROLS.
  '
  Call frmD5A_CSTR_Repopulate_Values(Temp_Plant)
  '
  ' SET VISIBILITY OF ROWS.
  '
  LastRow = Temp_Plant.AerationBasin.CSTR.Count - 1
  For i = 0 To 8
    IsVisible_ShouldBe = IIf(i <= LastRow, True, False)
    IsVisible_Currently = Frm.LeftLabel(i).Visible
    If (IsVisible_ShouldBe <> IsVisible_Currently) Then
      Frm.LeftLabel(i).Visible = IsVisible_ShouldBe
      Frm.txtFeed(i).Visible = IsVisible_ShouldBe
      Frm.txtVolume(i).Visible = IsVisible_ShouldBe
      Frm.txtGasFlow(i).Visible = IsVisible_ShouldBe
      Frm.txtBioMass(i).Visible = IsVisible_ShouldBe
    End If
  Next i
  '
  ' UPDATE LOCKED STATUS OF EACH COLUMN.
  '
  For i = 0 To 8
    For j = 1 To 4
      With Temp_Plant.AerationBasin.CSTR
        Select Case j
          Case 1:
            IsEnabled = IIf(.UseStepFeed, Not .UniformFeed, False)
            Set Ctl = Frm.txtFeed(i)
            Set Ctl9 = Nothing
          Case 2:
            IsEnabled = Not .UniformVolume
            Set Ctl = Frm.txtVolume(i)
            Set Ctl9 = Frm.txtVolume(9)
          Case 3:
            IsEnabled = Not .UniformGasFlow
            Set Ctl = Frm.txtGasFlow(i)
            Set Ctl9 = Frm.txtGasFlow(9)
          Case 4:
            IsEnabled = Not .UniformBioMass
            Set Ctl = Frm.txtBioMass(i)
            Set Ctl9 = Frm.txtBioMass(9)
        End Select
      End With
      IsLocked = Not IsEnabled
      Ctl.Locked = IsLocked
      Ctl.BackColor = IIf(IsLocked, QBColor(7), QBColor(15))
      Ctl.ForeColor = IIf(IsLocked, QBColor(8), QBColor(0))
      If (i = 0) And (j <> 1) Then
        IsLocked = Not IsLocked
        Ctl9.Locked = IsLocked
        Ctl9.BackColor = IIf(IsLocked, QBColor(7), QBColor(15))
        Ctl9.ForeColor = IIf(IsLocked, QBColor(8), QBColor(0))
      End If
    Next j
  Next i
  '
  ' SET APPROPRIATE VALUES FOR chkUniform(*) CONTROLS.
  '
  With Temp_Plant.AerationBasin.CSTR
    For i = 0 To 3
      Select Case i
        Case 0: NewSetting = .UniformFeed
        Case 1: NewSetting = .UniformVolume
        Case 2: NewSetting = .UniformGasFlow
        Case 3: NewSetting = .UniformBioMass
      End Select
      Frm.HALT_chkUniform = True
      Set Ctl = Frm.chkUniform(i)
      Ctl.Value = NewSetting
      Ctl.Tag = Trim$(Str$(CInt(NewSetting)))
      Frm.HALT_chkUniform = False
    Next i
  End With
  '
  ' SET APPROPRIATE VALUE FOR chkStepFeed CONTROL.
  '
  With Temp_Plant.AerationBasin.CSTR
    NewSetting = .UseStepFeed
    Frm.HALT_chkStepFeed = True
    Set Ctl = Frm.chkStepFeed
    Ctl.Value = NewSetting
    Ctl.Tag = Trim$(Str$(CInt(NewSetting)))
    Frm.HALT_chkStepFeed = False
  End With
  '
  ' SET VISIBILITY OF chkUniform(0) CONTROL.
  '
  With Temp_Plant.AerationBasin.CSTR
    IsVisible_ShouldBe = .UseStepFeed
  End With
  Frm.chkUniform(0).Visible = IsVisible_ShouldBe
  '
  ' DISABLE/ENABLE THE chkUniform(3) CONTROL.
  '
  With Temp_Plant.AerationBasin.CSTR
    IsEnabled = True
    If (.UseStepFeed = True) Then IsEnabled = False
    If (.UniformFeed = False) Then IsEnabled = False
    If (.UniformVolume = False) Then IsEnabled = False
    If (.UniformGasFlow = False) Then IsEnabled = False
    Frm.chkUniform(3).Enabled = IsEnabled
  End With
End Sub


Sub frmD5B_Biomass_Repopulate_Values(Temp_Plant As TYPE_PlantDiagram)
Dim Frm As Form
Set Frm = frmD5B_Biomass
  '
  ' MAIN BLOCK.
  '
  With Temp_Plant.AerationBasin.BioTreat
    Call unitsys_set_number_in_base_units(Frm.txtData(0), .MaxGrowthRate)
    Call unitsys_set_number_in_base_units(Frm.txtData(1), .HalfVelocityConst)
    Call unitsys_set_number_in_base_units(Frm.txtData(2), .BacterialDecay)
    Call unitsys_set_number_in_base_units(Frm.txtData(3), .YieldCoeff)
    Call unitsys_set_number_in_base_units(Frm.txtData(4), .BOD5Conc)
  End With
End Sub
Sub frmD5B_Biomass_Refresh(Temp_Plant As TYPE_PlantDiagram)
Dim Frm As Form
Set Frm = frmD5B_Biomass
Dim NewSetting As Boolean
Dim NewTag As Integer
  Call frmD5B_Biomass_Repopulate_Values(Temp_Plant)
End Sub


Sub frmD6_SecondaryClarifier_Repopulate_Values(Temp_Plant As TYPE_PlantDiagram)
Dim Frm As Form
Set Frm = frmD6_SecondaryClarifier
  '
  ' MAIN BLOCK.
  '
  With Temp_Plant.SecondaryClarifier
    Call unitsys_set_number_in_base_units(Frm.txtData(0), CDbl(.Count))
    Call unitsys_set_number_in_base_units(Frm.txtData(1), .VentilationRate)
    Call unitsys_set_number_in_base_units(Frm.txtData(2), .Depth)
    Call unitsys_set_number_in_base_units(Frm.txtData(3), .Volume)
    Call unitsys_set_number_in_base_units(Frm.txtData(4), .EffluentSolidsConc)
  End With
End Sub
Sub frmD6_SecondaryClarifier_Refresh(Temp_Plant As TYPE_PlantDiagram)
Dim Frm As Form
Set Frm = frmD6_SecondaryClarifier
Dim NewSetting As Boolean
Dim NewTag As Integer
Dim i As Integer
Dim j As Integer
Dim LookingFor As Integer
Dim Ctl As Control
Dim New_Tag As Integer
  Call frmD6_SecondaryClarifier_Repopulate_Values(Temp_Plant)
  '
  ' SELECT APPROPRIATE opt_IsCovered SETTING.
  '
  With Temp_Plant.SecondaryClarifier
    NewSetting = .IsCovered
  End With
  NewTag = IIf(NewSetting, 1, 0)
  Frm.HALT_opt_IsCovered = True
  Frm.opt_IsCovered(NewTag).Value = True
  Frm.opt_IsCovered(1 - NewTag).Value = False
  Frm.opt_IsCovered(0).Enabled = True
  Frm.opt_IsCovered(1).Enabled = True
  Frm.opt_IsCovered(0).Tag = Trim$(Str$(NewTag))
  'Frm.cmdSpecifyTimeVariableInfluentConc.Enabled = NewSetting
  Frm.lblData(1).Enabled = NewSetting
  Frm.txtData(1).Enabled = NewSetting
  Frm.cboUnits(1).Enabled = NewSetting
  Frm.HALT_opt_IsCovered = False
End Sub


