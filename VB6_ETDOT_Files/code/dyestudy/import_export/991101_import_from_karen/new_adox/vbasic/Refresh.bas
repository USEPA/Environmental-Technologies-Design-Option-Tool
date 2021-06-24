Attribute VB_Name = "Refresh"
Option Explicit
Sub frmMain_Repopulate_Values()
Dim Frm As Form
Set Frm = frmMain
  Call unitsys_set_number_in_base_units(Frm.txtdata(0), NowProj.volume)
  Call unitsys_set_number_in_base_units(Frm.txtdata(1), NowProj.tau)
  Call unitsys_set_number_in_base_units(Frm.txtdata(2), NowProj.opsize)
  Call unitsys_set_number_in_base_units(Frm.txtdata(3), NowProj.ttotal)
  Call unitsys_set_number_in_base_units(Frm.txtdata(4), NowProj.xntimes)
  Call unitsys_set_number_in_base_units(Frm.txtdata(5), NowProj.inf_h2o2)
  Call unitsys_set_number_in_base_units(Frm.txtdata(6), NowProj.ph0)
  Call unitsys_set_number_in_base_units(Frm.txtdata(7), NowProj.phosph)
  Call unitsys_set_number_in_base_units(Frm.txtdata(8), NowProj.ticarbn)
  Call unitsys_set_number_in_base_units(Frm.txtdata(9), NowProj.alk)
  Call unitsys_set_number_in_base_units(Frm.txtdata(10), CDbl(NowProj.num_tanks))
  
End Sub

Sub frmPhotoChem_Repopulate_Values()
Dim Frm As Form
Set Frm = frmPhotoChem
  Call unitsys_set_number_in_base_units(Frm.txtdata(0), CDbl(NowProj.lamp_power))
  Call unitsys_set_number_in_base_units(Frm.txtdata(2), CDbl(NowProj.uvpathl))
End Sub
Sub refresh_frmMain()
Dim new_tag As Integer
Dim i As Integer
Dim old_listindex As Integer

  Call frmMain_Repopulate_Values
  
  'REACTOR PROPERTIES.
  Select Case NowProj.idreact
    Case IDREACT_CMBR:
      new_tag = 0
      frmMain.lbldesc(1).Enabled = False
      frmMain.txtdata(1).Enabled = False
      frmMain.lbldesc2(1).Enabled = False
      frmMain.cboUnits(1).Enabled = False
      frmMain.txtdata(1).Text = "n/a"
      frmMain.lbldesc(3).Enabled = True
      frmMain.txtdata(3).Enabled = True
      frmMain.cboUnits(3).Enabled = True
'      frmMain.lbldesc2(3).Enabled = True
      frmMain.lbldesc(4).Enabled = False
      frmMain.txtdata(4).Enabled = False
      frmMain.txtdata(4).Text = "n/a"
      frmMain.lbldesc(10).Enabled = False
      frmMain.lbldesc2(10).Enabled = False
      frmMain.txtdata(10).Enabled = False
      frmMain.txtdata(10).Text = "n/a"
      'UPDATE LABELS FOR "Influent" AND "Initial".
      frmMain.lbldesc(5).Caption = "Initial H2O2:"
      frmMain.lbldesc(6).Caption = "Initial pH:"
      frmMain.lbldesc(7).Caption = "Initial Phosphate Conc.:"
      frmMain.lbldesc(8).Caption = "Initial TIC Concentration:"
      frmMain.lbldesc(9).Caption = "Initial Alkalinity:"
    Case IDREACT_CMFR:
      new_tag = 1
      frmMain.lbldesc(1).Enabled = True
      frmMain.txtdata(1).Enabled = True
      frmMain.lbldesc2(1).Enabled = True
      frmMain.cboUnits(1).Enabled = True
      frmMain.lbldesc(3).Enabled = False
      frmMain.txtdata(3).Enabled = False
      frmMain.lbldesc(3).Enabled = False
      frmMain.cboUnits(3).Enabled = False
      frmMain.lbldesc(4).Enabled = True
      frmMain.txtdata(4).Enabled = True
      frmMain.lbldesc(10).Enabled = True
      frmMain.lbldesc2(10).Enabled = True
      frmMain.txtdata(10).Enabled = True
      'UPDATE LABELS FOR "Influent" AND "Initial".
      frmMain.lbldesc(5).Caption = "Influent H2O2:"
      frmMain.lbldesc(6).Caption = "Influent pH:"
      frmMain.lbldesc(7).Caption = "Influent Phosphate Conc.:"
      frmMain.lbldesc(8).Caption = "Influent TIC Concentration:"
      frmMain.lbldesc(9).Caption = "Influent Alkalinity:"
  End Select
  Call AssignTag_Scrollbox(frmMain.cboReactorType, new_tag)
  frmMain.cboReactorType.ListIndex = new_tag
  
  Call AssignTextAndTag_WithRange(frmMain.txtdata(0), NowProj.volume, 1E-20, 1E+20)
  If (NowProj.idreact = IDREACT_CMFR) Then
    Call AssignTextAndTag_WithRange(frmMain.txtdata(1), NowProj.tau, 1E-20, 1E+20)
  End If
  Call AssignTextAndTag_WithRange(frmMain.txtdata(5), NowProj.inf_h2o2, 1E-20, 1E+20)

  'NUMERICAL SIMULATION PARAMETERS.
  Call AssignTextAndTag_WithRange(frmMain.txtdata(2), NowProj.ssize / 60#, 1E-20, 1E+20)
  Call AssignTextAndTag_WithRange(frmMain.txtdata(3), NowProj.ttotal, 1E-20, 1E+20)
  If (NowProj.idreact = IDREACT_CMFR) Then
    Call AssignTextAndTag_WithRange(frmMain.txtdata(4), NowProj.xntimes, 1E-20, 1E+20)
    Call AssignTextAndTag_WithRange(frmMain.txtdata(10), NowProj.num_tanks, 1#, 1E+20)
  End If

  'WATER QUALITY PROPERTIES.
  Call AssignTextAndTag_WithRange(frmMain.txtdata(6), NowProj.ph0, 1E-20, 1E+20)
  Call AssignTextAndTag_WithRange(frmMain.txtdata(7), NowProj.phosph, 1E-20, 1E+20)
  Call AssignTextAndTag_WithRange(frmMain.txtdata(8), NowProj.ticarbn, 1E-20, 1E+20)
  Call AssignTextAndTag_WithRange(frmMain.txtdata(9), NowProj.alk, 1E-20, 1E+20)
  
  frmMain.ssframe_alk.top = frmMain.ssframe_tic.top
  frmMain.ssframe_alk.left = frmMain.ssframe_tic.left
  Select Case NowProj.idcarbn
    Case IDCARBN_TIC:
      new_tag = 0
      frmMain.ssframe_alk.visible = False
      frmMain.ssframe_tic.visible = True
    Case IDCARBN_ALKALINITY:
      new_tag = 1
      frmMain.ssframe_alk.visible = True
      frmMain.ssframe_tic.visible = False
  End Select
  Call AssignTag_OptionChecks(frmMain.optTICInput(0), new_tag)
  frmMain.optTICInput(CInt(new_tag)).Value = True

  'TARGET COMPOUNDS.
  old_listindex = frmMain.cboTarget.ListIndex
  frmMain.cboTarget.Clear
  For i = 1 To NowProj.TargetCompounds_Count
    frmMain.cboTarget.AddItem Trim$(NowProj.TargetCompounds(i).comname)
  Next i
  If (old_listindex < 0) And (frmMain.cboTarget.ListCount >= 1) Then
    old_listindex = 0
  End If
  If (old_listindex > frmMain.cboTarget.ListCount - 1) Then
    old_listindex = frmMain.cboTarget.ListCount - 1
  End If
  frmMain.cboTarget.ListIndex = old_listindex
  
  'PHOTOCHEMICAL PARAMETERS.


End Sub


Sub refresh_frmTarget(tc As TargetCompound_Type)
Dim is_nom As Integer
Dim i As Integer
Dim j As Integer

  'PROPERTIES OF PROTONATED FORM.
  Call AssignTextAndTag_WithRange(frmTarget.txtdata(0), tc.concini, 0#, 1E+20)
  Call AssignTextAndTag_WithRange(frmTarget.txtdata(1), tc.val, 1E-20, 1E+20)
  Call AssignTextAndTag_WithRange(frmTarget.txtdata(2), tc.mw, 1E-20, 1E+20)
  Call AssignTextAndTag_WithRange(frmTarget.txtdata(3), tc.ncarbn, 1E-20, 1E+20)
  Call AssignTextAndTag_WithRange(frmTarget.txtdata(4), tc.nsubstt, 1E-20, 1E+20)
  Call AssignTextAndTag_WithRange(frmTarget.txtdata(5), tc.xk, 1E-20, 1E+20)
  Select Case NowProj.idreact
    Case IDREACT_CMBR:
      frmTarget.lbldesc(0).Caption = "Initial Concentration"
    Case IDREACT_CMFR:
      frmTarget.lbldesc(0).Caption = "Influent Concentration"
  End Select

  'EQUILBRIUM REACTION.
  Call AssignTextAndTag_WithRange(frmTarget.txtdata(6), tc.dep_xke, 1E-20, 1E+20)
  
  'PROPERTIES OF DEPROTONATED FORM.
  Call AssignTextAndTag_WithRange(frmTarget.txtdata(7), tc.dep_val, 1E-20, 1E+20)
  Call AssignTextAndTag_WithRange(frmTarget.txtdata(8), tc.dep_mw, 1E-20, 1E+20)
  Call AssignTextAndTag_WithRange(frmTarget.txtdata(9), tc.dep_xk, 1E-20, 1E+20)

  'OTHER RATE CONSTANTS.
  Call AssignTextAndTag_WithRange(frmTarget.txtdata(10), tc.xk_co3XM, 1E-20, 1E+20)
  Call AssignTextAndTag_WithRange(frmTarget.txtdata(11), tc.xk_hpo4XM, 1E-20, 1E+20)
  Call AssignTextAndTag_WithRange(frmTarget.txtdata(12), tc.xk_o2XM, 1E-20, 1E+20)
  Call AssignTextAndTag_WithRange(frmTarget.txtdata(13), tc.xk_ho2X, 1E-20, 1E+20)

  'IF THIS IS THE "NOM" COMPOUND, DISABLE LOTS OF TEXTBOXES.
  is_nom = False
  If (UCase$(Trim$(tc.comname)) = UCase$(Trim$("NOM"))) Then
    is_nom = True
  End If
  'keep: 0,2,5; disable: 1,3,4,6,7,8,9,10,11,12,13
  For i = 1 To 11
    j = Choose(i, 1, 3, 4, 6, 7, 8, 9, 10, 11, 12, 13)
    frmTarget.lbldesc(j).Enabled = Not is_nom
    frmTarget.txtdata(j).Enabled = Not is_nom
    If j <> 6 Then
      frmTarget.lbldesc2(j).Enabled = Not is_nom
    End If
    If (is_nom) Then
      frmTarget.txtdata(j).Text = "n/a"
    End If
  Next i
  If (is_nom) Then
    frmTarget.lbldesc2(0).Caption = "mg/L"
    frmTarget.lbldesc2(5).Caption = "1/(mg/L)-s"
  Else
    frmTarget.lbldesc2(0).Caption = "gmol/L"
    frmTarget.lbldesc2(5).Caption = "L/gmol-s"
  End If
  
End Sub


Sub refresh_frmPhotoChem(proj As Project_Type)
Dim i As Integer
Dim old_listindex As Integer
Dim idx_target As Integer
Dim new_tag As Integer
Dim temp As String

  Call frmPhotoChem_Repopulate_Values
  
  'WAVELENGTHS.
  frmPhotoChem.f1book_wavelen.visible = False
  frmPhotoChem.f1book_wavelen.MaxRow = proj.Wavelength_Count
  For i = 1 To proj.Wavelength_Count
    frmPhotoChem.f1book_wavelen.EntryRC(i, 1) = _
        Trim$(Str$(proj.Wavelengths(i).lwave))
    frmPhotoChem.f1book_wavelen.EntryRC(i, 2) = _
        Trim$(Str$(proj.Wavelengths(i).uvi))
  Next i
  frmPhotoChem.f1book_wavelen.visible = True
  
  'VALUES FOR EACH COMPOUND.
  '.... COMPOUND SCROLLBOX.
  old_listindex = frmPhotoChem.cboTarget.ListIndex
  frmPhotoChem.cboTarget.Clear
  For i = 1 To proj.TargetCompounds_Count
    frmPhotoChem.cboTarget.AddItem Trim$(proj.TargetCompounds(i).comname)
  Next i
  frmPhotoChem.cboTarget.AddItem "H2O2"
  If (old_listindex < 0) And (frmPhotoChem.cboTarget.ListCount >= 1) Then
    old_listindex = 0
  End If
  If (old_listindex > frmPhotoChem.cboTarget.ListCount - 1) Then
    old_listindex = frmPhotoChem.cboTarget.ListCount - 1
  End If
  Call AssignTag_Scrollbox(frmPhotoChem.cboTarget, old_listindex)
  frmPhotoChem.cboTarget.ListIndex = old_listindex

  '.... GRID.
  idx_target = old_listindex + 1
  If (idx_target > NowProj.TargetCompounds_Count) Then
    'THIS IS THE H2O2 COMPOUND.
    idx_target = -1
  End If
  frmPhotoChem.f1book_vals.visible = False
  frmPhotoChem.f1book_vals.MaxRow = proj.Wavelength_Count
  If (idx_target = -1) Then
    'DISPLAY H2O2 DATA.
    For i = 1 To proj.Wavelength_Count
      frmPhotoChem.f1book_vals.EntryRC(i, 1) = _
          Trim$(Str$(proj.Wavelengths(i).lwave))
      frmPhotoChem.f1book_vals.EntryRC(i, 2) = _
          Trim$(Str$(proj.extcoef_h2o2(i)))
      frmPhotoChem.f1book_vals.EntryRC(i, 3) = _
          Trim$(Str$(proj.quatyd_h2o2(i)))
    Next i
  Else
    'DISPLAY NON-H2O2 DATA.
    For i = 1 To proj.Wavelength_Count
      frmPhotoChem.f1book_vals.EntryRC(i, 1) = _
          Trim$(Str$(proj.Wavelengths(i).lwave))
      frmPhotoChem.f1book_vals.EntryRC(i, 2) = _
          Trim$(Str$(proj.extcoef(idx_target, i)))
      frmPhotoChem.f1book_vals.EntryRC(i, 3) = _
          Trim$(Str$(proj.quatyd(idx_target, i)))
    Next i
  End If
  frmPhotoChem.f1book_vals.visible = True
  
  'LAMP PARAMETERS FRAME.
  Call AssignTextAndTag_WithRange(frmPhotoChem.txtdata(0), proj.lamp_power, 1E-20, 1E+20)
  Call AssignTextAndTag_WithRange(frmPhotoChem.txtdata(2), proj.uvpathl, 1E-20, 1E+20)
  frmPhotoChem.txtDataStr(0).Text = Trim$(proj.lamp_name)
  frmPhotoChem.txtDataStr(0).Tag = Trim$(proj.lamp_name)
  
  'TRACK DOWN SPECIFICATION SCROLLBOX VALUE.
  new_tag = 0
  For i = 0 To frmPhotoChem.cboLightSpecMethod.ListCount - 1
    If (frmPhotoChem.cboLightSpecMethod.ItemData(i) = proj.iduvi) Then
      new_tag = i
      Exit For
    End If
  Next i
  Call AssignTag_Scrollbox(frmPhotoChem.cboLightSpecMethod, new_tag)
  frmPhotoChem.cboLightSpecMethod.ListIndex = new_tag
  Select Case proj.iduvi
    Case IDUVI_EINSTEINS_L_S: temp = "UV Light" & vbCrLf & "Intensity (Einsteins/L-s)"
    Case IDUVI_WATTS: temp = "UV Light" & vbCrLf & "Intensity (watts)"
    Case IDUVI_EFFICIENCY: temp = vbCrLf & "Efficiency" & vbCrLf & "(dim'less)"
  End Select
  frmPhotoChem.f1book_wavelen.ColText(2) = temp

End Sub

'''Sub refresh_frmDyeStudy(proj As Project_Type)
'''Dim i As Integer
'''
'''
'''  frmDyeStudy.f1book_dyestudy.visible = False
'''  frmDyeStudy.f1book_dyestudy.MaxRow = proj.dyestudy_count
'''  For i = 1 To proj.dyestudy_count
'''    frmDyeStudy.f1book_dyestudy.EntryRC(i, 1) = _
'''        Trim$(Str$(proj.DyeStudy(i).time))
'''    frmDyeStudy.f1book_dyestudy.EntryRC(i, 2) = _
'''        Trim$(Str$(proj.DyeStudy(i).concentration))
'''  Next i
'''  frmDyeStudy.f1book_dyestudy.visible = True
'''  Call AssignTextAndTag(frmDyeStudy.txtData(0), proj.dyestudy_calcdate)
'''
'''End Sub

