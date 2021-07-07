Attribute VB_Name = "StructsDo"
Option Explicit




Global Const EXTCOEF_DEFAULT_VALUE = 0#
Global Const QUATYD_DEFAULT_VALUE = 0#
Global Const UVI_DEFAULT_VALUE = 34.3  '0.0000041

Global Const EXTCOEF_H2O2_DEFAULT_VALUE = 19#
Global Const QUATYD_H2O2_DEFAULT_VALUE = 0.5


']']
'"1.0.3","Version"
'"Advanced Oxidation Process Software","App.Title"
'"7/1/98 10:14:43 PM","Save date/time"



']']



Const StructsDo_declarations_end = 0


Sub Project_New(proj As Project_Type)
  Call Project_SetDefaults(proj)
  NowProj.dirty = False
  Call DirtyFlag_Refresh(proj)
End Sub


Sub Project_SetDefaults(ByRef p As Project_Type)

  p.Filename = ""
  p.dirty = False
  
  'REACTOR PROPERTIES.
  p.idreact = IDREACT_CMFR
  p.volume = 70#
  p.tau = 25
  p.num_tanks = 2
  p.iduvi = 1

  'NUMERICAL SIMULATION PARAMETERS.
  p.ssize = 300#
  p.ttotal = 75#     '50#
  p.opsize = 5#
  p.xntimes = 3#
  
  'WATER QUALITY PROPERTIES.
  p.ph0 = 8.4
  p.phosph = 0#
  p.idcarbn = IDCARBN_TIC
  p.alk = 250#
  p.ticarbn = 0.004
  p.inf_h2o2 = 0.001
  
  'TARGET COMPOUNDS.
  p.TargetCompounds_Count = 2
  ReDim p.TargetCompounds(1 To 2)
  
  '...... SET DEFAULTS FOR NOM.
  Call TargetCompound_SetDefaults(p.TargetCompounds(1))
  p.TargetCompounds(1).comname = "NOM"
  p.TargetCompounds(1).concini = 0   'mg/L
  p.TargetCompounds(1).val = 0
  p.TargetCompounds(1).mw = 200#
  p.TargetCompounds(1).ncarbn = 3
  p.TargetCompounds(1).nsubstt = 1
  p.TargetCompounds(1).xk = 20000#

  '...... SET DEFAULTS FOR R1.
'  Call TargetCompound_SetDefaults(p.TargetCompounds(2))
  p.TargetCompounds(1).dep_comname = "R1-"
  p.TargetCompounds(1).dep_val = -1
  p.TargetCompounds(1).dep_mw = 99
  p.TargetCompounds(1).dep_xk = 3000000000#
  p.TargetCompounds(1).dep_xke = 11.6
  p.TargetCompounds(1).xk_co3XM = 0
  p.TargetCompounds(1).xk_hpo4XM = 0
  p.TargetCompounds(1).xk_o2XM = 0
  p.TargetCompounds(1).xk_ho2X = 0
  
  '..... SET DEFAULTS FOR DBCP
   Call TargetCompound_SetDefaults(p.TargetCompounds(2))
  p.TargetCompounds(2).comname = "DBCP"
  p.TargetCompounds(2).concini = 0.00000183
  p.TargetCompounds(2).val = 0
  p.TargetCompounds(2).mw = 100
  p.TargetCompounds(2).ncarbn = 3
  p.TargetCompounds(2).nsubstt = 3
  p.TargetCompounds(2).xk = 150000000
  p.TargetCompounds(2).dep_comname = "R1-"
  p.TargetCompounds(2).dep_val = -1
  p.TargetCompounds(2).dep_mw = 99
  p.TargetCompounds(2).dep_xk = 3000000000#
  p.TargetCompounds(2).dep_xke = 14
  p.TargetCompounds(2).xk_co3XM = 0
  p.TargetCompounds(2).xk_hpo4XM = 0
  p.TargetCompounds(2).xk_o2XM = 0
  p.TargetCompounds(2).xk_ho2X = 0
  
  'PHOTOCHEMICAL PARAMETERS.
  p.Wavelength_Count = 1
  ReDim p.Wavelengths(1 To 1)
  p.Wavelengths(1).lwave = 253.7
  p.Wavelengths(1).uvi = UVI_DEFAULT_VALUE  '0.0000041
  ReDim p.extcoef(1 To 2, 1 To 1)
  p.extcoef(1, 1) = 0.0867  '0#
  p.extcoef(2, 1) = 14.7
  ReDim p.quatyd(1 To 2, 1 To 1)
  p.quatyd(1, 1) = 0   '0#
  p.quatyd(2, 1) = 0.47
  p.uvpathl = 6.8
  ReDim p.extcoef_h2o2(1 To 1)
  p.extcoef_h2o2(1) = 19  '19#
  ReDim p.quatyd_h2o2(1 To 1)
  p.quatyd_h2o2(1) = 0.5  '0.5
  'p.lamp_eff = 35#
  p.lamp_power = 500#
  p.lamp_name = "Low Pressure UV Lamp"
  
End Sub


Sub DirtyFlag_Refresh(proj As Project_Type)
  If (proj.dirty) Then
    frmMain.panDirty = ""
    frmMain.panDirty.ForeColor = QBColor(4 + 8)
    frmMain.panDirty = "Data Changed"
  Else
    frmMain.panDirty = ""
    frmMain.panDirty.ForeColor = QBColor(0)
    frmMain.panDirty = "Data Unchanged"
  End If
End Sub


Sub DirtyFlag_Throw(ByRef proj As Project_Type)
  proj.dirty = True
  Call DirtyFlag_Refresh(proj)
  'DELETE EXISTING RESULTS, IF ANY.
  Call FortranLink_SetFilenames
  Call Kill_If_It_Exists(FortranLink_fn_MainOutput)
End Sub



Sub TargetCompound_SetDefaults(ByRef tc As TargetCompound_Type)
  'tc.comname = ""        'not set here!
  tc.concini = 0.000008
  tc.val = 0
  tc.mw = 100#
  tc.ncarbn = 3
  tc.nsubstt = 1
  tc.xk = 30000000#
  tc.dep_comname = "R1-"
  tc.dep_val = -1#
  tc.dep_mw = tc.mw - 1#
  tc.dep_xk = 3000000000#
  tc.dep_xke = 11.6
  tc.xk_co3XM = 0#
  tc.xk_hpo4XM = 0#
  tc.xk_o2XM = 0#
  tc.xk_ho2X = 0#
End Sub


Function TargetCompound_GetIndex(proj As Project_Type, recname As String) As Integer
Dim found As Integer
Dim i As Integer

  found = False
  For i = 1 To proj.TargetCompounds_Count
    If (Trim$(UCase$(proj.TargetCompounds(i).comname)) = Trim$(UCase$(recname))) Then
      found = True
      Exit For
    End If
  Next i
  If (found) Then
    TargetCompound_GetIndex = i
  Else
    TargetCompound_GetIndex = 0
  End If

End Function


Function TargetCompound_IsKeyExist(proj As Project_Type, recname As String) As Integer
Dim retval As Integer

  retval = TargetCompound_GetIndex(proj, recname)
  If (retval = 0) Then
    TargetCompound_IsKeyExist = False
  Else
    TargetCompound_IsKeyExist = True
  End If

End Function



