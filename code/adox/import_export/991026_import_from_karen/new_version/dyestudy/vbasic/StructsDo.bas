Attribute VB_Name = "StructsDo"
Option Explicit



Const StructsDo_declarations_end = True


Sub Project_SetDefaults(ByRef p As Project_Type)
Dim i As Integer


  'DYESTUDY PARAMETERS
  p.dyestudy_output = ""
  ReDim p.DyeStudy(1 To 400)
  p.dyestudy_count = 400
  
  For i = 1 To 400
    p.DyeStudy(i).concentration = " "
    p.DyeStudy(i).time = " "
  Next i
  p.dyestudy_output = ""
  p.dyestudy_calcdate = ""
  p.Predicted_Available = False
  p.Predicted_count = 0
  p.Experimental_count = 0
  p.dirty = False
  Call Kill_If_It_Exists(App.Path & "\exes\outpt.txt")
  Call Kill_If_It_Exists(App.Path & "\exes\outpt2.txt")
  Call Kill_If_It_Exists(App.Path & "\exes\outpt3.txt")
  
End Sub


Sub DirtyFlag_Refresh(proj As Project_Type)
  If (proj.dirty) Then
    frmMain.sspanel_Dirty = ""
    frmMain.sspanel_Dirty.ForeColor = QBColor(4 + 8)
    frmMain.sspanel_Dirty = "Data Changed"
  Else
    frmMain.sspanel_Dirty = ""
    frmMain.sspanel_Dirty.ForeColor = QBColor(0)
    frmMain.sspanel_Dirty = "Data Unchanged"
  End If
End Sub


Sub DirtyFlag_Throw(ByRef proj As Project_Type)
  proj.dirty = True
  Call DirtyFlag_Refresh(proj)
  'ADDED THIS BECAUSE OF NEW PEC FORTRAN CODE
  FortranLink_fn_MainOutput = App.Path & "\exes\adox_out.txt"
  Call Kill_If_It_Exists(FortranLink_fn_MainOutput)
End Sub

