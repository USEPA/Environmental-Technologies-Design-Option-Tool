Attribute VB_Name = "Refresh"
Option Explicit




Const Refresh_declarations_end = True


Sub frmMain_Repopulate_Values()
Dim Frm As Form
Set Frm = frmMain
  Call unitsys_set_number_in_base_units(frmMain.txtData(0), NowProj.length)
  Call unitsys_set_number_in_base_units(frmMain.txtData(1), NowProj.Diameter)
  Call unitsys_set_number_in_base_units(frmMain.txtData(2), NowProj.Mass)
  Call unitsys_set_number_in_base_units(frmMain.txtData(3), NowProj.FlowRate)
End Sub
Sub frmMain_Refresh()
  Call frmMain_Repopulate_Values
End Sub


