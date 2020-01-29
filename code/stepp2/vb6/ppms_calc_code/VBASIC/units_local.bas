Attribute VB_Name = "UNITS_LOCAL"

''COMMUNICATIONS WITH frmBed.
'Global frmBed_Copy_Of_CURRENT_BEDCOMPONENT As Integer
'Global frmBed_Copy_Of_TempBedDef As BedDefinition_Type
'Global frmBed_Copy_Of_TempProj As Project_Type
'      'NOTE: THIS RECORD IS USED TO OBTAIN THE VALUES
'      'OF THE FOLLOWING VARIABLES (CURRENTLY DISPLAYED COMPONENT):
'      '    - molecular weight
'      '    - Freundlich 1/n





Const UNITS_LOCAL_declarations_end = 0


Function local_unitsys_convert_getfactor_FreundlichK( _
    UnitType As String, _
    factor1 As Double, _
    factor2 As Double) As Double
Dim X As Double
  Select Case Trim$(UCase$(UnitType))
    Case "(MG/G)*(L/MG)^(1/N)":
      X = 1#
    Case "(MMOL/G)*(L/MMOL)^(1/N)":
      X = 1# / factor1
    Case "(µG/G)*(L/µG)^(1/N)":
      X = 1# / factor2
    Case "(µMOL/G)*(L/µMOL)^(1/N)":
      X = 1# / factor1 / factor2
  End Select
  local_unitsys_convert_getfactor_FreundlichK = X
End Function


Sub local_unitsys_convert( _
    UnitType As String, _
    unit_from As String, _
    unit_to As String, _
    val_from As Double, _
    val_to As Double)
'Dim now_MW As Double
'Dim now_OneOverN As Double
'Dim factor1 As Double
'Dim factor2 As Double
'Dim factor_from As Double
'Dim factor_to As Double
'  unittype = Trim$(UCase$(unittype))
'  If (unittype = Trim$(UCase$("freundlich_k"))) Then
'    'now_MW = frmBed_Copy_Of_TempProj.Components( _
'    '    frmBed_Copy_Of_CURRENT_BEDCOMPONENT).xwt
'    'now_OneOverN = frmBed_Copy_Of_TempBedDef.BedComponents( _
'    '    frmBed_Copy_Of_CURRENT_BEDCOMPONENT).XN
'    now_MW = Component(0).MW
'    now_OneOverN = Component(0).Use_OneOverN
'    factor1 = (now_MW) ^ (now_OneOverN - 1#)
'    factor2 = (1000#) ^ (1# - now_OneOverN)
'    factor_from = local_unitsys_convert_getfactor_FreundlichK(unit_from, factor1, factor2)
'    factor_to = local_unitsys_convert_getfactor_FreundlichK(unit_to, factor1, factor2)
'    'PERFORM THE CONVERSION.
'    val_to = val_from / factor_to * factor_from
'  End If
End Sub
Sub local_unitsys_populate_units( _
    cbc As Control, _
    UnitType As String)
  UnitType = Trim$(UCase$(UnitType))
  If (UnitType = Trim$(UCase$("freundlich_k"))) Then
    cbc.AddItem "(mg/g)*(L/mg)^(1/n)"
    cbc.AddItem "(mmol/g)*(L/mmol)^(1/n)"
    cbc.AddItem "(µg/g)*(L/µg)^(1/n)"
    cbc.AddItem "(µmol/g)*(L/µmol)^(1/n)"
  End If
End Sub
'Sub local_unitsys_populate_units(H As Integer, UnitType As String)
'  UnitType = Trim$(UCase$(UnitType))
'  If (UnitType = Trim$(UCase$("freundlich_k"))) Then
'    unitsys(H).CboX.AddItem "(mg/g)*(L/mg)^(1/n)"
'    unitsys(H).CboX.AddItem "(mmol/g)*(L/mmol)^(1/n)"
'    unitsys(H).CboX.AddItem "(µg/g)*(L/µg)^(1/n)"
'    unitsys(H).CboX.AddItem "(µmol/g)*(L/µmol)^(1/n)"
'  End If
'End Sub


