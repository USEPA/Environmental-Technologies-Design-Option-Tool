Attribute VB_Name = "FileIO_Legacy"
Option Explicit




Const FileIO_Legacy_declarations_end = True


'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////     VERSION 1.00     ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Sub File_Open_Legacy_v1_00(f As Integer)
Dim i As Integer, j As Integer
Dim VersionST As String, msg As String
Dim PSDR As Double
'---Declaration for tempo variables
ReDim State_Check_WaterT(2) As Integer
Dim SPDFR_Low_Concentration_tempo As Integer
Dim Use_SPDFR_Correlation_Tempo As Integer
Dim Use_Tortuosity_Correlation_Tempo As Integer
Dim Constant_Tortuosity_Tempo As Integer
Dim PSDRTempo As Double
Dim Number_ComponentTempo As Integer
Dim BedTempo As BedPropertyType
ReDim CompoTempo(Number_Compo_Max) As ComponentPropertyType
Dim CarbonTempo As CarbonPropertyType
Dim MCTempo As Integer, NCTempo As Integer
Dim TimePTempo As Para_Int
Dim Number_Influent_PointsTempo As Integer
ReDim C_InfluentTempo(Number_Compo_Max, Number_Max_Influent_Points) As Double
ReDim T_InfluentTempo(Number_Max_Influent_Points) As Double
Dim NVersion_Temp As Double
Dim temp As String
ReDim u(5) As String
Dim TempName As String
'
'END OF DECLARATION.
'
  msg = ""
'-----------------------------------
'-----   Version 1.00 ---------
'-----------------------------------
  If batchrun <> 1 Then
   'temp = "Note--this is an Adsorption Simulation Software version 1.00 file."
   'temp = temp & NL & "If you save this file, it will be saved in version " & Format$(NVersion, "0.00") & " format."
   'MsgBox temp, MB_ICONEXCLAMATION, AppName_For_Display_long
  End If
  Input #f, Number_ComponentTempo
  For i = 1 To Number_ComponentTempo
    Call SetComponentDefaults(CompoTempo(i), -1)
    Input #f, CompoTempo(i).Name, CompoTempo(i).MW, CompoTempo(i).InitialConcentration, CompoTempo(i).MolarVolume, CompoTempo(i).BP, CompoTempo(i).Use_K, CompoTempo(i).Use_OneOverN
    CompoTempo(i).UserEntered_K = CompoTempo(i).Use_K
    CompoTempo(i).UserEntered_OneOverN = CompoTempo(i).Use_OneOverN
    Input #f, CompoTempo(i).SPDFR, CompoTempo(i).SPDFR_Low_Concentration, CompoTempo(i).Use_SPDFR_Correlation
    Input #f, CompoTempo(i).kf, CompoTempo(i).Ds, CompoTempo(i).Dp, CompoTempo(i).Corr(1), CompoTempo(i).Corr(2), CompoTempo(i).Corr(3)
    Input #f, CompoTempo(i).KP_User_Input(1), CompoTempo(i).KP_User_Input(2), CompoTempo(i).KP_User_Input(3)
    Input #f, CompoTempo(i).K_Reduction, CompoTempo(i).Correlation.Name, CompoTempo(i).Correlation.Coeff(1), CompoTempo(i).Correlation.Coeff(2)
  Next i
  Input #f, BedTempo.length, BedTempo.Diameter, BedTempo.Weight, BedTempo.Flowrate, BedTempo.WaterDensity, BedTempo.WaterViscosity, BedTempo.Temperature, BedTempo.Pressure, BedTempo.Phase, BedTempo.NumberOfBeds
  Input #f, BedTempo.Water_Correlation.Name, BedTempo.Water_Correlation.Coeff(1), BedTempo.Water_Correlation.Coeff(2), BedTempo.Water_Correlation.Coeff(3), BedTempo.Water_Correlation.Coeff(4)
  Input #f, CarbonTempo.Name, CarbonTempo.Porosity, CarbonTempo.Density, CarbonTempo.ParticleRadius, CarbonTempo.Tortuosity
  Input #f, State_Check_WaterT(1), State_Check_WaterT(2)
  Input #f, Use_Tortuosity_Correlation_Tempo, Constant_Tortuosity_Tempo
  Input #f, NCTempo, MCTempo
  Input #f, TimePTempo.Init, TimePTempo.End, TimePTempo.np, TimePTempo.Step
  Input #f, Number_Influent_PointsTempo
  If (Number_Influent_PointsTempo > 0) Then
    For i = 1 To Number_Influent_PointsTempo
      Input #f, T_InfluentTempo(i)
      For j = 1 To Number_ComponentTempo
       Input #f, C_InfluentTempo(j, i)
      Next j
    Next i
  End If
  On Error GoTo No_Data_Points
  Input #f, NData_Points, Number_Component
  On Error GoTo 0
  CarbonTempo.ShapeFactor = 1#
  FileNote = ""
Read_Data_Points_v1_00:
  If (NData_Points > 0 And Number_Component > 0) Then
    For i = 1 To NData_Points
      Input #f, T_Data_Points(i)
      For j = 1 To Number_Component
        Input #f, C_Data_Points(j, i)
      Next j
    Next i
  End If
Update_Tortuosities_from_v1_00:
  For i = 1 To Number_ComponentTempo
    CompoTempo(i).Tortuosity = CarbonTempo.Tortuosity
    CompoTempo(i).Use_Tortuosity_Correlation = Use_Tortuosity_Correlation_Tempo
    CompoTempo(i).Constant_Tortuosity = Constant_Tortuosity_Tempo
  Next i
  'IMPORT ALL TEMPORARY DATA INTO GLOBAL VARIABLES.
  'Update_All_Data
  'Store all the read variables in global variables
  If (Number_ComponentTempo > 0) Then
    Component_Number_Selected = 1
  Else
    Component_Number_Selected = -1
  End If
  Number_Component = Number_ComponentTempo
  State_Check_Water(2) = State_Check_WaterT(2)
  State_Check_Water(1) = State_Check_WaterT(1)
  'Use_SPDFR_Correlation = Use_SPDFR_Correlation_Tempo
  'SPDFR_Low_Concentration = SPDFR_Low_Concentration_tempo
'  Use_Tortuosity_Correlation = Use_Tortuosity_Correlation_Tempo
'  Constant_Tortuosity = Constant_Tortuosity_Tempo
  For i = 1 To Number_Component
    Component(i) = CompoTempo(i)
  Next i
  '========== What the heck does this code do!?!?
  '- Apparently all it does is cause problems.
  '- Commented out by ejo, 4/16/96. =================================
'  If (Mid$(VersionST, 1, 45) <> ("Input file for ASS - Windows - Version 0.80")) And (Mid$(VersionST, 1, 45) <> ("Input file for ASS - Windows - Version 1.00") And NVersion_Temp <> 1#) Then
'    If (Mid$(VersionST, 1, 45) <> ("Input file for ASS - Windows - Version 0.80")) And (Mid$(VersionST, 1, 45) <> ("Input file for ASS - Windows - Version 1.20") And NVersion_Temp <> 1.2) Then
'      If (Mid$(VersionST, 1, 45) <> ("Input file for ASS - Windows - Version 0.80")) And (Mid$(VersionST, 1, 45) <> ("Input file for ASS - Windows - Version 1.30") And NVersion_Temp <> 1.3) Then
  'If (NVersion_Temp <> 1#) Then
  '  If (NVersion_Temp <> 1.2) Then
  '    If (NVersion_Temp <> 1.3) Then
  '      For i = 1 To Number_Component
  '        Component(i).SPDFR = PSDRTempo
  '        Component(i).SPDFR_Low_Concentration = SPDFR_Low_Concentration_tempo
  '        Component(i).Use_SPDFR_Correlation = Use_SPDFR_Correlation_Tempo
  '      Next i
  '    End If
  '  End If
  'End If
  Carbon = CarbonTempo
  Bed = BedTempo
  TimeP = TimePTempo
  NC = NCTempo
  MC = MCTempo
  TimeP = TimePTempo
  Number_Influent_Points = Number_Influent_PointsTempo
  If (Number_Influent_Points > 0) Then
    For i = 1 To Number_Influent_Points
      T_Influent(i) = T_InfluentTempo(i)
      For j = 1 To Number_Component
        C_Influent(j, i) = C_InfluentTempo(j, i)
      Next j
    Next i
  End If
  Exit Sub
No_Data_Points:
  NData_Points = 0
  Number_Component = 0
  Resume Read_Data_Points_v1_00
  'Select Case NVersion_Temp
  '  Case 1#
  '    Resume Read_Data_Points_v1_00
  '  '     .
  '  '     .
  '  '     .
  '  '.... other versions ...
  '  '     .
  '  '     .
  '  '     .
  'End Select
  'Resume Read_Data_Points_v_LATEST
End Sub


'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////     VERSION 1.20     ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Sub File_Open_Legacy_v1_20(f As Integer)
Dim i As Integer, j As Integer
Dim VersionST As String, msg As String
Dim PSDR As Double
'---Declaration for tempo variables
ReDim State_Check_WaterT(2) As Integer
Dim SPDFR_Low_Concentration_tempo As Integer
Dim Use_SPDFR_Correlation_Tempo As Integer
Dim Use_Tortuosity_Correlation_Tempo As Integer
Dim Constant_Tortuosity_Tempo As Integer
Dim PSDRTempo As Double
Dim Number_ComponentTempo As Integer
Dim BedTempo As BedPropertyType
ReDim CompoTempo(Number_Compo_Max) As ComponentPropertyType
Dim CarbonTempo As CarbonPropertyType
Dim MCTempo As Integer, NCTempo As Integer
Dim TimePTempo As Para_Int
Dim Number_Influent_PointsTempo As Integer
ReDim C_InfluentTempo(Number_Compo_Max, Number_Max_Influent_Points) As Double
ReDim T_InfluentTempo(Number_Max_Influent_Points) As Double
Dim NVersion_Temp As Double
Dim temp As String
ReDim u(5) As String
Dim TempName As String
'
'END OF DECLARATION.
'
  msg = ""
'-----------------------------------
'-----   Version 1.20 ---------
'-----------------------------------
  If batchrun <> 1 Then
   'temp = "Note--this is an Adsorption Simulation Software version 1.20 file."
   'temp = temp & NL & "If you save this file, it will be saved in version " & Format$(NVersion, "0.00") & " format."
   'MsgBox temp, MB_ICONEXCLAMATION, AppName_For_Display_long
  End If
  Input #f, Number_ComponentTempo
  For i = 1 To Number_ComponentTempo
    Call SetComponentDefaults(CompoTempo(i), -1)
    Input #f, CompoTempo(i).Name, CompoTempo(i).MW, CompoTempo(i).InitialConcentration, CompoTempo(i).MolarVolume, CompoTempo(i).BP, CompoTempo(i).Use_K, CompoTempo(i).Use_OneOverN, CompoTempo(i).Liquid_Density, CompoTempo(i).Aqueous_Solubility, CompoTempo(i).Vapor_Pressure, CompoTempo(i).Refractive_Index
    Input #f, CompoTempo(i).SPDFR, CompoTempo(i).SPDFR_Low_Concentration, CompoTempo(i).Use_SPDFR_Correlation
    Input #f, CompoTempo(i).kf, CompoTempo(i).Ds, CompoTempo(i).Dp, CompoTempo(i).Corr(1), CompoTempo(i).Corr(2), CompoTempo(i).Corr(3)
    Input #f, CompoTempo(i).KP_User_Input(1), CompoTempo(i).KP_User_Input(2), CompoTempo(i).KP_User_Input(3)
    Input #f, CompoTempo(i).K_Reduction, CompoTempo(i).Correlation.Name, CompoTempo(i).Correlation.Coeff(1), CompoTempo(i).Correlation.Coeff(2)
    Input #f, CompoTempo(i).IsothermDB_Component_Name, CompoTempo(i).IsothermDB_Range_Num, CompoTempo(i).IPES_OrderOfMagnitude, CompoTempo(i).IPES_NumRegressionPts, CompoTempo(i).IPES_RelativeHumidity, CompoTempo(i).IPES_EstimationMethod, CompoTempo(i).Source_KandOneOverN
    Input #f, CompoTempo(i).IsothermDB_K, CompoTempo(i).IsothermDB_OneOverN, CompoTempo(i).IPESResult_K, CompoTempo(i).IPESResult_OneOverN, CompoTempo(i).UserEntered_K, CompoTempo(i).UserEntered_OneOverN
  Next i
  Input #f, BedTempo.length, BedTempo.Diameter, BedTempo.Weight, BedTempo.Flowrate, BedTempo.WaterDensity, BedTempo.WaterViscosity, BedTempo.Temperature, BedTempo.Pressure, BedTempo.Phase, BedTempo.NumberOfBeds
  Input #f, BedTempo.Water_Correlation.Name, BedTempo.Water_Correlation.Coeff(1), BedTempo.Water_Correlation.Coeff(2), BedTempo.Water_Correlation.Coeff(3), BedTempo.Water_Correlation.Coeff(4)
  Input #f, CarbonTempo.Name, CarbonTempo.Porosity, CarbonTempo.Density, CarbonTempo.ParticleRadius, CarbonTempo.Tortuosity, CarbonTempo.W0, CarbonTempo.BB, CarbonTempo.PolanyiExponent
  Input #f, State_Check_WaterT(1), State_Check_WaterT(2)
  Input #f, Use_Tortuosity_Correlation_Tempo, Constant_Tortuosity_Tempo
  Input #f, NCTempo, MCTempo
  Input #f, TimePTempo.Init, TimePTempo.End, TimePTempo.np, TimePTempo.Step
  CarbonTempo.ShapeFactor = 1#
  FileNote = ""
  Input #f, u(1), u(2), u(3), u(4), u(5)
  Call unitsys_set_units(frmMain.txtBedValue(0), u(1))
  Call unitsys_set_units(frmMain.txtBedValue(1), u(2))
  Call unitsys_set_units(frmMain.txtBedValue(2), u(3))
  Call unitsys_set_units(frmMain.txtBedValue(3), u(4))
  Call unitsys_set_units(frmMain.txtBedValue(4), u(5))
  'Call SetUnits(frmPFPSDM!txtBedUnits(0), u(1))
  'Call SetUnits(frmPFPSDM!txtBedUnits(1), u(2))
  'Call SetUnits(frmPFPSDM!txtBedUnits(2), u(3))
  'Call SetUnits(frmPFPSDM!txtBedUnits(3), u(4))
  'Call SetUnits(frmPFPSDM!txtBedUnits(4), u(5))
  Input #f, u(1), u(2)
  Call unitsys_set_units(frmMain.txtCarbon(1), u(1))
  Call unitsys_set_units(frmMain.txtCarbon(2), u(2))
  'Call SetUnits(frmPFPSDM!txtCarbonUnits(1), u(1))
  'Call SetUnits(frmPFPSDM!txtCarbonUnits(2), u(2))
  Input #f, PropertyUnits.MW, PropertyUnits.MolarVolume, PropertyUnits.BP, PropertyUnits.InitialConcentration
  Input #f, PropertyUnits.Liquid_Density, PropertyUnits.Aqueous_Solubility, PropertyUnits.Vapor_Pressure, PropertyUnits.k
  Input #f, Number_Influent_PointsTempo
  If (Number_Influent_PointsTempo > 0) Then
    For i = 1 To Number_Influent_PointsTempo
      Input #f, T_InfluentTempo(i)
      For j = 1 To Number_ComponentTempo
       Input #f, C_InfluentTempo(j, i)
      Next j
    Next i
  End If
  On Error GoTo No_Data_Points
  Input #f, NData_Points, Number_Component
  On Error GoTo 0
Read_Data_Points_v1_20:
  If (NData_Points > 0) And (Number_Component > 0) Then
    For i = 1 To NData_Points
      Input #f, T_Data_Points(i)
      For j = 1 To Number_Component
        Input #f, C_Data_Points(j, i)
      Next j
    Next i
  End If
Update_Tortuosities_from_v1_20:
  For i = 1 To Number_ComponentTempo
    CompoTempo(i).Tortuosity = CarbonTempo.Tortuosity
    CompoTempo(i).Use_Tortuosity_Correlation = Use_Tortuosity_Correlation_Tempo
    CompoTempo(i).Constant_Tortuosity = Constant_Tortuosity_Tempo
  Next i
  'IMPORT ALL TEMPORARY DATA INTO GLOBAL VARIABLES.
  'Update_All_Data
  'Store all the read variables in global variables
  If (Number_ComponentTempo > 0) Then
    Component_Number_Selected = 1
  Else
    Component_Number_Selected = -1
  End If
  Number_Component = Number_ComponentTempo
  State_Check_Water(2) = State_Check_WaterT(2)
  State_Check_Water(1) = State_Check_WaterT(1)
  'Use_SPDFR_Correlation = Use_SPDFR_Correlation_Tempo
  'SPDFR_Low_Concentration = SPDFR_Low_Concentration_tempo
'  Use_Tortuosity_Correlation = Use_Tortuosity_Correlation_Tempo
'  Constant_Tortuosity = Constant_Tortuosity_Tempo
  For i = 1 To Number_Component
    Component(i) = CompoTempo(i)
  Next i
  '========== What the heck does this code do!?!?
  '- Apparently all it does is cause problems.
  '- Commented out by ejo, 4/16/96. =================================
'  If (Mid$(VersionST, 1, 45) <> ("Input file for ASS - Windows - Version 0.80")) And (Mid$(VersionST, 1, 45) <> ("Input file for ASS - Windows - Version 1.00") And NVersion_Temp <> 1#) Then
'    If (Mid$(VersionST, 1, 45) <> ("Input file for ASS - Windows - Version 0.80")) And (Mid$(VersionST, 1, 45) <> ("Input file for ASS - Windows - Version 1.20") And NVersion_Temp <> 1.2) Then
'      If (Mid$(VersionST, 1, 45) <> ("Input file for ASS - Windows - Version 0.80")) And (Mid$(VersionST, 1, 45) <> ("Input file for ASS - Windows - Version 1.30") And NVersion_Temp <> 1.3) Then
  'If (NVersion_Temp <> 1#) Then
  '  If (NVersion_Temp <> 1.2) Then
  '    If (NVersion_Temp <> 1.3) Then
  '      For i = 1 To Number_Component
  '        Component(i).SPDFR = PSDRTempo
  '        Component(i).SPDFR_Low_Concentration = SPDFR_Low_Concentration_tempo
  '        Component(i).Use_SPDFR_Correlation = Use_SPDFR_Correlation_Tempo
  '      Next i
  '    End If
  '  End If
  'End If
  Carbon = CarbonTempo
  Bed = BedTempo
  TimeP = TimePTempo
  NC = NCTempo
  MC = MCTempo
  TimeP = TimePTempo
  Number_Influent_Points = Number_Influent_PointsTempo
  If (Number_Influent_Points > 0) Then
    For i = 1 To Number_Influent_Points
      T_Influent(i) = T_InfluentTempo(i)
      For j = 1 To Number_Component
        C_Influent(j, i) = C_InfluentTempo(j, i)
      Next j
    Next i
  End If
  Exit Sub
No_Data_Points:
  NData_Points = 0
  Number_Component = 0
  Resume Read_Data_Points_v1_20
  'Select Case NVersion_Temp
  '  Case 1#
  '    Resume Read_Data_Points_v1_00
  '  '     .
  '  '     .
  '  '     .
  '  '.... other versions ...
  '  '     .
  '  '     .
  '  '     .
  'End Select
  'Resume Read_Data_Points_v_LATEST
End Sub


'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////     VERSION 1.30     ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Sub File_Open_Legacy_v1_30(f As Integer)
Dim i As Integer, j As Integer
Dim VersionST As String, msg As String
Dim PSDR As Double
'---Declaration for tempo variables
ReDim State_Check_WaterT(2) As Integer
Dim SPDFR_Low_Concentration_tempo As Integer
Dim Use_SPDFR_Correlation_Tempo As Integer
Dim Use_Tortuosity_Correlation_Tempo As Integer
Dim Constant_Tortuosity_Tempo As Integer
Dim PSDRTempo As Double
Dim Number_ComponentTempo As Integer
Dim BedTempo As BedPropertyType
ReDim CompoTempo(Number_Compo_Max) As ComponentPropertyType
Dim CarbonTempo As CarbonPropertyType
Dim MCTempo As Integer, NCTempo As Integer
Dim TimePTempo As Para_Int
Dim Number_Influent_PointsTempo As Integer
ReDim C_InfluentTempo(Number_Compo_Max, Number_Max_Influent_Points) As Double
ReDim T_InfluentTempo(Number_Max_Influent_Points) As Double
Dim NVersion_Temp As Double
Dim temp As String
ReDim u(5) As String
Dim TempName As String
'
'END OF DECLARATION.
'
  msg = ""
'-----------------------------------
'-----   Version 1.30 ---------
'-----------------------------------
  Input #f, Number_ComponentTempo
  For i = 1 To Number_ComponentTempo
    Call SetComponentDefaults(CompoTempo(i), -1)
    '---- Modified by Eric J. Oman 8/8/97 BEGINS:
    Input #f, TempName, CompoTempo(i).MW, CompoTempo(i).InitialConcentration, CompoTempo(i).MolarVolume, CompoTempo(i).BP, CompoTempo(i).Use_K, CompoTempo(i).Use_OneOverN, CompoTempo(i).Liquid_Density, CompoTempo(i).Aqueous_Solubility, CompoTempo(i).Vapor_Pressure, CompoTempo(i).Refractive_Index
    If (Right$(Trim$(TempName), 5) = "#$1$#") Then
      Input #f, CompoTempo(i).CAS
      CompoTempo(i).Name = Left$(TempName, Len(TempName) - 5)
    Else
      CompoTempo(i).Name = TempName
    End If
    '---- Modified by Eric J. Oman 8/8/97 ENDS.
    Input #f, CompoTempo(i).SPDFR, CompoTempo(i).SPDFR_Low_Concentration, CompoTempo(i).Use_SPDFR_Correlation
    Input #f, CompoTempo(i).kf, CompoTempo(i).Ds, CompoTempo(i).Dp, CompoTempo(i).Corr(1), CompoTempo(i).Corr(2), CompoTempo(i).Corr(3)
    Input #f, CompoTempo(i).KP_User_Input(1), CompoTempo(i).KP_User_Input(2), CompoTempo(i).KP_User_Input(3)
    Input #f, CompoTempo(i).K_Reduction, CompoTempo(i).Correlation.Name, CompoTempo(i).Correlation.Coeff(1), CompoTempo(i).Correlation.Coeff(2)
    Input #f, CompoTempo(i).IsothermDB_Component_Name, CompoTempo(i).IsothermDB_Range_Num, CompoTempo(i).IPES_OrderOfMagnitude, CompoTempo(i).IPES_NumRegressionPts, CompoTempo(i).IPES_RelativeHumidity, CompoTempo(i).IPES_EstimationMethod, CompoTempo(i).Source_KandOneOverN
    Input #f, CompoTempo(i).IsothermDB_K, CompoTempo(i).IsothermDB_OneOverN, CompoTempo(i).IPESResult_K, CompoTempo(i).IPESResult_OneOverN, CompoTempo(i).UserEntered_K, CompoTempo(i).UserEntered_OneOverN
    Input #f, CompoTempo(i).Tortuosity, CompoTempo(i).Use_Tortuosity_Correlation, CompoTempo(i).Constant_Tortuosity
  Next i
  Input #f, BedTempo.length, BedTempo.Diameter, BedTempo.Weight, BedTempo.Flowrate, BedTempo.WaterDensity, BedTempo.WaterViscosity, BedTempo.Temperature, BedTempo.Pressure, BedTempo.Phase, BedTempo.NumberOfBeds
  Input #f, BedTempo.Water_Correlation.Name, BedTempo.Water_Correlation.Coeff(1), BedTempo.Water_Correlation.Coeff(2), BedTempo.Water_Correlation.Coeff(3), BedTempo.Water_Correlation.Coeff(4)
  Input #f, CarbonTempo.Name, CarbonTempo.Porosity, CarbonTempo.Density, CarbonTempo.ParticleRadius, CarbonTempo.Tortuosity, CarbonTempo.W0, CarbonTempo.BB, CarbonTempo.PolanyiExponent
  Input #f, State_Check_WaterT(1), State_Check_WaterT(2)
  Input #f, CarbonTempo.ShapeFactor, Constant_Tortuosity_Tempo       'Constant_Tortuosity_Tempo is Unused!!
  If (CarbonTempo.ShapeFactor <= 0#) Then
    CarbonTempo.ShapeFactor = 1#
  End If
  FileNote = ""
  Input #f, NCTempo, MCTempo
  Input #f, TimePTempo.Init, TimePTempo.End, TimePTempo.np, TimePTempo.Step
  Input #f, u(1), u(2), u(3), u(4), u(5)
  Call unitsys_set_units(frmMain.txtBedValue(0), u(1))
  Call unitsys_set_units(frmMain.txtBedValue(1), u(2))
  Call unitsys_set_units(frmMain.txtBedValue(2), u(3))
  Call unitsys_set_units(frmMain.txtBedValue(3), u(4))
  Call unitsys_set_units(frmMain.txtBedValue(4), u(5))
  'Call SetUnits(frmPFPSDM!txtBedUnits(0), u(1))
  'Call SetUnits(frmPFPSDM!txtBedUnits(1), u(2))
  'Call SetUnits(frmPFPSDM!txtBedUnits(2), u(3))
  'Call SetUnits(frmPFPSDM!txtBedUnits(3), u(4))
  'Call SetUnits(frmPFPSDM!txtBedUnits(4), u(5))
  Input #f, u(1), u(2)
  Call unitsys_set_units(frmMain.txtCarbon(1), u(1))
  Call unitsys_set_units(frmMain.txtCarbon(2), u(2))
  'Call SetUnits(frmPFPSDM!txtCarbonUnits(1), u(1))
  'Call SetUnits(frmPFPSDM!txtCarbonUnits(2), u(2))
  Input #f, PropertyUnits.MW, PropertyUnits.MolarVolume, PropertyUnits.BP, PropertyUnits.InitialConcentration
  Input #f, PropertyUnits.Liquid_Density, PropertyUnits.Aqueous_Solubility, PropertyUnits.Vapor_Pressure, PropertyUnits.k
  Input #f, Number_Influent_PointsTempo
  If (Number_Influent_PointsTempo > 0) Then
    For i = 1 To Number_Influent_PointsTempo
      Input #f, T_InfluentTempo(i)
      For j = 1 To Number_ComponentTempo
       Input #f, C_InfluentTempo(j, i)
      Next j
    Next i
  End If
  On Error GoTo No_Data_Points
  Input #f, NData_Points, Number_Component
  On Error GoTo 0
Read_Data_Points_v1_30:
  If (NData_Points > 0) And (Number_Component > 0) Then
    For i = 1 To NData_Points
      Input #f, T_Data_Points(i)
      For j = 1 To Number_Component
        Input #f, C_Data_Points(j, i)
      Next j
    Next i
  End If
  'IMPORT ALL TEMPORARY DATA INTO GLOBAL VARIABLES.
  'Update_All_Data
  'Store all the read variables in global variables
  If (Number_ComponentTempo > 0) Then
    Component_Number_Selected = 1
  Else
    Component_Number_Selected = -1
  End If
  Number_Component = Number_ComponentTempo
  State_Check_Water(2) = State_Check_WaterT(2)
  State_Check_Water(1) = State_Check_WaterT(1)
  'Use_SPDFR_Correlation = Use_SPDFR_Correlation_Tempo
  'SPDFR_Low_Concentration = SPDFR_Low_Concentration_tempo
'  Use_Tortuosity_Correlation = Use_Tortuosity_Correlation_Tempo
'  Constant_Tortuosity = Constant_Tortuosity_Tempo
  For i = 1 To Number_Component
    Component(i) = CompoTempo(i)
  Next i
  '========== What the heck does this code do!?!?
  '- Apparently all it does is cause problems.
  '- Commented out by ejo, 4/16/96. =================================
'  If (Mid$(VersionST, 1, 45) <> ("Input file for ASS - Windows - Version 0.80")) And (Mid$(VersionST, 1, 45) <> ("Input file for ASS - Windows - Version 1.00") And NVersion_Temp <> 1#) Then
'    If (Mid$(VersionST, 1, 45) <> ("Input file for ASS - Windows - Version 0.80")) And (Mid$(VersionST, 1, 45) <> ("Input file for ASS - Windows - Version 1.20") And NVersion_Temp <> 1.2) Then
'      If (Mid$(VersionST, 1, 45) <> ("Input file for ASS - Windows - Version 0.80")) And (Mid$(VersionST, 1, 45) <> ("Input file for ASS - Windows - Version 1.30") And NVersion_Temp <> 1.3) Then
  'If (NVersion_Temp <> 1#) Then
  '  If (NVersion_Temp <> 1.2) Then
  '    If (NVersion_Temp <> 1.3) Then
  '      For i = 1 To Number_Component
  '        Component(i).SPDFR = PSDRTempo
  '        Component(i).SPDFR_Low_Concentration = SPDFR_Low_Concentration_tempo
  '        Component(i).Use_SPDFR_Correlation = Use_SPDFR_Correlation_Tempo
  '      Next i
  '    End If
  '  End If
  'End If
  Carbon = CarbonTempo
  Bed = BedTempo
  TimeP = TimePTempo
  NC = NCTempo
  MC = MCTempo
  TimeP = TimePTempo
  Number_Influent_Points = Number_Influent_PointsTempo
  If (Number_Influent_Points > 0) Then
    For i = 1 To Number_Influent_Points
      T_Influent(i) = T_InfluentTempo(i)
      For j = 1 To Number_Component
        C_Influent(j, i) = C_InfluentTempo(j, i)
      Next j
    Next i
  End If
  Exit Sub
No_Data_Points:
  NData_Points = 0
  Number_Component = 0
  Resume Read_Data_Points_v1_30
  'Select Case NVersion_Temp
  '  Case 1#
  '    Resume Read_Data_Points_v1_00
  '  '     .
  '  '     .
  '  '     .
  '  '.... other versions ...
  '  '     .
  '  '     .
  '  '     .
  'End Select
  'Resume Read_Data_Points_v_LATEST
End Sub


'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////     VERSION 1.42     ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Sub File_Open_Legacy_v1_42(f As Integer)
Dim i As Integer, j As Integer
Dim VersionST As String, msg As String
Dim PSDR As Double
'---Declaration for tempo variables
ReDim State_Check_WaterT(2) As Integer
Dim SPDFR_Low_Concentration_tempo As Integer
Dim Use_SPDFR_Correlation_Tempo As Integer
Dim Use_Tortuosity_Correlation_Tempo As Integer
Dim Constant_Tortuosity_Tempo As Integer
Dim PSDRTempo As Double
Dim Number_ComponentTempo As Integer
Dim BedTempo As BedPropertyType
ReDim CompoTempo(Number_Compo_Max) As ComponentPropertyType
Dim CarbonTempo As CarbonPropertyType
Dim MCTempo As Integer, NCTempo As Integer
Dim TimePTempo As Para_Int
Dim Number_Influent_PointsTempo As Integer
ReDim C_InfluentTempo(Number_Compo_Max, Number_Max_Influent_Points) As Double
ReDim T_InfluentTempo(Number_Max_Influent_Points) As Double
Dim NVersion_Temp As Double
Dim temp As String
ReDim u(5) As String
Dim TempName As String
Dim DummyStr As String
'
'END OF DECLARATION.
'
'-----------------------------------
'-----   Version 1.42 ---------
'-----------------------------------
    Input #f, Number_ComponentTempo
    For i = 1 To Number_ComponentTempo
      Call SetComponentDefaults(CompoTempo(i), -1)
      '---- Modified by Eric J. Oman 8/8/97 BEGINS:
      Input #f, TempName, CompoTempo(i).MW, CompoTempo(i).InitialConcentration, CompoTempo(i).MolarVolume, CompoTempo(i).BP, CompoTempo(i).Use_K, CompoTempo(i).Use_OneOverN, CompoTempo(i).Liquid_Density, CompoTempo(i).Aqueous_Solubility, CompoTempo(i).Vapor_Pressure, CompoTempo(i).Refractive_Index
      If (Right$(Trim$(TempName), 5) = "#$1$#") Then
        Input #f, CompoTempo(i).CAS
        CompoTempo(i).Name = Left$(TempName, Len(TempName) - 5)
      Else
        CompoTempo(i).Name = TempName
      End If
      '---- Modified by Eric J. Oman 8/8/97 ENDS.
      Input #f, CompoTempo(i).SPDFR, CompoTempo(i).SPDFR_Low_Concentration, CompoTempo(i).Use_SPDFR_Correlation
      Input #f, CompoTempo(i).kf, CompoTempo(i).Ds, CompoTempo(i).Dp, CompoTempo(i).Corr(1), CompoTempo(i).Corr(2), CompoTempo(i).Corr(3)
      Input #f, CompoTempo(i).KP_User_Input(1), CompoTempo(i).KP_User_Input(2), CompoTempo(i).KP_User_Input(3)
      Input #f, CompoTempo(i).K_Reduction, CompoTempo(i).Correlation.Name, CompoTempo(i).Correlation.Coeff(1), CompoTempo(i).Correlation.Coeff(2)
      Input #f, CompoTempo(i).IsothermDB_Component_Name, CompoTempo(i).IsothermDB_Range_Num, CompoTempo(i).IPES_OrderOfMagnitude, CompoTempo(i).IPES_NumRegressionPts, CompoTempo(i).IPES_RelativeHumidity, CompoTempo(i).IPES_EstimationMethod, CompoTempo(i).Source_KandOneOverN
      Input #f, CompoTempo(i).IsothermDB_K, CompoTempo(i).IsothermDB_OneOverN, CompoTempo(i).IPESResult_K, CompoTempo(i).IPESResult_OneOverN, CompoTempo(i).UserEntered_K, CompoTempo(i).UserEntered_OneOverN
      Input #f, CompoTempo(i).Tortuosity, CompoTempo(i).Use_Tortuosity_Correlation, CompoTempo(i).Constant_Tortuosity
    Next i
    Input #f, BedTempo.length, BedTempo.Diameter, BedTempo.Weight, BedTempo.Flowrate, BedTempo.WaterDensity, BedTempo.WaterViscosity, BedTempo.Temperature, BedTempo.Pressure, BedTempo.Phase, BedTempo.NumberOfBeds
    Input #f, BedTempo.Water_Correlation.Name, BedTempo.Water_Correlation.Coeff(1), BedTempo.Water_Correlation.Coeff(2), BedTempo.Water_Correlation.Coeff(3), BedTempo.Water_Correlation.Coeff(4)
    Input #f, CarbonTempo.Name, CarbonTempo.Porosity, CarbonTempo.Density, CarbonTempo.ParticleRadius, CarbonTempo.Tortuosity, CarbonTempo.W0, CarbonTempo.BB, CarbonTempo.PolanyiExponent
    Input #f, State_Check_WaterT(1), State_Check_WaterT(2)
    Input #f, CarbonTempo.ShapeFactor, Constant_Tortuosity_Tempo       'Constant_Tortuosity_Tempo is Unused!!
    If (CarbonTempo.ShapeFactor <= 0#) Then
      CarbonTempo.ShapeFactor = 1#
    End If
    FileNote = ""
    Input #f, NCTempo, MCTempo
    Input #f, TimePTempo.Init, TimePTempo.End, TimePTempo.np, TimePTempo.Step
    Input #f, u(1), u(2), u(3), u(4), u(5)
    Call unitsys_set_units(frmMain.txtBedValue(0), u(1))
    Call unitsys_set_units(frmMain.txtBedValue(1), u(2))
    Call unitsys_set_units(frmMain.txtBedValue(2), u(3))
    Call unitsys_set_units(frmMain.txtBedValue(3), u(4))
    Call unitsys_set_units(frmMain.txtBedValue(4), u(5))
    'Call SetUnits(frmPFPSDM!txtBedUnits(0), u(1))
    'Call SetUnits(frmPFPSDM!txtBedUnits(1), u(2))
    'Call SetUnits(frmPFPSDM!txtBedUnits(2), u(3))
    'Call SetUnits(frmPFPSDM!txtBedUnits(3), u(4))
    'Call SetUnits(frmPFPSDM!txtBedUnits(4), u(5))
    Input #f, u(1), u(2)
    Call unitsys_set_units(frmMain.txtCarbon(1), u(1))
    Call unitsys_set_units(frmMain.txtCarbon(2), u(2))
    'Call SetUnits(frmPFPSDM!txtCarbonUnits(1), u(1))
    'Call SetUnits(frmPFPSDM!txtCarbonUnits(2), u(2))
    Input #f, PropertyUnits.MW, PropertyUnits.MolarVolume, PropertyUnits.BP, PropertyUnits.InitialConcentration
    Input #f, PropertyUnits.Liquid_Density, PropertyUnits.Aqueous_Solubility, PropertyUnits.Vapor_Pressure, PropertyUnits.k
    '---- Modified by Eric J. Oman 6/25/98 BEGINS:
    'NOTE: THIS IS A MODIFICATION FOR V1.40 OF THE DATA FILE.
    Input #f, DummyStr
    Input #f, RoomParams.ROOM_VOL, DummyStr
    Input #f, RoomParams.ROOM_FLOWRATE, DummyStr
    For i = 1 To Number_ComponentTempo
      Input #f, RoomParams.ROOM_C0(i), DummyStr
      Input #f, RoomParams.ROOM_EMIT(i), DummyStr
    Next i
    Input #f, RoomParams.ROOM_VOL_Units
    Input #f, RoomParams.ROOM_FLOWRATE_Units
    Input #f, RoomParams.ROOM_C0_Units
    Input #f, RoomParams.ROOM_EMIT_Units
    RoomParams.COUNT_CONTAMINANT = Number_ComponentTempo
    Call RoomParam_Recalculate(RoomParams)
    'MsgBox RoomParams.ROOM_VOL_Units
    '---- Modified by Eric J. Oman 6/25/98 ENDS.
    
    '---- Modified by Eric J. Oman 9/16/98 BEGINS:
    'NOTE: THIS IS A MODIFICATION FOR V1.42 OF THE DATA FILE.
    For i = 1 To Number_ComponentTempo
      Input #f, RoomParams.INITIAL_ROOM_CONC(i)
    Next i
    Input #f, RoomParams.INITIAL_ROOM_CONC_Units
    '---- Modified by Eric J. Oman 9/16/98 ENDS.

    Input #f, Number_Influent_PointsTempo
    If (Number_Influent_PointsTempo > 0) Then
      For i = 1 To Number_Influent_PointsTempo
        Input #f, T_InfluentTempo(i)
        For j = 1 To Number_ComponentTempo
         Input #f, C_InfluentTempo(j, i)
        Next j
      Next i
    End If
    On Error GoTo No_Data_Points
    Input #f, NData_Points, Number_Component
    On Error GoTo 0
Read_Data_Points_v1_42:
    If (NData_Points > 0) And (Number_Component > 0) Then
      For i = 1 To NData_Points
        Input #f, T_Data_Points(i)
        For j = 1 To Number_Component
          Input #f, C_Data_Points(j, i)
        Next j
      Next i
    End If
    ''''Close (f)
    ''''GoTo Start_Updating
  ''''End If
  'IMPORT ALL TEMPORARY DATA INTO GLOBAL VARIABLES.
  'Update_All_Data
  'Store all the read variables in global variables
  If (Number_ComponentTempo > 0) Then
    Component_Number_Selected = 1
  Else
    Component_Number_Selected = -1
  End If
  Number_Component = Number_ComponentTempo
  State_Check_Water(2) = State_Check_WaterT(2)
  State_Check_Water(1) = State_Check_WaterT(1)
  'Use_SPDFR_Correlation = Use_SPDFR_Correlation_Tempo
  'SPDFR_Low_Concentration = SPDFR_Low_Concentration_tempo
'  Use_Tortuosity_Correlation = Use_Tortuosity_Correlation_Tempo
'  Constant_Tortuosity = Constant_Tortuosity_Tempo
  For i = 1 To Number_Component
    Component(i) = CompoTempo(i)
  Next i
  '========== What the heck does this code do!?!?
  '- Apparently all it does is cause problems.
  '- Commented out by ejo, 4/16/96. =================================
'  If (Mid$(VersionST, 1, 45) <> ("Input file for ASS - Windows - Version 0.80")) And (Mid$(VersionST, 1, 45) <> ("Input file for ASS - Windows - Version 1.00") And NVersion_Temp <> 1#) Then
'    If (Mid$(VersionST, 1, 45) <> ("Input file for ASS - Windows - Version 0.80")) And (Mid$(VersionST, 1, 45) <> ("Input file for ASS - Windows - Version 1.20") And NVersion_Temp <> 1.2) Then
'      If (Mid$(VersionST, 1, 45) <> ("Input file for ASS - Windows - Version 0.80")) And (Mid$(VersionST, 1, 45) <> ("Input file for ASS - Windows - Version 1.30") And NVersion_Temp <> 1.3) Then
  'If (NVersion_Temp <> 1#) Then
  '  If (NVersion_Temp <> 1.2) Then
  '    If (NVersion_Temp <> 1.3) Then
  '      For i = 1 To Number_Component
  '        Component(i).SPDFR = PSDRTempo
  '        Component(i).SPDFR_Low_Concentration = SPDFR_Low_Concentration_tempo
  '        Component(i).Use_SPDFR_Correlation = Use_SPDFR_Correlation_Tempo
  '      Next i
  '    End If
  '  End If
  'End If
  Carbon = CarbonTempo
  Bed = BedTempo
  TimeP = TimePTempo
  NC = NCTempo
  MC = MCTempo
  TimeP = TimePTempo
  Number_Influent_Points = Number_Influent_PointsTempo
  If (Number_Influent_Points > 0) Then
    For i = 1 To Number_Influent_Points
      T_Influent(i) = T_InfluentTempo(i)
      For j = 1 To Number_Component
        C_Influent(j, i) = C_InfluentTempo(j, i)
      Next j
    Next i
  End If
  Exit Sub
No_Data_Points:
  NData_Points = 0
  Number_Component = 0
  Resume Read_Data_Points_v1_42

End Sub




'Sub openfile()
'    Dim f As Integer, i  As Integer, J As Integer
'    Dim VersionST As String, msg As String
'    Dim PSDR As Double
'    '---Declaration for tempo variables
'    ReDim State_Check_WaterT(2) As Integer
'    Dim SPDFR_Low_Concentration_tempo As Integer
'    Dim Use_SPDFR_Correlation_Tempo As Integer
'    Dim Use_Tortuosity_Correlation_Tempo As Integer
'    Dim Constant_Tortuosity_Tempo As Integer
'    Dim PSDRTempo As Double
'    Dim Number_ComponentTempo As Integer
'    Dim BedTempo As BedPropertyType
'    ReDim CompoTempo(Number_Compo_Max) As ComponentPropertyType
'    Dim CarbonTempo As CarbonPropertyType
'    Dim MCTempo As Integer, NCTempo As Integer
'    Dim TimePTempo As Para_Int
'    Dim Number_Influent_PointsTempo As Integer
'    ReDim C_InfluentTempo(Number_Compo_Max, Number_Max_Influent_Points) As Double
'    ReDim T_InfluentTempo(Number_Max_Influent_Points) As Double
'    Dim NVersion_Temp As Double
'    Dim temp As String
'    ReDim u(5) As String
'
'    Dim TempName As String
''---End declaration
'
'  msg = ""
'  On Error GoTo OpenError
'  f = FreeFile
'  Open Filename For Input As f
'
'  Input #f, NVersion_Temp
'  If (NVersion_Temp = 1#) Then
''-----------------------------------
''-----   Version 1.00 ---------
''-----------------------------------
'    If batchrun <> 1 Then
'     'temp = "Note--this is an Adsorption Simulation Software version 1.00 file."
'     'temp = temp & NL & "If you save this file, it will be saved in version " & Format$(NVersion, "0.00") & " format."
'     'MsgBox temp, MB_ICONEXCLAMATION, AppName_For_Display_long
'    End If
'
'    Input #f, Number_ComponentTempo
'    For i = 1 To Number_ComponentTempo
'      Call SetComponentDefaults(CompoTempo(i), -1)
'      Input #f, CompoTempo(i).Name, CompoTempo(i).MW, CompoTempo(i).InitialConcentration, CompoTempo(i).MolarVolume, CompoTempo(i).BP, CompoTempo(i).Use_K, CompoTempo(i).Use_OneOverN
'      CompoTempo(i).UserEntered_K = CompoTempo(i).Use_K
'      CompoTempo(i).UserEntered_OneOverN = CompoTempo(i).Use_OneOverN
'      Input #f, CompoTempo(i).SPDFR, CompoTempo(i).SPDFR_Low_Concentration, CompoTempo(i).Use_SPDFR_Correlation
'      Input #f, CompoTempo(i).kf, CompoTempo(i).Ds, CompoTempo(i).Dp, CompoTempo(i).Corr(1), CompoTempo(i).Corr(2), CompoTempo(i).Corr(3)
'      Input #f, CompoTempo(i).KP_User_Input(1), CompoTempo(i).KP_User_Input(2), CompoTempo(i).KP_User_Input(3)
'      Input #f, CompoTempo(i).K_Reduction, CompoTempo(i).Correlation.Name, CompoTempo(i).Correlation.Coeff(1), CompoTempo(i).Correlation.Coeff(2)
'    Next i
'
'    Input #f, BedTempo.Length, BedTempo.Diameter, BedTempo.Weight, BedTempo.Flowrate, BedTempo.WaterDensity, BedTempo.WaterViscosity, BedTempo.Temperature, BedTempo.Pressure, BedTempo.Phase, BedTempo.NumberOfBeds
'    Input #f, BedTempo.Water_Correlation.Name, BedTempo.Water_Correlation.Coeff(1), BedTempo.Water_Correlation.Coeff(2), BedTempo.Water_Correlation.Coeff(3), BedTempo.Water_Correlation.Coeff(4)
'    Input #f, CarbonTempo.Name, CarbonTempo.Porosity, CarbonTempo.Density, CarbonTempo.ParticleRadius, CarbonTempo.tortuosity
'    Input #f, State_Check_WaterT(1), State_Check_WaterT(2)
'    Input #f, Use_Tortuosity_Correlation_Tempo, Constant_Tortuosity_Tempo
'    Input #f, NCTempo, MCTempo
'    Input #f, TimePTempo.Init, TimePTempo.End, TimePTempo.np, TimePTempo.Step
'    Input #f, Number_Influent_PointsTempo
'    If (Number_Influent_PointsTempo > 0) Then
'      For i = 1 To Number_Influent_PointsTempo
'        Input #f, T_InfluentTempo(i)
'        For J = 1 To Number_ComponentTempo
'         Input #f, C_InfluentTempo(J, i)
'        Next J
'      Next i
'    End If
'    On Error GoTo No_Data_Points
'    Input #f, NData_Points, Number_Component
'    On Error GoTo OpenError
'    CarbonTempo.ShapeFactor = 1#
'
'Read_Data_Points_v1_00:
'    If (NData_Points > 0 And Number_Component > 0) Then
'      For i = 1 To NData_Points
'        Input #f, T_Data_Points(i)
'        For J = 1 To Number_Component
'          Input #f, C_Data_Points(J, i)
'        Next J
'      Next i
'    End If
'
'Update_Tortuosities_from_v1_00:
'    For i = 1 To Number_ComponentTempo
'      CompoTempo(i).tortuosity = CarbonTempo.tortuosity
'      CompoTempo(i).Use_Tortuosity_Correlation = Use_Tortuosity_Correlation_Tempo
'      CompoTempo(i).Constant_Tortuosity = Constant_Tortuosity_Tempo
'    Next i
'
'    Close (f)
'    GoTo Start_Updating
'  End If
'
'  If (NVersion_Temp = 1.2) Then
''-----------------------------------
''-----   Version 1.20 ---------
''-----------------------------------
'    If batchrun <> 1 Then
'     'temp = "Note--this is an Adsorption Simulation Software version 1.20 file."
'     'temp = temp & NL & "If you save this file, it will be saved in version " & Format$(NVersion, "0.00") & " format."
'     'MsgBox temp, MB_ICONEXCLAMATION, AppName_For_Display_long
'    End If
'
'    Input #f, Number_ComponentTempo
'    For i = 1 To Number_ComponentTempo
'      Call SetComponentDefaults(CompoTempo(i), -1)
'      Input #f, CompoTempo(i).Name, CompoTempo(i).MW, CompoTempo(i).InitialConcentration, CompoTempo(i).MolarVolume, CompoTempo(i).BP, CompoTempo(i).Use_K, CompoTempo(i).Use_OneOverN, CompoTempo(i).Liquid_Density, CompoTempo(i).Aqueous_Solubility, CompoTempo(i).Vapor_Pressure, CompoTempo(i).Refractive_Index
'      Input #f, CompoTempo(i).SPDFR, CompoTempo(i).SPDFR_Low_Concentration, CompoTempo(i).Use_SPDFR_Correlation
'      Input #f, CompoTempo(i).kf, CompoTempo(i).Ds, CompoTempo(i).Dp, CompoTempo(i).Corr(1), CompoTempo(i).Corr(2), CompoTempo(i).Corr(3)
'      Input #f, CompoTempo(i).KP_User_Input(1), CompoTempo(i).KP_User_Input(2), CompoTempo(i).KP_User_Input(3)
'      Input #f, CompoTempo(i).K_Reduction, CompoTempo(i).Correlation.Name, CompoTempo(i).Correlation.Coeff(1), CompoTempo(i).Correlation.Coeff(2)
'      Input #f, CompoTempo(i).IsothermDB_Component_Name, CompoTempo(i).IsothermDB_Range_Num, CompoTempo(i).IPES_OrderOfMagnitude, CompoTempo(i).IPES_NumRegressionPts, CompoTempo(i).IPES_RelativeHumidity, CompoTempo(i).IPES_EstimationMethod, CompoTempo(i).Source_KandOneOverN
'      Input #f, CompoTempo(i).IsothermDB_K, CompoTempo(i).IsothermDB_OneOverN, CompoTempo(i).IPESResult_K, CompoTempo(i).IPESResult_OneOverN, CompoTempo(i).UserEntered_K, CompoTempo(i).UserEntered_OneOverN
'    Next i
'
'    Input #f, BedTempo.Length, BedTempo.Diameter, BedTempo.Weight, BedTempo.Flowrate, BedTempo.WaterDensity, BedTempo.WaterViscosity, BedTempo.Temperature, BedTempo.Pressure, BedTempo.Phase, BedTempo.NumberOfBeds
'    Input #f, BedTempo.Water_Correlation.Name, BedTempo.Water_Correlation.Coeff(1), BedTempo.Water_Correlation.Coeff(2), BedTempo.Water_Correlation.Coeff(3), BedTempo.Water_Correlation.Coeff(4)
'    Input #f, CarbonTempo.Name, CarbonTempo.Porosity, CarbonTempo.Density, CarbonTempo.ParticleRadius, CarbonTempo.tortuosity, CarbonTempo.W0, CarbonTempo.BB, CarbonTempo.PolanyiExponent
'    Input #f, State_Check_WaterT(1), State_Check_WaterT(2)
'    Input #f, Use_Tortuosity_Correlation_Tempo, Constant_Tortuosity_Tempo
'    Input #f, NCTempo, MCTempo
'    Input #f, TimePTempo.Init, TimePTempo.End, TimePTempo.np, TimePTempo.Step
'
'    CarbonTempo.ShapeFactor = 1#
'
'    Input #f, u(1), u(2), u(3), u(4), u(5)
'    Call SetUnits(frmPFPSDM!txtBedUnits(0), u(1))
'    Call SetUnits(frmPFPSDM!txtBedUnits(1), u(2))
'    Call SetUnits(frmPFPSDM!txtBedUnits(2), u(3))
'    Call SetUnits(frmPFPSDM!txtBedUnits(3), u(4))
'    Call SetUnits(frmPFPSDM!txtBedUnits(4), u(5))
'
'    Input #f, u(1), u(2)
'    Call SetUnits(frmPFPSDM!txtCarbonUnits(1), u(1))
'    Call SetUnits(frmPFPSDM!txtCarbonUnits(2), u(2))
'
'    Input #f, PropertyUnits.MW, PropertyUnits.MolarVolume, PropertyUnits.BP, PropertyUnits.InitialConcentration
'    Input #f, PropertyUnits.Liquid_Density, PropertyUnits.Aqueous_Solubility, PropertyUnits.Vapor_Pressure, PropertyUnits.k
'
'    Input #f, Number_Influent_PointsTempo
'    If (Number_Influent_PointsTempo > 0) Then
'      For i = 1 To Number_Influent_PointsTempo
'        Input #f, T_InfluentTempo(i)
'        For J = 1 To Number_ComponentTempo
'         Input #f, C_InfluentTempo(J, i)
'        Next J
'      Next i
'    End If
'    On Error GoTo No_Data_Points
'    Input #f, NData_Points, Number_Component
'    On Error GoTo OpenError
'
'Read_Data_Points_v1_20:
'    If (NData_Points > 0) And (Number_Component > 0) Then
'      For i = 1 To NData_Points
'        Input #f, T_Data_Points(i)
'        For J = 1 To Number_Component
'          Input #f, C_Data_Points(J, i)
'        Next J
'      Next i
'    End If
'
'    Close (f)
'
'Update_Tortuosities_from_v1_20:
'    For i = 1 To Number_ComponentTempo
'      CompoTempo(i).tortuosity = CarbonTempo.tortuosity
'      CompoTempo(i).Use_Tortuosity_Correlation = Use_Tortuosity_Correlation_Tempo
'      CompoTempo(i).Constant_Tortuosity = Constant_Tortuosity_Tempo
'    Next i
'
'    GoTo Start_Updating
'  End If
'
'  If (NVersion_Temp = 1.3) Then
''-----------------------------------
''-----   Version 1.30 ---------
''-----------------------------------
'    Input #f, Number_ComponentTempo
'    For i = 1 To Number_ComponentTempo
'      Call SetComponentDefaults(CompoTempo(i), -1)
'      '---- Modified by Eric J. Oman 8/8/97 BEGINS:
'      Input #f, TempName, CompoTempo(i).MW, CompoTempo(i).InitialConcentration, CompoTempo(i).MolarVolume, CompoTempo(i).BP, CompoTempo(i).Use_K, CompoTempo(i).Use_OneOverN, CompoTempo(i).Liquid_Density, CompoTempo(i).Aqueous_Solubility, CompoTempo(i).Vapor_Pressure, CompoTempo(i).Refractive_Index
'      If (Right$(Trim$(TempName), 5) = "#$1$#") Then
'        Input #f, CompoTempo(i).CAS
'        CompoTempo(i).Name = Left$(TempName, Len(TempName) - 5)
'      Else
'        CompoTempo(i).Name = TempName
'      End If
'      '---- Modified by Eric J. Oman 8/8/97 ENDS.
'      Input #f, CompoTempo(i).SPDFR, CompoTempo(i).SPDFR_Low_Concentration, CompoTempo(i).Use_SPDFR_Correlation
'      Input #f, CompoTempo(i).kf, CompoTempo(i).Ds, CompoTempo(i).Dp, CompoTempo(i).Corr(1), CompoTempo(i).Corr(2), CompoTempo(i).Corr(3)
'      Input #f, CompoTempo(i).KP_User_Input(1), CompoTempo(i).KP_User_Input(2), CompoTempo(i).KP_User_Input(3)
'      Input #f, CompoTempo(i).K_Reduction, CompoTempo(i).Correlation.Name, CompoTempo(i).Correlation.Coeff(1), CompoTempo(i).Correlation.Coeff(2)
'      Input #f, CompoTempo(i).IsothermDB_Component_Name, CompoTempo(i).IsothermDB_Range_Num, CompoTempo(i).IPES_OrderOfMagnitude, CompoTempo(i).IPES_NumRegressionPts, CompoTempo(i).IPES_RelativeHumidity, CompoTempo(i).IPES_EstimationMethod, CompoTempo(i).Source_KandOneOverN
'      Input #f, CompoTempo(i).IsothermDB_K, CompoTempo(i).IsothermDB_OneOverN, CompoTempo(i).IPESResult_K, CompoTempo(i).IPESResult_OneOverN, CompoTempo(i).UserEntered_K, CompoTempo(i).UserEntered_OneOverN
'      Input #f, CompoTempo(i).tortuosity, CompoTempo(i).Use_Tortuosity_Correlation, CompoTempo(i).Constant_Tortuosity
'    Next i
'
'    Input #f, BedTempo.Length, BedTempo.Diameter, BedTempo.Weight, BedTempo.Flowrate, BedTempo.WaterDensity, BedTempo.WaterViscosity, BedTempo.Temperature, BedTempo.Pressure, BedTempo.Phase, BedTempo.NumberOfBeds
'    Input #f, BedTempo.Water_Correlation.Name, BedTempo.Water_Correlation.Coeff(1), BedTempo.Water_Correlation.Coeff(2), BedTempo.Water_Correlation.Coeff(3), BedTempo.Water_Correlation.Coeff(4)
'    Input #f, CarbonTempo.Name, CarbonTempo.Porosity, CarbonTempo.Density, CarbonTempo.ParticleRadius, CarbonTempo.tortuosity, CarbonTempo.W0, CarbonTempo.BB, CarbonTempo.PolanyiExponent
'    Input #f, State_Check_WaterT(1), State_Check_WaterT(2)
'    Input #f, CarbonTempo.ShapeFactor, Constant_Tortuosity_Tempo       'Constant_Tortuosity_Tempo is Unused!!
'    If (CarbonTempo.ShapeFactor <= 0#) Then
'      CarbonTempo.ShapeFactor = 1#
'    End If
'    Input #f, NCTempo, MCTempo
'    Input #f, TimePTempo.Init, TimePTempo.End, TimePTempo.np, TimePTempo.Step
'
'    Input #f, u(1), u(2), u(3), u(4), u(5)
'    Call SetUnits(frmPFPSDM!txtBedUnits(0), u(1))
'    Call SetUnits(frmPFPSDM!txtBedUnits(1), u(2))
'    Call SetUnits(frmPFPSDM!txtBedUnits(2), u(3))
'    Call SetUnits(frmPFPSDM!txtBedUnits(3), u(4))
'    Call SetUnits(frmPFPSDM!txtBedUnits(4), u(5))
'
'    Input #f, u(1), u(2)
'    Call SetUnits(frmPFPSDM!txtCarbonUnits(1), u(1))
'    Call SetUnits(frmPFPSDM!txtCarbonUnits(2), u(2))
'
'    Input #f, PropertyUnits.MW, PropertyUnits.MolarVolume, PropertyUnits.BP, PropertyUnits.InitialConcentration
'    Input #f, PropertyUnits.Liquid_Density, PropertyUnits.Aqueous_Solubility, PropertyUnits.Vapor_Pressure, PropertyUnits.k
'
'    Input #f, Number_Influent_PointsTempo
'    If (Number_Influent_PointsTempo > 0) Then
'      For i = 1 To Number_Influent_PointsTempo
'        Input #f, T_InfluentTempo(i)
'        For J = 1 To Number_ComponentTempo
'         Input #f, C_InfluentTempo(J, i)
'        Next J
'      Next i
'    End If
'    On Error GoTo No_Data_Points
'    Input #f, NData_Points, Number_Component
'    On Error GoTo OpenError
'
'Read_Data_Points_v_LATEST:
'    If (NData_Points > 0) And (Number_Component > 0) Then
'      For i = 1 To NData_Points
'        Input #f, T_Data_Points(i)
'        For J = 1 To Number_Component
'          Input #f, C_Data_Points(J, i)
'        Next J
'      Next i
'    End If
'
'    Close (f)
'    GoTo Start_Updating
'  Else
'    'MsgBox "This file is not a file for Adsorption Simulation Software Version " & Format$(NVersion, "0.00"), MB_ICONEXCLAMATION, AppName_For_Display_long
'    MsgBox "This file is not a file for AdXsorption Design Software Version " & Format$(NVersion, "0.00"), MB_ICONEXCLAMATION, AppName_For_Display_long
'    Close (f)
'    GoTo Exit_Open
'  End If
'
'Start_Updating:
'  'Update_All_Data
'  'Store all the read variables in global variables
'  If (Number_ComponentTempo > 0) Then
'    Component_Number_Selected = 1
'  Else
'    Component_Number_Selected = -1
'  End If
'
'  Number_Component = Number_ComponentTempo
'  State_Check_Water(2) = State_Check_WaterT(2)
'  State_Check_Water(1) = State_Check_WaterT(1)
'  'Use_SPDFR_Correlation = Use_SPDFR_Correlation_Tempo
'  'SPDFR_Low_Concentration = SPDFR_Low_Concentration_tempo
''  Use_Tortuosity_Correlation = Use_Tortuosity_Correlation_Tempo
''  Constant_Tortuosity = Constant_Tortuosity_Tempo
'
'  For i = 1 To Number_Component
'    Component(i) = CompoTempo(i)
'  Next i
'  '========== What the heck does this code do!?!?
'  '- Apparently all it does is cause problems.
'  '- Commented out by ejo, 4/16/96. =================================
''  If (Mid$(VersionST, 1, 45) <> ("Input file for ASS - Windows - Version 0.80")) And (Mid$(VersionST, 1, 45) <> ("Input file for ASS - Windows - Version 1.00") And NVersion_Temp <> 1#) Then
''    If (Mid$(VersionST, 1, 45) <> ("Input file for ASS - Windows - Version 0.80")) And (Mid$(VersionST, 1, 45) <> ("Input file for ASS - Windows - Version 1.20") And NVersion_Temp <> 1.2) Then
''      If (Mid$(VersionST, 1, 45) <> ("Input file for ASS - Windows - Version 0.80")) And (Mid$(VersionST, 1, 45) <> ("Input file for ASS - Windows - Version 1.30") And NVersion_Temp <> 1.3) Then
'  'If (NVersion_Temp <> 1#) Then
'  '  If (NVersion_Temp <> 1.2) Then
'  '    If (NVersion_Temp <> 1.3) Then
'  '      For i = 1 To Number_Component
'  '        Component(i).SPDFR = PSDRTempo
'  '        Component(i).SPDFR_Low_Concentration = SPDFR_Low_Concentration_tempo
'  '        Component(i).Use_SPDFR_Correlation = Use_SPDFR_Correlation_Tempo
'  '      Next i
'  '    End If
'  '  End If
'  'End If
'
'  Carbon = CarbonTempo
'  Bed = BedTempo
'  TimeP = TimePTempo
'
'  NC = NCTempo
'  MC = MCTempo
'  TimeP = TimePTempo
'
'  Number_Influent_Points = Number_Influent_PointsTempo
'  If (Number_Influent_Points > 0) Then
'    For i = 1 To Number_Influent_Points
'      T_Influent(i) = T_InfluentTempo(i)
'      For J = 1 To Number_Component
'        C_Influent(J, i) = C_InfluentTempo(J, i)
'      Next J
'    Next i
'  End If
''Update the display
'  Call Update_Display_Data
'  Call Update_Display_Kinetic
'  Call Update_Bed_Density_Display
'  Call Update_Several_Bed_Properties(3)
'
''------------------------------------------------
'  frmPFPSDM!mnuFileItem(2).Enabled = True
'  frmPFPSDM!mnuFileItem(3).Enabled = True
'  If (Number_Component > 0) Then
'    frmPFPSDM!mnuRunItem(0).Enabled = True
'    frmPFPSDM!mnuRunItem(1).Enabled = True
'    frmPFPSDM!mnuRunItem(2).Enabled = True
'
'    frmPFPSDM!mnuResultsItem(0).Enabled = False
'    frmPFPSDM!mnuResultsItem(1).Enabled = False
'    frmPFPSDM!mnuResultsItem(2).Enabled = False
'    frmPFPSDM!mnuResultsItem(3).Enabled = False
'    frmPFPSDM!mnuResultsItem(4).Enabled = False
'
'    frmPFPSDM!mnuOptionsItem(0).Enabled = True
'    frmPFPSDM!mnuOptionsItem(1).Enabled = True
'    frmPFPSDM!mnuOptionsItem(2).Enabled = True
'  Else
'    frmPFPSDM!mnuRunItem(0).Enabled = False
'    frmPFPSDM!mnuRunItem(1).Enabled = False
'    frmPFPSDM!mnuRunItem(2).Enabled = False
'
'    frmPFPSDM!mnuResultsItem(0).Enabled = False
'    frmPFPSDM!mnuResultsItem(1).Enabled = False
'    frmPFPSDM!mnuResultsItem(2).Enabled = False
'    frmPFPSDM!mnuResultsItem(3).Enabled = False
'    frmPFPSDM!mnuResultsItem(4).Enabled = False
'
'    frmPFPSDM!mnuOptionsItem(0).Enabled = False
'    frmPFPSDM!mnuOptionsItem(1).Enabled = False
'    frmPFPSDM!mnuOptionsItem(2).Enabled = False
'  End If
'  frmPFPSDM.Caption = AppName_For_Display_long & "  -  " & Trim$(Filename)
'
'  Exit Sub
'
''if file from Version 1.00, NData_Points and Number_Component are not available and set to 0.
'No_Data_Points:
'  NData_Points = 0
'  Number_Component = 0
'  Select Case NVersion_Temp
'    Case 1#
'      Resume Read_Data_Points_v1_00
'    '     .
'    '     .
'    '     .
'    '.... other versions ...
'    '     .
'    '     .
'    '     .
'  End Select
'  Resume Read_Data_Points_v_LATEST
'
'OpenError:
'  MsgBox msg & Chr$(10) & "Error occurred trying to open file, please retry.", MB_ICONEXCLAMATION, AppName_For_Display_long
'  msg = ""
'  Close (f)
'  Resume Exit_Open
'
'Exit_Open:
'
'End Sub


