Option Strict Off
Option Explicit On
Module FileIO_Legacy2
	
	Const STR_EOF_MARKER As String = "12345-EOF-MARKER-12345"
	
	
	
	
	
	Const FileIO_Latest_declarations_end As Boolean = True
	
	
	Sub Bed_ProjectFile_Read(ByRef f As Short, ByRef be As BedPropertyType)
		Call ProjectFile_Read(f, be.length, "be.Length")
		Call ProjectFile_Read(f, be.Diameter, "be.Diameter")
		Call ProjectFile_Read(f, be.Weight, "be.Weight")
		Call ProjectFile_Read(f, be.Flowrate, "be.Flowrate")
		Call ProjectFile_Read(f, be.WaterDensity, "be.WaterDensity")
		Call ProjectFile_Read(f, be.WaterViscosity, "be.WaterViscosity")
		Call ProjectFile_Read(f, be.Temperature, "be.Temperature")
		Call ProjectFile_Read(f, be.Pressure, "be.Pressure")
		Call ProjectFile_Read(f, be.Phase, "be.Phase")
		Call ProjectFile_Read(f, be.NumberOfBeds, "be.NumberOfBeds")
		Call ProjectFile_Read(f, be.Water_Correlation.Name, "be.Water_Correlation.Name")
		Call ProjectFile_Read(f, be.Water_Correlation.Coeff(1), "be.Water_Correlation.Coeff(1)")
		Call ProjectFile_Read(f, be.Water_Correlation.Coeff(2), "be.Water_Correlation.Coeff(2)")
		Call ProjectFile_Read(f, be.Water_Correlation.Coeff(3), "be.Water_Correlation.Coeff(3)")
		Call ProjectFile_Read(f, be.Water_Correlation.Coeff(4), "be.Water_Correlation.Coeff(4)")
	End Sub
	Sub Bed_ProjectFile_Write(ByRef f As Short, ByRef be As BedPropertyType)
		Call ProjectFile_Write(f, be.length, "be.Length")
		Call ProjectFile_Write(f, be.Diameter, "be.Diameter")
		Call ProjectFile_Write(f, be.Weight, "be.Weight")
		Call ProjectFile_Write(f, be.Flowrate, "be.Flowrate")
		Call ProjectFile_Write(f, be.WaterDensity, "be.WaterDensity")
		Call ProjectFile_Write(f, be.WaterViscosity, "be.WaterViscosity")
		Call ProjectFile_Write(f, be.Temperature, "be.Temperature")
		Call ProjectFile_Write(f, be.Pressure, "be.Pressure")
		Call ProjectFile_Write(f, be.Phase, "be.Phase")
		Call ProjectFile_Write(f, be.NumberOfBeds, "be.NumberOfBeds")
		Call ProjectFile_Write(f, be.Water_Correlation.Name, "be.Water_Correlation.Name")
		Call ProjectFile_Write(f, be.Water_Correlation.Coeff(1), "be.Water_Correlation.Coeff(1)")
		Call ProjectFile_Write(f, be.Water_Correlation.Coeff(2), "be.Water_Correlation.Coeff(2)")
		Call ProjectFile_Write(f, be.Water_Correlation.Coeff(3), "be.Water_Correlation.Coeff(3)")
		Call ProjectFile_Write(f, be.Water_Correlation.Coeff(4), "be.Water_Correlation.Coeff(4)")
	End Sub
	
	
	Sub Component_ProjectFile_Read(ByRef f As Short, ByRef co As ComponentPropertyType)
		Call ProjectFile_Read(f, co.Name, "co.Name")
		Call ProjectFile_Read(f, co.CAS, "co.CAS")
		Call ProjectFile_Read(f, co.MW, "co.MW")
		Call ProjectFile_Read(f, co.InitialConcentration, "co.InitialConcentration")
		Call ProjectFile_Read(f, co.MolarVolume, "co.MolarVolume")
		Call ProjectFile_Read(f, co.BP, "co.BP")
		Call ProjectFile_Read(f, co.Use_K, "co.Use_K")
		Call ProjectFile_Read(f, co.Use_OneOverN, "co.Use_OneOverN")
		Call ProjectFile_Read(f, co.Liquid_Density, "co.Liquid_Density")
		Call ProjectFile_Read(f, co.Aqueous_Solubility, "co.Aqueous_Solubility")
		Call ProjectFile_Read(f, co.Vapor_Pressure, "co.Vapor_Pressure")
		Call ProjectFile_Read(f, co.Refractive_Index, "co.Refractive_Index")
		Call ProjectFile_Read(f, co.SPDFR, "co.SPDFR")
		Call ProjectFile_Read(f, co.SPDFR_Low_Concentration, "co.SPDFR_Low_Concentration")
		Call ProjectFile_Read(f, co.Use_SPDFR_Correlation, "co.Use_SPDFR_Correlation")
		Call ProjectFile_Read(f, co.kf, "co.kf")
		Call ProjectFile_Read(f, co.Ds, "co.Ds")
		Call ProjectFile_Read(f, co.Dp, "co.Dp")
		Call ProjectFile_Read(f, co.Corr(1), "co.Corr(1)")
		Call ProjectFile_Read(f, co.Corr(2), "co.Corr(2)")
		Call ProjectFile_Read(f, co.Corr(3), "co.Corr(3)")
		Call ProjectFile_Read(f, co.KP_User_Input(1), "co.KP_User_Input(1)")
		Call ProjectFile_Read(f, co.KP_User_Input(2), "co.KP_User_Input(2)")
		Call ProjectFile_Read(f, co.KP_User_Input(3), "co.KP_User_Input(3)")
		Call ProjectFile_Read(f, co.K_Reduction, "co.K_Reduction")
		Call ProjectFile_Read(f, co.Correlation.Name, "co.Correlation.Name")
		Call ProjectFile_Read(f, co.Correlation.Coeff(1), "co.Correlation.Coeff(1)")
		Call ProjectFile_Read(f, co.Correlation.Coeff(2), "co.Correlation.Coeff(2)")
		Call ProjectFile_Read(f, co.IsothermDB_Component_Name, "co.IsothermDB_Component_Name")
		Call ProjectFile_Read(f, co.IsothermDB_Range_Num, "co.IsothermDB_Range_Num")
		Call ProjectFile_Read(f, co.IPES_OrderOfMagnitude, "co.IPES_OrderOfMagnitude")
		Call ProjectFile_Read(f, co.IPES_NumRegressionPts, "co.IPES_NumRegressionPts")
		Call ProjectFile_Read(f, co.IPES_RelativeHumidity, "co.IPES_RelativeHumidity")
		Call ProjectFile_Read(f, co.IPES_EstimationMethod, "co.IPES_EstimationMethod")
		Call ProjectFile_Read(f, co.Source_KandOneOverN, "co.Source_KandOneOverN")
		Call ProjectFile_Read(f, co.IsothermDB_K, "co.IsothermDB_K")
		Call ProjectFile_Read(f, co.IsothermDB_OneOverN, "co.IsothermDB_OneOverN")
		Call ProjectFile_Read(f, co.IPESResult_K, "co.IPESResult_K")
		Call ProjectFile_Read(f, co.IPESResult_OneOverN, "co.IPESResult_OneOverN")
		Call ProjectFile_Read(f, co.UserEntered_K, "co.UserEntered_K")
		Call ProjectFile_Read(f, co.UserEntered_OneOverN, "co.UserEntered_OneOverN")
		Call ProjectFile_Read(f, co.Tortuosity, "co.tortuosity")
		Call ProjectFile_Read(f, co.Use_Tortuosity_Correlation, "co.Use_Tortuosity_Correlation")
		Call ProjectFile_Read(f, co.Constant_Tortuosity, "co.Constant_Tortuosity")
	End Sub
	Sub Component_ProjectFile_Write(ByRef f As Short, ByRef co As ComponentPropertyType)
		Call ProjectFile_Write(f, co.Name, "co.Name")
		Call ProjectFile_Write(f, co.CAS, "co.CAS")
		Call ProjectFile_Write(f, co.MW, "co.MW")
		Call ProjectFile_Write(f, co.InitialConcentration, "co.InitialConcentration")
		Call ProjectFile_Write(f, co.MolarVolume, "co.MolarVolume")
		Call ProjectFile_Write(f, co.BP, "co.BP")
		Call ProjectFile_Write(f, co.Use_K, "co.Use_K")
		Call ProjectFile_Write(f, co.Use_OneOverN, "co.Use_OneOverN")
		Call ProjectFile_Write(f, co.Liquid_Density, "co.Liquid_Density")
		Call ProjectFile_Write(f, co.Aqueous_Solubility, "co.Aqueous_Solubility")
		Call ProjectFile_Write(f, co.Vapor_Pressure, "co.Vapor_Pressure")
		Call ProjectFile_Write(f, co.Refractive_Index, "co.Refractive_Index")
		Call ProjectFile_Write(f, co.SPDFR, "co.SPDFR")
		Call ProjectFile_Write(f, co.SPDFR_Low_Concentration, "co.SPDFR_Low_Concentration")
		Call ProjectFile_Write(f, co.Use_SPDFR_Correlation, "co.Use_SPDFR_Correlation")
		Call ProjectFile_Write(f, co.kf, "co.kf")
		Call ProjectFile_Write(f, co.Ds, "co.Ds")
		Call ProjectFile_Write(f, co.Dp, "co.Dp")
		Call ProjectFile_Write(f, co.Corr(1), "co.Corr(1)")
		Call ProjectFile_Write(f, co.Corr(2), "co.Corr(2)")
		Call ProjectFile_Write(f, co.Corr(3), "co.Corr(3)")
		Call ProjectFile_Write(f, co.KP_User_Input(1), "co.KP_User_Input(1)")
		Call ProjectFile_Write(f, co.KP_User_Input(2), "co.KP_User_Input(2)")
		Call ProjectFile_Write(f, co.KP_User_Input(3), "co.KP_User_Input(3)")
		Call ProjectFile_Write(f, co.K_Reduction, "co.K_Reduction")
		Call ProjectFile_Write(f, co.Correlation.Name, "co.Correlation.Name")
		Call ProjectFile_Write(f, co.Correlation.Coeff(1), "co.Correlation.Coeff(1)")
		Call ProjectFile_Write(f, co.Correlation.Coeff(2), "co.Correlation.Coeff(2)")
		Call ProjectFile_Write(f, co.IsothermDB_Component_Name, "co.IsothermDB_Component_Name")
		Call ProjectFile_Write(f, co.IsothermDB_Range_Num, "co.IsothermDB_Range_Num")
		Call ProjectFile_Write(f, co.IPES_OrderOfMagnitude, "co.IPES_OrderOfMagnitude")
		Call ProjectFile_Write(f, co.IPES_NumRegressionPts, "co.IPES_NumRegressionPts")
		Call ProjectFile_Write(f, co.IPES_RelativeHumidity, "co.IPES_RelativeHumidity")
		Call ProjectFile_Write(f, co.IPES_EstimationMethod, "co.IPES_EstimationMethod")
		Call ProjectFile_Write(f, co.Source_KandOneOverN, "co.Source_KandOneOverN")
		Call ProjectFile_Write(f, co.IsothermDB_K, "co.IsothermDB_K")
		Call ProjectFile_Write(f, co.IsothermDB_OneOverN, "co.IsothermDB_OneOverN")
		Call ProjectFile_Write(f, co.IPESResult_K, "co.IPESResult_K")
		Call ProjectFile_Write(f, co.IPESResult_OneOverN, "co.IPESResult_OneOverN")
		Call ProjectFile_Write(f, co.UserEntered_K, "co.UserEntered_K")
		Call ProjectFile_Write(f, co.UserEntered_OneOverN, "co.UserEntered_OneOverN")
		Call ProjectFile_Write(f, co.Tortuosity, "co.tortuosity")
		Call ProjectFile_Write(f, co.Use_Tortuosity_Correlation, "co.Use_Tortuosity_Correlation")
		Call ProjectFile_Write(f, co.Constant_Tortuosity, "co.Constant_Tortuosity")
	End Sub
	
	
	'RETURNS:
	'- true = open went okay.
	'- false = open failed.
	''''Function File_Open_Latest_v1_40(f As Integer) As Boolean
	Function File_Open_Legacy_v1_40(ByRef f As Short) As Boolean
		Dim i As Short
		Dim J As Short
		Dim Test_STR_EOF_MARKER As String
		Call ProjectFile_Read(f, FileNote, "FileNote")
		Call ProjectFile_Read(f, Number_Component, "Number_Component")
		For i = 1 To Number_Component
			Call Component_ProjectFile_Read(f, Component(i))
		Next i
		Call Bed_ProjectFile_Read(f, Bed)
		Call UnitsOfDisplay_ProjectFile_Read(f)
		'MISCELLANEOUS BLOCK.
		Call ProjectFile_Read(f, Carbon.Name, "Carbon.Name")
		Call ProjectFile_Read(f, Carbon.Porosity, "Carbon.Porosity")
		Call ProjectFile_Read(f, Carbon.Density, "Carbon.Density")
		Call ProjectFile_Read(f, Carbon.ParticleRadius, "Carbon.ParticleRadius")
		Call ProjectFile_Read(f, Carbon.Tortuosity, "Carbon.tortuosity")
		Call ProjectFile_Read(f, Carbon.W0, "Carbon.W0")
		Call ProjectFile_Read(f, Carbon.BB, "Carbon.BB")
		Call ProjectFile_Read(f, Carbon.PolanyiExponent, "Carbon.PolanyiExponent")
		Call ProjectFile_Read(f, State_Check_Water(1), "State_Check_Water(1)")
		Call ProjectFile_Read(f, State_Check_Water(2), "State_Check_Water(2)")
		Call ProjectFile_Read(f, Carbon.ShapeFactor, "Carbon.ShapeFactor")
		Call ProjectFile_Read(f, Constant_Tortuosity, "Constant_Tortuosity")
		Call ProjectFile_Read(f, Carbon.ShapeFactor, "Carbon.ShapeFactor")
		Call ProjectFile_Read(f, NC, "NC")
		Call ProjectFile_Read(f, MC, "MC")
		Call ProjectFile_Read(f, TimeP.Init, "TimeP.Init")
		Call ProjectFile_Read(f, TimeP.End_Renamed, "TimeP.End")
		Call ProjectFile_Read(f, TimeP.np, "TimeP.np")
		Call ProjectFile_Read(f, TimeP.Step_Renamed, "TimeP.Step")
		'INFLUENT POINTS.
		Call ProjectFile_Read(f, Number_Influent_Points, "Number_Influent_Points")
		If (Number_Influent_Points > 0) Then
			For i = 1 To Number_Influent_Points
				Input(f, T_Influent(i))
				For J = 1 To Number_Component
					Input(f, C_Influent(J, i))
				Next J
			Next i
		End If
		'EFFLUENT POINTS.
		Call ProjectFile_Read(f, NData_Points, "NData_Points")
		If (NData_Points > 0) Then
			For i = 1 To NData_Points
				Input(f, T_Data_Points(i))
				For J = 1 To Number_Component
					Input(f, C_Data_Points(J, i))
				Next J
			Next i
		End If
		Call ProjectFile_Read(f, Test_STR_EOF_MARKER, "STR_EOF_MARKER")
		If (Test_STR_EOF_MARKER <> STR_EOF_MARKER) Then
			Call Show_Error("This file has an invalid file format.")
			File_Open_Legacy_v1_40 = False
		Else
			File_Open_Legacy_v1_40 = True
		End If
	End Function
	Sub File_Save_Legacy_v1_40(ByRef f As Short)
		Dim i As Short
		Dim J As Short
		PrintLine(f, 1.4)
		Call ProjectFile_Write(f, FileNote, "FileNote")
		Call ProjectFile_Write(f, Number_Component, "Number_Component")
		For i = 1 To Number_Component
			Call Component_ProjectFile_Write(f, Component(i))
		Next i
		Call Bed_ProjectFile_Write(f, Bed)
		Call UnitsOfDisplay_ProjectFile_Write(f)
		'MISCELLANEOUS BLOCK.
		Call ProjectFile_Write(f, Carbon.Name, "Carbon.Name")
		Call ProjectFile_Write(f, Carbon.Porosity, "Carbon.Porosity")
		Call ProjectFile_Write(f, Carbon.Density, "Carbon.Density")
		Call ProjectFile_Write(f, Carbon.ParticleRadius, "Carbon.ParticleRadius")
		Call ProjectFile_Write(f, Carbon.Tortuosity, "Carbon.tortuosity")
		Call ProjectFile_Write(f, Carbon.W0, "Carbon.W0")
		Call ProjectFile_Write(f, Carbon.BB, "Carbon.BB")
		Call ProjectFile_Write(f, Carbon.PolanyiExponent, "Carbon.PolanyiExponent")
		Call ProjectFile_Write(f, State_Check_Water(1), "State_Check_Water(1)")
		Call ProjectFile_Write(f, State_Check_Water(2), "State_Check_Water(2)")
		Call ProjectFile_Write(f, Carbon.ShapeFactor, "Carbon.ShapeFactor")
		Call ProjectFile_Write(f, Constant_Tortuosity, "Constant_Tortuosity")
		Call ProjectFile_Write(f, Carbon.ShapeFactor, "Carbon.ShapeFactor")
		Call ProjectFile_Write(f, NC, "NC")
		Call ProjectFile_Write(f, MC, "MC")
		Call ProjectFile_Write(f, TimeP.Init, "TimeP.Init")
		Call ProjectFile_Write(f, TimeP.End_Renamed, "TimeP.End")
		Call ProjectFile_Write(f, TimeP.np, "TimeP.np")
		Call ProjectFile_Write(f, TimeP.Step_Renamed, "TimeP.Step")
		'INFLUENT POINTS.
		Call ProjectFile_Write(f, Number_Influent_Points, "Number_Influent_Points")
		If (Number_Influent_Points > 0) Then
			For i = 1 To Number_Influent_Points
				WriteLine(f, T_Influent(i))
				For J = 1 To Number_Component
					WriteLine(f, C_Influent(J, i))
				Next J
			Next i
		End If
		'EFFLUENT POINTS.
		Call ProjectFile_Write(f, NData_Points, "NData_Points")
		If (NData_Points > 0) Then
			For i = 1 To NData_Points
				WriteLine(f, T_Data_Points(i))
				For J = 1 To Number_Component
					WriteLine(f, C_Data_Points(J, i))
				Next J
			Next i
		End If
		Call ProjectFile_Write(f, STR_EOF_MARKER, "STR_EOF_MARKER")
	End Sub
	
	
	Sub Units1_ProjectFile_Read(ByRef f As Short, ByRef CboX As System.Windows.Forms.Control, ByRef Desc As String)
		Dim TxtX As System.Windows.Forms.Control
		Dim InLine As String
		Dim Dummy1 As String
		Dim NewUnits As String
		Dim H As Short
		Call ProjectFile_Read(f, InLine, Dummy1)
		NewUnits = InLine
		'UPGRADE_WARNING: Couldn't resolve default property of object unitsys_lookup_cbox(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		H = unitsys_lookup_cbox(CboX)
		TxtX = unitsys(H).TxtX
		Call unitsys_set_units(TxtX, NewUnits)
	End Sub
	Sub Units1_ProjectFile_Write(ByRef f As Short, ByRef CboX As ComboBox, ByRef Desc As String)
		Dim OutStr As String
		'UPGRADE_WARNING: Couldn't resolve default property of object CboX.ListIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (CboX.SelectedIndex >= 0) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object CboX.ListIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object CboX.List. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			OutStr = CboX.Items(CboX.SelectedIndex)
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object CboX.ListCount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If (CboX.Items.Count > 0) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object CboX.List. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				OutStr = CboX.Items(0)
			Else
				OutStr = "" 'NOT LIKELY TO GET HERE!
			End If
		End If
		Call ProjectFile_Write(f, OutStr, Desc)
	End Sub


	Sub UnitsOfDisplay_ProjectFile_Read(ByRef f As Short)
		Call Units1_ProjectFile_Read(f, frmMain.txtBedUnits(0), "frmMain.txtBedUnits(0)")
		Call Units1_ProjectFile_Read(f, frmMain.txtBedUnits(1), "frmMain.txtBedUnits(1)")
		Call Units1_ProjectFile_Read(f, frmMain.txtBedUnits(2), "frmMain.txtBedUnits(2)")
		Call Units1_ProjectFile_Read(f, frmMain.txtBedUnits(3), "frmMain.txtBedUnits(3)")
		Call Units1_ProjectFile_Read(f, frmMain.txtBedUnits(4), "frmMain.txtBedUnits(4)")
		Call Units1_ProjectFile_Read(f, frmMain.txtCarbonUnits(1), "frmMain.txtCarbonUnits(1)")
		Call Units1_ProjectFile_Read(f, frmMain.txtCarbonUnits(2), "frmMain.txtCarbonUnits(2)")
		Call Units1_ProjectFile_Read(f, frmMain.txtTimeUnits(0), "frmMain.txtTimeUnits(0)")
		Call Units1_ProjectFile_Read(f, frmMain.txtTimeUnits(1), "frmMain.txtTimeUnits(1)")
		Call Units1_ProjectFile_Read(f, frmMain.txtTimeUnits(2), "frmMain.txtTimeUnits(2)")
		Call ProjectFile_Read(f, PropertyUnits.MW, "PropertyUnits.MW")
		Call ProjectFile_Read(f, PropertyUnits.MolarVolume, "PropertyUnits.MolarVolume")
		Call ProjectFile_Read(f, PropertyUnits.BP, "PropertyUnits.BP")
		Call ProjectFile_Read(f, PropertyUnits.InitialConcentration, "PropertyUnits.InitialConcentration")
		Call ProjectFile_Read(f, PropertyUnits.Liquid_Density, "PropertyUnits.Liquid_Density")
		Call ProjectFile_Read(f, PropertyUnits.Aqueous_Solubility, "PropertyUnits.Aqueous_Solubility")
		Call ProjectFile_Read(f, PropertyUnits.Vapor_Pressure, "PropertyUnits.Vapor_Pressure")
		Call ProjectFile_Read(f, PropertyUnits.k, "PropertyUnits.k")
	End Sub
	Sub UnitsOfDisplay_ProjectFile_Write(ByRef f As Short)
		Call Units1_ProjectFile_Write(f, frmMain.txtBedUnits(0), "frmMain.txtBedUnits(0)")
		Call Units1_ProjectFile_Write(f, frmMain.txtBedUnits(1), "frmMain.txtBedUnits(1)")
		Call Units1_ProjectFile_Write(f, frmMain.txtBedUnits(2), "frmMain.txtBedUnits(2)")
		Call Units1_ProjectFile_Write(f, frmMain.txtBedUnits(3), "frmMain.txtBedUnits(3)")
		Call Units1_ProjectFile_Write(f, frmMain.txtBedUnits(4), "frmMain.txtBedUnits(4)")
		Call Units1_ProjectFile_Write(f, frmMain.txtCarbonUnits(1), "frmMain.txtCarbonUnits(1)")
		Call Units1_ProjectFile_Write(f, frmMain.txtCarbonUnits(2), "frmMain.txtCarbonUnits(2)")
		Call Units1_ProjectFile_Write(f, frmMain.txtTimeUnits(0), "frmMain.txtTimeUnits(0)")
		Call Units1_ProjectFile_Write(f, frmMain.txtTimeUnits(1), "frmMain.txtTimeUnits(1)")
		Call Units1_ProjectFile_Write(f, frmMain.txtTimeUnits(2), "frmMain.txtTimeUnits(2)")
		Call ProjectFile_Write(f, PropertyUnits.MW, "PropertyUnits.MW")
		Call ProjectFile_Write(f, PropertyUnits.MolarVolume, "PropertyUnits.MolarVolume")
		Call ProjectFile_Write(f, PropertyUnits.BP, "PropertyUnits.BP")
		Call ProjectFile_Write(f, PropertyUnits.InitialConcentration, "PropertyUnits.InitialConcentration")
		Call ProjectFile_Write(f, PropertyUnits.Liquid_Density, "PropertyUnits.Liquid_Density")
		Call ProjectFile_Write(f, PropertyUnits.Aqueous_Solubility, "PropertyUnits.Aqueous_Solubility")
		Call ProjectFile_Write(f, PropertyUnits.Vapor_Pressure, "PropertyUnits.Vapor_Pressure")
		Call ProjectFile_Write(f, PropertyUnits.k, "PropertyUnits.k")
	End Sub
End Module