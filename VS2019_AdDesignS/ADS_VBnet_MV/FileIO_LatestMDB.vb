Option Strict Off
Option Explicit On
Module FileIO_LatestMDB


    Const FileIO_LatestMDB_declarations_end As Boolean = True


    'RETURNS:
    '         TRUE = SUCCEEDED IN LOADING.
    '         FALSE = FAILED IN LOADING.
    Function File_Open_Latest_v1_60(ByRef fn_This As String) As Boolean
		'	Dim OpenDatabase As Object
		'UPGRADE_ISSUE: Workspace object was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim Ws1 As dao.Workspace
		'UPGRADE_ISSUE: Database object was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim Db1 As dao.Database
		'UPGRADE_ISSUE: Recordset object was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim Rs1 As dao.Recordset
		Dim Use_FieldIndex As Short
		Dim Use_FieldIndex2 As Short
		Dim ContainsTable_PSDMInRoomData As Boolean
		'UPGRADE_WARNING: Arrays in structure rp may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim rp As RoomParam_Type
		
		'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
		'>>>>>>>>>>>>>>>>>>>>>>>>>>>  INPUT FROM MAIN DATABASE  <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
		'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
		If (Not FileExists(fn_This)) Then
			'ERROR: UNABLE TO FIND THE FILE!
			File_Open_Latest_v1_60 = False
			Exit Function
		End If
		'OPEN DATABASE.
		Db1 = DAOEngine.OpenDatabase(fn_This)  'From Ws1 to Daoengine ??? Shang

		'=========== INPUT DATA FROM DATABASE TABLES. =================

		'------ INPUT DATA FROM TABLE "Version". ------------------------------------------------------------------------------------------------------
		'APPLICABLE DEFAULT VALUES:
		ContainsTable_PSDMInRoomData = False
		'UPGRADE_WARNING: Untranslated statement in File_Open_Latest_v1_60. Please check source code.
		
		'------ INPUT DATA FROM TABLE "Main". ------------------------------------------------------------------------------------------------------
		Dim booDoDemoCalc As Boolean
		Dim dblDemoChecksum As Double
		Dim dblThisVal As Double
		Dim lngThisVal As Integer
		booDoDemoCalc = False
		dblDemoChecksum = 0#
		If (IsThisADemo() = True) Then
			' STORE TO VARIABLE TO SAVE A LITTLE TIME.
			booDoDemoCalc = True
		End If

        'UPGRADE_WARNING: Untranslated statement in File_Open_Latest_v1_60. Please check source code.
        If (Database_IsTableExist(Db1, "Main") = False) Then
            'DO NOTHING: USE DEFAULT VALUES.
        Else
            Rs1 = Db1.OpenRecordset("Main")
			' If (Database_NoRecordsInRecordset(Rs1)) Then
			If (Rs1.RecordCount = 0) Then
				'DO NOTHING: USE DEFAULT VALUES.
			Else
				Rs1.MoveFirst()

                Do Until Rs1.EOF
                    Use_FieldIndex = CInt(Database_Get_Long(Rs1, "FieldIndex"))
					If (booDoDemoCalc = True) Then
						Call Database_LoadProperty(Rs1, dblThisVal)
						Call Database_LoadProperty(Rs1, lngThisVal)
						If (dblThisVal = 0#) Then dblThisVal = CDbl(lngThisVal)
						If (dblThisVal = 0#) Then dblThisVal = 0.1
						dblDemoChecksum = dblDemoChecksum + Math.Abs(Math.Log(Math.Abs(dblThisVal)))
					End If

					If (Component(Use_FieldIndex).Correlation.Coeff Is Nothing Or
						Component(Use_FieldIndex).Corr Is Nothing Or
						Component(Use_FieldIndex).KP_User_Input Is Nothing) Then   'Shang   initialization is needed

						Component(Use_FieldIndex).Initialize()
					End If

					If (Component(Use_FieldIndex).Name Is Nothing) Then   'Shang
						Component(Use_FieldIndex).Name = ""
					End If

					Select Case Trim$(UCase$(Database_Get_String(Rs1, "FieldName")))
          'HEADER BLOCK.
                        Case Trim$(UCase$("FileNote")) : Call Database_LoadProperty(Rs1, FileNote, True)
                        Case Trim$(UCase$("Number_Component")) : Call Database_LoadProperty(Rs1, Number_Component)
          'COMPONENT PROPERTIES BLOCK.
                        Case Trim$(UCase$("Co.Name")) : Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).Name)
                        Case Trim$(UCase$("Co.CAS")) : Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).CAS)
                        Case Trim$(UCase$("Co.MW")) : Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).MW)
                        Case Trim$(UCase$("Co.InitialConcentration")) : Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).InitialConcentration)
                        Case Trim$(UCase$("Co.MolarVolume")) : Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).MolarVolume)
                        Case Trim$(UCase$("Co.BP")) : Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).BP)
                        Case Trim$(UCase$("Co.Use_K")) : Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).Use_K)
                        Case Trim$(UCase$("Co.Use_OneOverN")) : Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).Use_OneOverN)
                        Case Trim$(UCase$("Co.Liquid_Density")) : Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).Liquid_Density)
                        Case Trim$(UCase$("Co.Aqueous_Solubility")) : Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).Aqueous_Solubility)
                        Case Trim$(UCase$("Co.Vapor_Pressure")) : Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).Vapor_Pressure)
                        Case Trim$(UCase$("Co.Refractive_Index")) : Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).Refractive_Index)
                        Case Trim$(UCase$("Co.SPDFR")) : Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).SPDFR)
                        Case Trim$(UCase$("Co.SPDFR_Low_Concentration")) : Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).SPDFR_Low_Concentration)
                        Case Trim$(UCase$("Co.Use_SPDFR_Correlation")) : Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).Use_SPDFR_Correlation)
                        Case Trim$(UCase$("Co.kf")) : Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).kf)
                        Case Trim$(UCase$("Co.Ds")) : Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).Ds)
                        Case Trim$(UCase$("Co.Dp")) : Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).Dp)
						Case Trim$(UCase$("Co.Corr(1)"))
							Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).Corr(1))
						Case Trim$(UCase$("Co.Corr(2)"))
							Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).Corr(2))
						Case Trim$(UCase$("Co.Corr(3)"))
							Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).Corr(3))
						Case Trim$(UCase$("Co.KP_User_Input(1)")) : Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).KP_User_Input(1))
                        Case Trim$(UCase$("Co.KP_User_Input(2)")) : Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).KP_User_Input(2))
                        Case Trim$(UCase$("Co.KP_User_Input(3)")) : Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).KP_User_Input(3))
                        Case Trim$(UCase$("Co.K_Reduction")) : Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).K_Reduction)
                        Case Trim$(UCase$("Co.Correlation.Name")) : Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).Correlation.Name)
                        Case Trim$(UCase$("Co.Correlation.Coeff(1)")) : Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).Correlation.Coeff(1))
                        Case Trim$(UCase$("Co.Correlation.Coeff(2)")) : Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).Correlation.Coeff(2))
                        Case Trim$(UCase$("Co.IsothermDB_Component_Name")) : Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).IsothermDB_Component_Name)
                        Case Trim$(UCase$("Co.IsothermDB_Range_Num")) : Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).IsothermDB_Range_Num)
                        Case Trim$(UCase$("Co.IPES_OrderOfMagnitude")) : Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).IPES_OrderOfMagnitude)
                        Case Trim$(UCase$("Co.IPES_NumRegressionPts")) : Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).IPES_NumRegressionPts)
                        Case Trim$(UCase$("Co.IPES_RelativeHumidity")) : Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).IPES_RelativeHumidity)
                        Case Trim$(UCase$("Co.IPES_EstimationMethod")) : Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).IPES_EstimationMethod)
                        Case Trim$(UCase$("Co.Source_KandOneOverN")) : Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).Source_KandOneOverN)
                        Case Trim$(UCase$("Co.IsothermDB_K")) : Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).IsothermDB_K)
                        Case Trim$(UCase$("Co.IsothermDB_OneOverN")) : Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).IsothermDB_OneOverN)
                        Case Trim$(UCase$("Co.IPESResult_K")) : Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).IPESResult_K)
                        Case Trim$(UCase$("Co.IPESResult_OneOverN")) : Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).IPESResult_OneOverN)
                        Case Trim$(UCase$("Co.UserEntered_K")) : Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).UserEntered_K)
                        Case Trim$(UCase$("Co.UserEntered_OneOverN")) : Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).UserEntered_OneOverN)
                        Case Trim$(UCase$("Co.Tortuosity")) : Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).Tortuosity)
                        Case Trim$(UCase$("Co.Use_Tortuosity_Correlation")) : Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).Use_Tortuosity_Correlation)
                        Case Trim$(UCase$("Co.Constant_Tortuosity")) : Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).Constant_Tortuosity)
          'BED PROPERTIES BLOCK.
                        Case Trim$(UCase$("Be.length")) : Call Database_LoadProperty(Rs1, Bed.length)
                        Case Trim$(UCase$("Be.Diameter")) : Call Database_LoadProperty(Rs1, Bed.Diameter)
                        Case Trim$(UCase$("Be.Weight")) : Call Database_LoadProperty(Rs1, Bed.Weight)
                        Case Trim$(UCase$("Be.Flowrate")) : Call Database_LoadProperty(Rs1, Bed.Flowrate)
                        Case Trim$(UCase$("Be.WaterDensity")) : Call Database_LoadProperty(Rs1, Bed.WaterDensity)
                        Case Trim$(UCase$("Be.WaterViscosity")) : Call Database_LoadProperty(Rs1, Bed.WaterViscosity)
                        Case Trim$(UCase$("Be.Temperature")) : Call Database_LoadProperty(Rs1, Bed.Temperature)
                        Case Trim$(UCase$("Be.Pressure")) : Call Database_LoadProperty(Rs1, Bed.Pressure)
                        Case Trim$(UCase$("Be.Phase")) : Call Database_LoadProperty(Rs1, Bed.Phase)
                        Case Trim$(UCase$("Be.NumberOfBeds")) : Call Database_LoadProperty(Rs1, Bed.NumberOfBeds)
                        Case Trim$(UCase$("Be.Water_Correlation.Name")) : Call Database_LoadProperty(Rs1, Bed.Water_Correlation.Name)
                        Case Trim$(UCase$("Be.Water_Correlation.Coeff(1)")) : Call Database_LoadProperty(Rs1, Bed.Water_Correlation.Coeff(1))
                        Case Trim$(UCase$("Be.Water_Correlation.Coeff(2)")) : Call Database_LoadProperty(Rs1, Bed.Water_Correlation.Coeff(2))
                        Case Trim$(UCase$("Be.Water_Correlation.Coeff(3)")) : Call Database_LoadProperty(Rs1, Bed.Water_Correlation.Coeff(3))
                        Case Trim$(UCase$("Be.Water_Correlation.Coeff(4)")) : Call Database_LoadProperty(Rs1, Bed.Water_Correlation.Coeff(4))
          'UNITS BLOCK.
                        Case Trim$(UCase$("frmMain.txtBedUnits(0)")) : Call Units1_Database_LoadProperty(Rs1, frmMain.txtBedUnits(0))
                        Case Trim$(UCase$("frmMain.txtBedUnits(1)")) : Call Units1_Database_LoadProperty(Rs1, frmMain.txtBedUnits(1))
                        Case Trim$(UCase$("frmMain.txtBedUnits(2)")) : Call Units1_Database_LoadProperty(Rs1, frmMain.txtBedUnits(2))
                        Case Trim$(UCase$("frmMain.txtBedUnits(3)")) : Call Units1_Database_LoadProperty(Rs1, frmMain.txtBedUnits(3))
                        Case Trim$(UCase$("frmMain.txtBedUnits(4)")) : Call Units1_Database_LoadProperty(Rs1, frmMain.txtBedUnits(4))
                        Case Trim$(UCase$("frmMain.txtCarbonUnits(1)")) : Call Units1_Database_LoadProperty(Rs1, frmMain.txtCarbonUnits(1))
                        Case Trim$(UCase$("frmMain.txtCarbonUnits(2)")) : Call Units1_Database_LoadProperty(Rs1, frmMain.txtCarbonUnits(2))
                        Case Trim$(UCase$("frmMain.txtTimeUnits(0)")) : Call Units1_Database_LoadProperty(Rs1, frmMain.txtTimeUnits(0))
                        Case Trim$(UCase$("frmMain.txtTimeUnits(1)")) : Call Units1_Database_LoadProperty(Rs1, frmMain.txtTimeUnits(1))
                        Case Trim$(UCase$("frmMain.txtTimeUnits(2)")) : Call Units1_Database_LoadProperty(Rs1, frmMain.txtTimeUnits(2))
                        Case Trim$(UCase$("PropertyUnits.MW")) : Call Database_LoadProperty(Rs1, PropertyUnits.MW)
                        Case Trim$(UCase$("PropertyUnits.MolarVolume")) : Call Database_LoadProperty(Rs1, PropertyUnits.MolarVolume)
                        Case Trim$(UCase$("PropertyUnits.BP")) : Call Database_LoadProperty(Rs1, PropertyUnits.BP)
                        Case Trim$(UCase$("PropertyUnits.InitialConcentration")) : Call Database_LoadProperty(Rs1, PropertyUnits.InitialConcentration)
                        Case Trim$(UCase$("PropertyUnits.Liquid_Density")) : Call Database_LoadProperty(Rs1, PropertyUnits.Liquid_Density)
                        Case Trim$(UCase$("PropertyUnits.Aqueous_Solubility")) : Call Database_LoadProperty(Rs1, PropertyUnits.Aqueous_Solubility)
                        Case Trim$(UCase$("PropertyUnits.Vapor_Pressure")) : Call Database_LoadProperty(Rs1, PropertyUnits.Vapor_Pressure)
                        Case Trim$(UCase$("PropertyUnits.k")) : Call Database_LoadProperty(Rs1, PropertyUnits.k)
          'MISCELLANEOUS BLOCK.
                        Case Trim$(UCase$("Carbon.Name")) : Call Database_LoadProperty(Rs1, Carbon.Name)
                        Case Trim$(UCase$("Carbon.Porosity")) : Call Database_LoadProperty(Rs1, Carbon.Porosity)
                        Case Trim$(UCase$("Carbon.Density")) : Call Database_LoadProperty(Rs1, Carbon.Density)
                        Case Trim$(UCase$("Carbon.ParticleRadius")) : Call Database_LoadProperty(Rs1, Carbon.ParticleRadius)
                        Case Trim$(UCase$("Carbon.Tortuosity")) : Call Database_LoadProperty(Rs1, Carbon.Tortuosity)
                        Case Trim$(UCase$("Carbon.W0")) : Call Database_LoadProperty(Rs1, Carbon.W0)
                        Case Trim$(UCase$("Carbon.BB")) : Call Database_LoadProperty(Rs1, Carbon.BB)
                        Case Trim$(UCase$("Carbon.PolanyiExponent")) : Call Database_LoadProperty(Rs1, Carbon.PolanyiExponent)
                        Case Trim$(UCase$("State_Check_Water(1)")) : Call Database_LoadProperty(Rs1, State_Check_Water(1))
                        Case Trim$(UCase$("State_Check_Water(2)")) : Call Database_LoadProperty(Rs1, State_Check_Water(2))
                        Case Trim$(UCase$("Carbon.ShapeFactor")) : Call Database_LoadProperty(Rs1, Carbon.ShapeFactor)
                        Case Trim$(UCase$("Constant_Tortuosity")) : Call Database_LoadProperty(Rs1, Constant_Tortuosity)
                        Case Trim$(UCase$("Carbon.ShapeFactor")) : Call Database_LoadProperty(Rs1, Carbon.ShapeFactor)
                        Case Trim$(UCase$("NC")) : Call Database_LoadProperty(Rs1, NC)
                        Case Trim$(UCase$("MC")) : Call Database_LoadProperty(Rs1, MC)
						Case Trim$(UCase$("TimeP.Init")) : Call Database_LoadProperty(Rs1, TimeP.Init)
						Case Trim$(UCase$("TimeP.End")) : Call Database_LoadProperty(Rs1, TimeP.End_Renamed)
						Case Trim$(UCase$("TimeP.np")) : Call Database_LoadProperty(Rs1, TimeP.np)
						Case Trim$(UCase$("TimeP.Step")) : Call Database_LoadProperty(Rs1, TimeP.Step_Renamed)
		  'INFLUENT/EFFLUENT POINT COUNTS.
						Case Trim$(UCase$("Number_Influent_Points")) : Call Database_LoadProperty(Rs1, Number_Influent_Points)
                        Case Trim$(UCase$("NData_Points")) : Call Database_LoadProperty(Rs1, NData_Points)
                    End Select
                    Rs1.MoveNext()
                Loop
            End If
            Rs1.Close()
        End If
        If (booDoDemoCalc = True) Then
			''''Call Show_Message("Demo Version value of dblDemoChecksum = " & _
			Trim$(Str$(dblDemoChecksum) & ".")
			If (Demo_CheckForValidFile(dblDemoChecksum) = False) Then
                Call file_new()
                Call Demo_ShowError("In the demonstration version, only the example files may be opened.")
                File_Open_Latest_v1_60 = False
                Exit Function
            End If
        End If

        '------ INPUT DATA FROM TABLE "InfluentPoints". ------------------------------------------------------------------------------------------------------
        If (Database_IsTableExist(Db1, "InfluentPoints") = False) Then
            'DO NOTHING: USE DEFAULT VALUES.
        Else
            Rs1 = Db1.OpenRecordset("InfluentPoints")
			If (Rs1.RecordCount = 0) Then
				'DO NOTHING: USE DEFAULT VALUES.
			Else
				Rs1.MoveFirst()

                Do Until Rs1.EOF
                    Use_FieldIndex = CInt(Database_Get_Long(Rs1, "FieldIndex"))
                    Use_FieldIndex2 = CInt(Database_Get_Long(Rs1, "FieldIndex2"))
                    Select Case Trim$(UCase$(Database_Get_String(Rs1, "FieldName")))
                        Case Trim$(UCase$("T_Influent")) : Call Database_LoadProperty(Rs1, T_Influent(Use_FieldIndex))
                        Case Trim$(UCase$("C_Influent")) : Call Database_LoadProperty(Rs1, C_Influent(Use_FieldIndex, Use_FieldIndex2))
                    End Select
                    Rs1.MoveNext()
                Loop
            End If
            Rs1.Close()
        End If

        '------ INPUT DATA FROM TABLE "EffluentPoints". ------------------------------------------------------------------------------------------------------
        If (Database_IsTableExist(Db1, "EffluentPoints") = False) Then
            'DO NOTHING: USE DEFAULT VALUES.
        Else
            Rs1 = Db1.OpenRecordset("EffluentPoints")
			If (Rs1.RecordCount = 0) Then
				'DO NOTHING: USE DEFAULT VALUES.
			Else
				Rs1.MoveFirst()

                Do Until Rs1.EOF
                    Use_FieldIndex = CInt(Database_Get_Long(Rs1, "FieldIndex"))
                    Use_FieldIndex2 = CInt(Database_Get_Long(Rs1, "FieldIndex2"))
                    Select Case Trim$(UCase$(Database_Get_String(Rs1, "FieldName")))
                        Case Trim$(UCase$("T_Data_Points")) : Call Database_LoadProperty(Rs1, T_Data_Points(Use_FieldIndex))
                        Case Trim$(UCase$("C_Data_Points")) : Call Database_LoadProperty(Rs1, C_Data_Points(Use_FieldIndex, Use_FieldIndex2))
                    End Select
                    Rs1.MoveNext()
                Loop
            End If
            Rs1.Close()
        End If

        '------ INPUT DATA FROM TABLE "PSDMInRoomData". ------------------------------------------------------------------------------------------------------
        If (ContainsTable_PSDMInRoomData) Then
            rp = RoomParams
            If (Database_IsTableExist(Db1, "PSDMInRoomData") = False) Then
                'DO NOTHING: USE DEFAULT VALUES.
            Else
                Rs1 = Db1.OpenRecordset("PSDMInRoomData")
				If (Rs1.RecordCount = 0) Then
					'DO NOTHING: USE DEFAULT VALUES.
				Else
					Rs1.MoveFirst()

                    Do Until Rs1.EOF
                        Use_FieldIndex = CInt(Database_Get_Long(Rs1, "FieldIndex"))
                        Select Case Trim$(UCase$(Database_Get_String(Rs1, "FieldName")))
                            ' INPUT ROOM PARAMETERS.
                            Case Trim$(UCase$("RP.COUNT_CONTAMINANT")) : Call Database_LoadProperty(Rs1, rp.COUNT_CONTAMINANT)
                            Case Trim$(UCase$("RP.ROOM_VOL")) : Call Database_LoadProperty(Rs1, rp.ROOM_VOL)
                            Case Trim$(UCase$("RP.ROOM_FLOWRATE")) : Call Database_LoadProperty(Rs1, rp.ROOM_FLOWRATE)
                            Case Trim$(UCase$("RP.ROOM_C0")) : Call Database_LoadProperty(Rs1, rp.ROOM_C0(Use_FieldIndex))
                            Case Trim$(UCase$("RP.ROOM_EMIT")) : Call Database_LoadProperty(Rs1, rp.ROOM_EMIT(Use_FieldIndex))
                            ' CALCULATED ROOM PARAMETERS.
                            Case Trim$(UCase$("RP.ROOM_CHANGE_RATE")) : Call Database_LoadProperty(Rs1, rp.ROOM_CHANGE_RATE)
                            Case Trim$(UCase$("RP.ROOM_SS_VALUE")) : Call Database_LoadProperty(Rs1, rp.ROOM_SS_VALUE(Use_FieldIndex))
                            ' UNITS FOR ALL VARIABLES.
                            Case Trim$(UCase$("RP.ROOM_VOL_Units")) : Call Database_LoadProperty(Rs1, rp.ROOM_VOL_Units)
                            Case Trim$(UCase$("RP.ROOM_FLOWRATE_Units")) : Call Database_LoadProperty(Rs1, rp.ROOM_FLOWRATE_Units)
                            Case Trim$(UCase$("RP.ROOM_C0_Units")) : Call Database_LoadProperty(Rs1, rp.ROOM_C0_Units)
                            Case Trim$(UCase$("RP.ROOM_EMIT_Units")) : Call Database_LoadProperty(Rs1, rp.ROOM_EMIT_Units)
                            Case Trim$(UCase$("RP.INITIAL_ROOM_CONC_Units")) : Call Database_LoadProperty(Rs1, rp.INITIAL_ROOM_CONC_Units)
            ' NEW AS OF 9/16/98.
                            Case Trim$(UCase$("RP.INITIAL_ROOM_CONC")) : Call Database_LoadProperty(Rs1, rp.INITIAL_ROOM_CONC(Use_FieldIndex))
            ' NEW AS OF 9/16/98 ENDS.
            ' NEW AS OF 8/18/99.
                            Case Trim$(UCase$("RP.RXN_RATE_CONSTANT")) : Call Database_LoadProperty(Rs1, rp.RXN_RATE_CONSTANT(Use_FieldIndex))
                            Case Trim$(UCase$("RP.RXN_PRODUCT")) : Call Database_LoadProperty(Rs1, rp.RXN_PRODUCT(Use_FieldIndex))
                            Case Trim$(UCase$("RP.RXN_RATIO")) : Call Database_LoadProperty(Rs1, rp.RXN_RATIO(Use_FieldIndex))
            ' NEW AS OF 8/18/99 ENDS.
            '---- NEW AS OF 11/11/99 BEGINS: ---------------------------------------------------------
            '
            '/////////   TIME-VARIABLE Co   //////////////////////////////////
                            Case Trim$(UCase$("RP.bool_ROOM_COINI_ISTIMEVAR")) : Call Database_LoadProperty(Rs1, rp.bool_ROOM_COINI_ISTIMEVAR(Use_FieldIndex))
                            Case Trim$(UCase$("RP.int_ROOM_NCOINI")) : Call Database_LoadProperty(Rs1, rp.int_ROOM_NCOINI(Use_FieldIndex))
                            Case Trim$(UCase$("RP.u_ROOM_TCOINI")) : Call Database_LoadProperty(Rs1, rp.u_ROOM_TCOINI)
                            Case Trim$(UCase$("RP.u_ROOM_COINI")) : Call Database_LoadProperty(Rs1, rp.u_ROOM_COINI)
            ''''dbl_ROOM_TCOINI() As Double   '(x,y): x=chemical, y=row
            ''''dbl_ROOM_COINI() As Double    '(x,y): x=chemical, y=row
            '
            '/////////   TIME-VARIABLE w*A   /////////////////////////////////
                            Case Trim$(UCase$("RP.bool_ROOM_EMITINI_ISTIMEVAR")) : Call Database_LoadProperty(Rs1, rp.bool_ROOM_EMITINI_ISTIMEVAR(Use_FieldIndex))
                            Case Trim$(UCase$("RP.int_ROOM_NEMITINI")) : Call Database_LoadProperty(Rs1, rp.int_ROOM_NEMITINI(Use_FieldIndex))
                            Case Trim$(UCase$("RP.u_ROOM_TEMITINI")) : Call Database_LoadProperty(Rs1, rp.u_ROOM_TEMITINI)
                            Case Trim$(UCase$("RP.u_ROOM_EMITINI")) : Call Database_LoadProperty(Rs1, rp.u_ROOM_EMITINI)
            ''''dbl_ROOM_TEMITINI() As Double   '(x,y): x=chemical, y=row
            ''''dbl_ROOM_EMITINI() As Double    '(x,y): x=chemical, y=row
            '---- NEW AS OF 11/11/99 ENDS. ---------------------------------------------------------
            '---- NEW AS OF 1/17/00 BEGINS: ---------------------------------------------------------
            '
            '/////////   TIME-VARIABLE K   /////////////////////////////////
                            Case Trim$(UCase$("RP.bool_ROOM_KINI_ISTIMEVAR")) : Call Database_LoadProperty(Rs1, rp.bool_ROOM_KINI_ISTIMEVAR(Use_FieldIndex))
                            Case Trim$(UCase$("RP.int_ROOM_NKINI")) : Call Database_LoadProperty(Rs1, rp.int_ROOM_NKINI(Use_FieldIndex))
                            Case Trim$(UCase$("RP.u_ROOM_TKINI")) : Call Database_LoadProperty(Rs1, rp.u_ROOM_TKINI)
                            Case Trim$(UCase$("RP.u_ROOM_KINI")) : Call Database_LoadProperty(Rs1, rp.u_ROOM_KINI)
                                ''''dbl_ROOM_TKINI() As Double   '(x,y): x=chemical, y=row
                                ''''dbl_ROOM_KINI() As Double    '(x,y): x=chemical, y=row
                                '---- NEW AS OF 1/17/00 ENDS. ---------------------------------------------------------
                        End Select
                        Rs1.MoveNext()
                    Loop
                End If
                Rs1.Close()
            End If
            RoomParams = rp
            Call RoomParam_Recalculate(RoomParams)
        End If

        '---- NEW AS OF 11/11/99 BEGINS: ---------------------------------------------------------
        '------ INPUT DATA FROM TABLE "PSDMInRoomData_CO_Data". ------------------------------------------------------------------------------------------------------
        If (Database_IsTableExist(Db1, "PSDMInRoomData_CO_Data") = False) Then
            'DO NOTHING: USE DEFAULT VALUES.
        Else
            Rs1 = Db1.OpenRecordset("PSDMInRoomData_CO_Data")
			If (Rs1.RecordCount = 0) Then
				'DO NOTHING: USE DEFAULT VALUES.
			Else
				Rs1.MoveFirst()

                Do Until Rs1.EOF
                    Use_FieldIndex = CInt(Database_Get_Long(Rs1, "FieldIndex"))
                    Use_FieldIndex2 = CInt(Database_Get_Long(Rs1, "FieldIndex2"))
                    Select Case Trim$(UCase$(Database_Get_String(Rs1, "FieldName")))
                        Case Trim$(UCase$("dbl_ROOM_TCOINI")) : Call Database_LoadProperty(Rs1, RoomParams.dbl_ROOM_TCOINI(Use_FieldIndex, Use_FieldIndex2))
                        Case Trim$(UCase$("dbl_ROOM_COINI")) : Call Database_LoadProperty(Rs1, RoomParams.dbl_ROOM_COINI(Use_FieldIndex, Use_FieldIndex2))
                    End Select
                    Rs1.MoveNext()
                Loop
            End If
            Rs1.Close()
        End If
        '---- NEW AS OF 11/11/99 ENDS. ---------------------------------------------------------

        '---- NEW AS OF 11/11/99 BEGINS: ---------------------------------------------------------
        '------ INPUT DATA FROM TABLE "PSDMInRoomData_WA_Data". ------------------------------------------------------------------------------------------------------
        If (Database_IsTableExist(Db1, "PSDMInRoomData_WA_Data") = False) Then
            'DO NOTHING: USE DEFAULT VALUES.
        Else
            Rs1 = Db1.OpenRecordset("PSDMInRoomData_WA_Data")
			If (Rs1.RecordCount = 0) Then
				'DO NOTHING: USE DEFAULT VALUES.
			Else
				Rs1.MoveFirst()

                Do Until Rs1.EOF
                    Use_FieldIndex = CInt(Database_Get_Long(Rs1, "FieldIndex"))
                    Use_FieldIndex2 = CInt(Database_Get_Long(Rs1, "FieldIndex2"))
                    Select Case Trim$(UCase$(Database_Get_String(Rs1, "FieldName")))
                        Case Trim$(UCase$("dbl_ROOM_TEMITINI")) : Call Database_LoadProperty(Rs1, RoomParams.dbl_ROOM_TEMITINI(Use_FieldIndex, Use_FieldIndex2))
                        Case Trim$(UCase$("dbl_ROOM_EMITINI")) : Call Database_LoadProperty(Rs1, RoomParams.dbl_ROOM_EMITINI(Use_FieldIndex, Use_FieldIndex2))
                    End Select
                    Rs1.MoveNext()
                Loop
            End If
            Rs1.Close()
        End If
        '---- NEW AS OF 11/11/99 ENDS. ---------------------------------------------------------

        '---- NEW AS OF 1/17/00 BEGINS: ---------------------------------------------------------
        '------ INPUT DATA FROM TABLE "PSDMInRoomData_K_Data". ------------------------------------------------------------------------------------------------------
        If (Database_IsTableExist(Db1, "PSDMInRoomData_K_Data") = False) Then
            'DO NOTHING: USE DEFAULT VALUES.
        Else
            Rs1 = Db1.OpenRecordset("PSDMInRoomData_K_Data")
			'			If (Database_NoRecordsInRecordset(Rs1)) Then
			If (Rs1.RecordCount = 0) Then

				'DO NOTHING: USE DEFAULT VALUES.
			Else
				Rs1.MoveFirst()

                Do Until Rs1.EOF
                    Use_FieldIndex = CInt(Database_Get_Long(Rs1, "FieldIndex"))
                    Use_FieldIndex2 = CInt(Database_Get_Long(Rs1, "FieldIndex2"))
                    Select Case Trim$(UCase$(Database_Get_String(Rs1, "FieldName")))
                        Case Trim$(UCase$("dbl_ROOM_TKINI")) : Call Database_LoadProperty(Rs1, RoomParams.dbl_ROOM_TKINI(Use_FieldIndex, Use_FieldIndex2))
                        Case Trim$(UCase$("dbl_ROOM_KINI")) : Call Database_LoadProperty(Rs1, RoomParams.dbl_ROOM_KINI(Use_FieldIndex, Use_FieldIndex2))
                    End Select
                    Rs1.MoveNext()
                Loop
            End If
            Rs1.Close()
        End If
        '---- NEW AS OF 1/17/00 ENDS. ---------------------------------------------------------

        'CLOSE THE DATABASE FILE.
        Db1.Close()

        'RETURN A "SUCCESS" MESSAGE TO CALLER.
        File_Open_Latest_v1_60 = True

    End Function












	Function File_Save_Latest_v1_60(ByRef fn_This As String) As Boolean
		'	Dim OpenDatabase As Object    'Out shang
		'UPGRADE_ISSUE: Workspace object was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim Ws1 As dao.Workspace
		'UPGRADE_ISSUE: Database object was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim Db1 As dao.Database
		'UPGRADE_ISSUE: Recordset object was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim Rs1 As dao.Recordset
		Dim i As Short
		Dim J As Short
		'UPGRADE_WARNING: Arrays in structure Co may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim Co As ComponentPropertyType
		Dim Be As BedPropertyType
		Dim IsLegacyVersion As Boolean
		Dim NeedToCreateNewDatabase As Boolean
		'UPGRADE_WARNING: Arrays in structure rp may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim rp As RoomParam_Type
		
		'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
		'>>>>>>>>>>>>>>>>>>>>>>>>>>>  SAVE TO MAIN DATABASE  <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
		'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
		
		'IF FILE DOES NOT EXIST, CREATE IT.
		'FOR EACH TABLE, IF IT EXISTS, DELETE IT.
		If (Not FileExists(fn_This)) Then
			'CREATE NEW DATABASE.
			NeedToCreateNewDatabase = True
		Else
			'DETERMINE WHETHER OLD FILE IS A LEGACY VERSION (i.e. A NON-MDB FILE).
			IsLegacyVersion = True
			On Error Resume Next
			'			Db1 = OpenDatabase(fn_This)
			Db1 = DAOEngine.OpenDatabase(fn_This)
			If (Err.Number = 0) Then
				IsLegacyVersion = False
				'UPGRADE_WARNING: Couldn't resolve default property of object Db1.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Db1.Close()
			End If
			On Error GoTo 0
			If (IsLegacyVersion) Then
				'DELETE OLD FILE, CREATE NEW DATABASE (SEE BELOW).
				Kill(fn_This)
				NeedToCreateNewDatabase = True
			Else
				'OPEN DATABASE NORMALLY.
				'Db1 = OpenDatabase(fn_This)
				Db1 = DAOEngine.OpenDatabase(fn_This)
			End If
		End If
		If (NeedToCreateNewDatabase) Then
			'	FileCopy(MAIN_APP_PATH & "\dbase\template.dat", fn_This)     'out shang
			Db1 = DAOEngine.CreateDatabase(fn_This, dao.LanguageConstants.dbLangGeneral)  'In shang
			'Db1.Close()

			Db1 = DAOEngine.OpenDatabase(fn_This)

		End If
		'CREATE NEW TABLES WITHIN DATABASE, IF NECESSARY.
		Call Database_CreateMFBTable_IfNoExist(Db1, "Version", True)
		Call Database_CreateMFBTable_IfNoExist(Db1, "Main", True)
		Call Database_CreateMFBTable_IfNoExist_TwoIndices(Db1, "InfluentPoints")
		Call Database_CreateMFBTable_IfNoExist_TwoIndices(Db1, "EffluentPoints")
		Call Database_CreateMFBTable_IfNoExist(Db1, "PSDMInRoomData", True)
		Call Database_CreateMFBTable_IfNoExist_TwoIndices(Db1, "PSDMInRoomData_CO_Data")
		Call Database_CreateMFBTable_IfNoExist_TwoIndices(Db1, "PSDMInRoomData_WA_Data")
		Call Database_CreateMFBTable_IfNoExist_TwoIndices(Db1, "PSDMInRoomData_K_Data")
		
		'=========== OUTPUT DATA TO DATABASE TABLES. =================
		
		'------ OUTPUT DATA TO TABLE "Version". ------------------------------------------------------------------------------------------------------
		'START SAVE TO THIS TABLE.
		Call Database_DeleteTableContents(Db1, "Version")
		'UPGRADE_WARNING: Couldn't resolve default property of object Db1.OpenRecordset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1 = Db1.OpenRecordset("Version")
		Call Database_SaveProperty(Rs1, "DataVersion_Major", CShort(Latest_DataVersion_Major))
		Call Database_SaveProperty(Rs1, "DataVersion_Minor", CShort(Latest_DataVersion_Minor))
		Call Database_SaveProperty(Rs1, "ContainsTable_PSDMInRoomData", True)
		'END SAVE TO THIS TABLE.
		'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1.Close()
		
		'------ OUTPUT DATA TO TABLE "Main". ------------------------------------------------------------------------------------------------------
		'START SAVE TO THIS TABLE.
		Call Database_DeleteTableContents(Db1, "Main")
		'UPGRADE_WARNING: Couldn't resolve default property of object Db1.OpenRecordset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1 = Db1.OpenRecordset("Main")
		'HEADER BLOCK.
		Call Database_SaveProperty(Rs1, "*FileNote", FileNote)
		Call Database_SaveProperty(Rs1, "Number_Component", Number_Component)
		'COMPONENT PROPERTIES BLOCK.
		For i = 1 To Number_Component
			'UPGRADE_WARNING: Couldn't resolve default property of object Co. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Co = Component(i)
			Call Database_SavePropertyWithIndex(Rs1, "Co.Name", i, Co.Name)
			Call Database_SavePropertyWithIndex(Rs1, "Co.CAS", i, Co.CAS)
			Call Database_SavePropertyWithIndex(Rs1, "Co.MW", i, Co.MW)
			Call Database_SavePropertyWithIndex(Rs1, "Co.InitialConcentration", i, Co.InitialConcentration)
			Call Database_SavePropertyWithIndex(Rs1, "Co.MolarVolume", i, Co.MolarVolume)
			Call Database_SavePropertyWithIndex(Rs1, "Co.BP", i, Co.BP)
			Call Database_SavePropertyWithIndex(Rs1, "Co.Use_K", i, Co.Use_K)
			Call Database_SavePropertyWithIndex(Rs1, "Co.Use_OneOverN", i, Co.Use_OneOverN)
			Call Database_SavePropertyWithIndex(Rs1, "Co.Liquid_Density", i, Co.Liquid_Density)
			Call Database_SavePropertyWithIndex(Rs1, "Co.Aqueous_Solubility", i, Co.Aqueous_Solubility)
			Call Database_SavePropertyWithIndex(Rs1, "Co.Vapor_Pressure", i, Co.Vapor_Pressure)
			Call Database_SavePropertyWithIndex(Rs1, "Co.Refractive_Index", i, Co.Refractive_Index)
			Call Database_SavePropertyWithIndex(Rs1, "Co.SPDFR", i, Co.SPDFR)
			Call Database_SavePropertyWithIndex(Rs1, "Co.SPDFR_Low_Concentration", i, Co.SPDFR_Low_Concentration)
			Call Database_SavePropertyWithIndex(Rs1, "Co.Use_SPDFR_Correlation", i, Co.Use_SPDFR_Correlation)
			Call Database_SavePropertyWithIndex(Rs1, "Co.kf", i, Co.kf)
			Call Database_SavePropertyWithIndex(Rs1, "Co.Ds", i, Co.Ds)
			Call Database_SavePropertyWithIndex(Rs1, "Co.Dp", i, Co.Dp)
			Call Database_SavePropertyWithIndex(Rs1, "Co.Corr(1)", i, Co.Corr(1))
			Call Database_SavePropertyWithIndex(Rs1, "Co.Corr(2)", i, Co.Corr(2))
			Call Database_SavePropertyWithIndex(Rs1, "Co.Corr(3)", i, Co.Corr(3))
			Call Database_SavePropertyWithIndex(Rs1, "Co.KP_User_Input(1)", i, Co.KP_User_Input(1))
			Call Database_SavePropertyWithIndex(Rs1, "Co.KP_User_Input(2)", i, Co.KP_User_Input(2))
			Call Database_SavePropertyWithIndex(Rs1, "Co.KP_User_Input(3)", i, Co.KP_User_Input(3))
			Call Database_SavePropertyWithIndex(Rs1, "Co.K_Reduction", i, Co.K_Reduction)
			Call Database_SavePropertyWithIndex(Rs1, "Co.Correlation.Name", i, Co.Correlation.Name)
			Call Database_SavePropertyWithIndex(Rs1, "Co.Correlation.Coeff(1)", i, Co.Correlation.Coeff(1))
			Call Database_SavePropertyWithIndex(Rs1, "Co.Correlation.Coeff(2)", i, Co.Correlation.Coeff(2))
			Call Database_SavePropertyWithIndex(Rs1, "Co.IsothermDB_Component_Name", i, Co.IsothermDB_Component_Name)
			Call Database_SavePropertyWithIndex(Rs1, "Co.IsothermDB_Range_Num", i, Co.IsothermDB_Range_Num)
			Call Database_SavePropertyWithIndex(Rs1, "Co.IPES_OrderOfMagnitude", i, Co.IPES_OrderOfMagnitude)
			Call Database_SavePropertyWithIndex(Rs1, "Co.IPES_NumRegressionPts", i, Co.IPES_NumRegressionPts)
			Call Database_SavePropertyWithIndex(Rs1, "Co.IPES_RelativeHumidity", i, Co.IPES_RelativeHumidity)
			Call Database_SavePropertyWithIndex(Rs1, "Co.IPES_EstimationMethod", i, Co.IPES_EstimationMethod)
			Call Database_SavePropertyWithIndex(Rs1, "Co.Source_KandOneOverN", i, Co.Source_KandOneOverN)
			Call Database_SavePropertyWithIndex(Rs1, "Co.IsothermDB_K", i, Co.IsothermDB_K)
			Call Database_SavePropertyWithIndex(Rs1, "Co.IsothermDB_OneOverN", i, Co.IsothermDB_OneOverN)
			Call Database_SavePropertyWithIndex(Rs1, "Co.IPESResult_K", i, Co.IPESResult_K)
			Call Database_SavePropertyWithIndex(Rs1, "Co.IPESResult_OneOverN", i, Co.IPESResult_OneOverN)
			Call Database_SavePropertyWithIndex(Rs1, "Co.UserEntered_K", i, Co.UserEntered_K)
			Call Database_SavePropertyWithIndex(Rs1, "Co.UserEntered_OneOverN", i, Co.UserEntered_OneOverN)
			Call Database_SavePropertyWithIndex(Rs1, "Co.Tortuosity", i, Co.Tortuosity)
			Call Database_SavePropertyWithIndex(Rs1, "Co.Use_Tortuosity_Correlation", i, Co.Use_Tortuosity_Correlation)
			Call Database_SavePropertyWithIndex(Rs1, "Co.Constant_Tortuosity", i, Co.Constant_Tortuosity)
		Next i
		'BED PROPERTIES BLOCK.
		'UPGRADE_WARNING: Couldn't resolve default property of object Be. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Be = Bed
		Call Database_SaveProperty(Rs1, "Be.length", Be.length)
		Call Database_SaveProperty(Rs1, "Be.Diameter", Be.Diameter)
		Call Database_SaveProperty(Rs1, "Be.Weight", Be.Weight)
		Call Database_SaveProperty(Rs1, "Be.Flowrate", Be.Flowrate)
		Call Database_SaveProperty(Rs1, "Be.WaterDensity", Be.WaterDensity)
		Call Database_SaveProperty(Rs1, "Be.WaterViscosity", Be.WaterViscosity)
		Call Database_SaveProperty(Rs1, "Be.Temperature", Be.Temperature)
		Call Database_SaveProperty(Rs1, "Be.Pressure", Be.Pressure)
		Call Database_SaveProperty(Rs1, "Be.Phase", Be.Phase)
		Call Database_SaveProperty(Rs1, "Be.NumberOfBeds", Be.NumberOfBeds)
		Call Database_SaveProperty(Rs1, "Be.Water_Correlation.Name", Be.Water_Correlation.Name)
		Call Database_SaveProperty(Rs1, "Be.Water_Correlation.Coeff(1)", Be.Water_Correlation.Coeff(1))
		Call Database_SaveProperty(Rs1, "Be.Water_Correlation.Coeff(2)", Be.Water_Correlation.Coeff(2))
		Call Database_SaveProperty(Rs1, "Be.Water_Correlation.Coeff(3)", Be.Water_Correlation.Coeff(3))
		Call Database_SaveProperty(Rs1, "Be.Water_Correlation.Coeff(4)", Be.Water_Correlation.Coeff(4))
		'UNITS BLOCK.
		Call Units1_Database_SaveProperty(Rs1, frmMain.txtBedUnits(0), "frmMain.txtBedUnits(0)")
		Call Units1_Database_SaveProperty(Rs1, frmMain.txtBedUnits(1), "frmMain.txtBedUnits(1)")
		Call Units1_Database_SaveProperty(Rs1, frmMain.txtBedUnits(2), "frmMain.txtBedUnits(2)")
		Call Units1_Database_SaveProperty(Rs1, frmMain.txtBedUnits(3), "frmMain.txtBedUnits(3)")
		Call Units1_Database_SaveProperty(Rs1, frmMain.txtBedUnits(4), "frmMain.txtBedUnits(4)")
		Call Units1_Database_SaveProperty(Rs1, frmMain.txtCarbonUnits(1), "frmMain.txtCarbonUnits(1)")
		Call Units1_Database_SaveProperty(Rs1, frmMain.txtCarbonUnits(2), "frmMain.txtCarbonUnits(2)")
		Call Units1_Database_SaveProperty(Rs1, frmMain.txtTimeUnits(0), "frmMain.txtTimeUnits(0)")
		Call Units1_Database_SaveProperty(Rs1, frmMain.txtTimeUnits(1), "frmMain.txtTimeUnits(1)")
		Call Units1_Database_SaveProperty(Rs1, frmMain.txtTimeUnits(2), "frmMain.txtTimeUnits(2)")
		Call Database_SaveProperty(Rs1, "PropertyUnits.MW", PropertyUnits.MW)
		Call Database_SaveProperty(Rs1, "PropertyUnits.MolarVolume", PropertyUnits.MolarVolume)
		Call Database_SaveProperty(Rs1, "PropertyUnits.BP", PropertyUnits.BP)
		Call Database_SaveProperty(Rs1, "PropertyUnits.InitialConcentration", PropertyUnits.InitialConcentration)
		Call Database_SaveProperty(Rs1, "PropertyUnits.Liquid_Density", PropertyUnits.Liquid_Density)
		Call Database_SaveProperty(Rs1, "PropertyUnits.Aqueous_Solubility", PropertyUnits.Aqueous_Solubility)
		Call Database_SaveProperty(Rs1, "PropertyUnits.Vapor_Pressure", PropertyUnits.Vapor_Pressure)
		Call Database_SaveProperty(Rs1, "PropertyUnits.k", PropertyUnits.k)
		'MISCELLANEOUS BLOCK.
		Call Database_SaveProperty(Rs1, "Carbon.Name", Carbon.Name)
		Call Database_SaveProperty(Rs1, "Carbon.Porosity", Carbon.Porosity)
		Call Database_SaveProperty(Rs1, "Carbon.Density", Carbon.Density)
		Call Database_SaveProperty(Rs1, "Carbon.ParticleRadius", Carbon.ParticleRadius)
		Call Database_SaveProperty(Rs1, "Carbon.Tortuosity", Carbon.Tortuosity)
		Call Database_SaveProperty(Rs1, "Carbon.W0", Carbon.W0)
		Call Database_SaveProperty(Rs1, "Carbon.BB", Carbon.BB)
		Call Database_SaveProperty(Rs1, "Carbon.PolanyiExponent", Carbon.PolanyiExponent)
		Call Database_SaveProperty(Rs1, "State_Check_Water(1)", State_Check_Water(1))
		Call Database_SaveProperty(Rs1, "State_Check_Water(2)", State_Check_Water(2))
		Call Database_SaveProperty(Rs1, "Carbon.ShapeFactor", Carbon.ShapeFactor)
		Call Database_SaveProperty(Rs1, "Constant_Tortuosity", Constant_Tortuosity)
		Call Database_SaveProperty(Rs1, "Carbon.ShapeFactor", Carbon.ShapeFactor)
		Call Database_SaveProperty(Rs1, "NC", NC)
		Call Database_SaveProperty(Rs1, "MC", MC)
		Call Database_SaveProperty(Rs1, "TimeP.Init", TimeP.Init)
		Call Database_SaveProperty(Rs1, "TimeP.End", TimeP.End_Renamed)
		Call Database_SaveProperty(Rs1, "TimeP.np", TimeP.np)
		Call Database_SaveProperty(Rs1, "TimeP.Step", TimeP.Step_Renamed)
		'INFLUENT/EFFLUENT POINT COUNTS.
		Call Database_SaveProperty(Rs1, "Number_Influent_Points", Number_Influent_Points)
		Call Database_SaveProperty(Rs1, "NData_Points", NData_Points)
		'END SAVE TO THIS TABLE.
		'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1.Close()
		
		'------ OUTPUT DATA TO TABLE "InfluentPoints". ------------------------------------------------------------------------------------------------------
		'START SAVE TO THIS TABLE.
		Call Database_DeleteTableContents(Db1, "InfluentPoints")
		'UPGRADE_WARNING: Couldn't resolve default property of object Db1.OpenRecordset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1 = Db1.OpenRecordset("InfluentPoints")
		'MAIN DATA SET.
		If (Number_Influent_Points > 0) Then
			For i = 1 To Number_Influent_Points
				''''Write #f, T_Influent(i)
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.AddNew. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.AddNew()
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("FieldName").Value = "T_Influent"
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("FieldIndex").Value = i
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("dblValue").Value = T_Influent(i)
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Update. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.Update()
				For J = 1 To Number_Component
					''''Write #f, C_Influent(j, i)
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.AddNew. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1.AddNew()
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1("FieldName").Value = "C_Influent"
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1("FieldIndex").Value = J
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1("FieldIndex2").Value = i
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1("dblValue").Value = C_Influent(J, i)
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Update. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1.Update()
				Next J
			Next i
		End If
		'END SAVE TO THIS TABLE.
		'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1.Close()
		
		'------ OUTPUT DATA TO TABLE "EffluentPoints". ------------------------------------------------------------------------------------------------------
		'START SAVE TO THIS TABLE.
		Call Database_DeleteTableContents(Db1, "EffluentPoints")
		'UPGRADE_WARNING: Couldn't resolve default property of object Db1.OpenRecordset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1 = Db1.OpenRecordset("EffluentPoints")
		'MAIN DATA SET.
		If (NData_Points > 0) Then
			For i = 1 To NData_Points
				''''Write #f, T_Data_Points(i)
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.AddNew. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.AddNew()
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("FieldName").Value = "T_Data_Points"
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("FieldIndex").Value = i
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("dblValue").Value = T_Data_Points(i)
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Update. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.Update()
				For J = 1 To Number_Component
					''''Write #f, C_Data_Points(j, i)
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.AddNew. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1.AddNew()
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1("FieldName").Value = "C_Data_Points"
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1("FieldIndex").Value = J
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1("FieldIndex2").Value = i
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1("dblValue").Value = C_Data_Points(J, i)
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Update. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1.Update()
				Next J
			Next i
		End If
		'END SAVE TO THIS TABLE.
		'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1.Close()
		
		'------ OUTPUT DATA TO TABLE "PSDMInRoomData". ------------------------------------------------------------------------------------------------------
		'START SAVE TO THIS TABLE.
		Call Database_DeleteTableContents(Db1, "PSDMInRoomData")
		'UPGRADE_WARNING: Couldn't resolve default property of object Db1.OpenRecordset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1 = Db1.OpenRecordset("PSDMInRoomData")
		' INPUT ROOM PARAMETERS.
		'UPGRADE_WARNING: Couldn't resolve default property of object rp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rp = RoomParams
		Call Database_SaveProperty(Rs1, "RP.COUNT_CONTAMINANT", rp.COUNT_CONTAMINANT)
		Call Database_SaveProperty(Rs1, "RP.ROOM_VOL", rp.ROOM_VOL)
		Call Database_SaveProperty(Rs1, "RP.ROOM_FLOWRATE", rp.ROOM_FLOWRATE)
		For i = 1 To rp.COUNT_CONTAMINANT
			Call Database_SavePropertyWithIndex(Rs1, "RP.ROOM_C0", i, rp.ROOM_C0(i))
			Call Database_SavePropertyWithIndex(Rs1, "RP.ROOM_EMIT", i, rp.ROOM_EMIT(i))
		Next i
		' CALCULATED ROOM PARAMETERS.
		Call Database_SaveProperty(Rs1, "RP.ROOM_CHANGE_RATE", rp.ROOM_CHANGE_RATE)
		For i = 1 To rp.COUNT_CONTAMINANT
			Call Database_SavePropertyWithIndex(Rs1, "RP.ROOM_SS_VALUE", i, rp.ROOM_SS_VALUE(i))
		Next i
		' UNITS FOR ALL VARIABLES.
		Call Database_SaveProperty(Rs1, "RP.ROOM_VOL_Units", rp.ROOM_VOL_Units)
		Call Database_SaveProperty(Rs1, "RP.ROOM_FLOWRATE_Units", rp.ROOM_FLOWRATE_Units)
		Call Database_SaveProperty(Rs1, "RP.ROOM_C0_Units", rp.ROOM_C0_Units)
		Call Database_SaveProperty(Rs1, "RP.ROOM_EMIT_Units", rp.ROOM_EMIT_Units)
		Call Database_SaveProperty(Rs1, "RP.INITIAL_ROOM_CONC_Units", rp.INITIAL_ROOM_CONC_Units)
		' NEW AS OF 9/16/98.
		For i = 1 To rp.COUNT_CONTAMINANT
			Call Database_SavePropertyWithIndex(Rs1, "RP.INITIAL_ROOM_CONC", i, rp.INITIAL_ROOM_CONC(i))
		Next i
		' NEW AS OF 9/16/98 ENDS.
		' NEW AS OF 8/18/99.
		For i = 1 To rp.COUNT_CONTAMINANT
			Call Database_SavePropertyWithIndex(Rs1, "RP.RXN_RATE_CONSTANT", i, rp.RXN_RATE_CONSTANT(i))
			Call Database_SavePropertyWithIndex(Rs1, "RP.RXN_PRODUCT", i, rp.RXN_PRODUCT(i))
			Call Database_SavePropertyWithIndex(Rs1, "RP.RXN_RATIO", i, rp.RXN_RATIO(i))
		Next i
		' NEW AS OF 8/18/99 ENDS.
		'---- NEW AS OF 11/11/99 BEGINS: ---------------------------------------------------------
		'
		'/////////   TIME-VARIABLE Co   //////////////////////////////////
		For i = 1 To rp.COUNT_CONTAMINANT
			Call Database_SavePropertyWithIndex(Rs1, "RP.bool_ROOM_COINI_ISTIMEVAR", i, rp.bool_ROOM_COINI_ISTIMEVAR(i))
		Next i
		For i = 1 To rp.COUNT_CONTAMINANT
			Call Database_SavePropertyWithIndex(Rs1, "RP.int_ROOM_NCOINI", i, rp.int_ROOM_NCOINI(i))
		Next i
		Call Database_SaveProperty(Rs1, "RP.u_ROOM_TCOINI", rp.u_ROOM_TCOINI)
		Call Database_SaveProperty(Rs1, "RP.u_ROOM_COINI", rp.u_ROOM_COINI)
		''''dbl_ROOM_TCOINI() As Double   '(x,y): x=chemical, y=row
		''''dbl_ROOM_COINI() As Double    '(x,y): x=chemical, y=row
		'
		'/////////   TIME-VARIABLE w*A   /////////////////////////////////
		For i = 1 To rp.COUNT_CONTAMINANT
			Call Database_SavePropertyWithIndex(Rs1, "RP.bool_ROOM_EMITINI_ISTIMEVAR", i, rp.bool_ROOM_EMITINI_ISTIMEVAR(i))
		Next i
		For i = 1 To rp.COUNT_CONTAMINANT
			Call Database_SavePropertyWithIndex(Rs1, "RP.int_ROOM_NEMITINI", i, rp.int_ROOM_NEMITINI(i))
		Next i
		Call Database_SaveProperty(Rs1, "RP.u_ROOM_TEMITINI", rp.u_ROOM_TEMITINI)
		Call Database_SaveProperty(Rs1, "RP.u_ROOM_EMITINI", rp.u_ROOM_EMITINI)
		''''dbl_ROOM_TEMITINI() As Double   '(x,y): x=chemical, y=row
		''''dbl_ROOM_EMITINI() As Double    '(x,y): x=chemical, y=row
		'---- NEW AS OF 11/11/99 ENDS. ---------------------------------------------------------
		
		'---- NEW AS OF 1/17/00 BEGINS: ---------------------------------------------------------
		'
		'/////////   TIME-VARIABLE K   /////////////////////////////////
		For i = 1 To rp.COUNT_CONTAMINANT
			Call Database_SavePropertyWithIndex(Rs1, "RP.bool_ROOM_KINI_ISTIMEVAR", i, rp.bool_ROOM_KINI_ISTIMEVAR(i))
		Next i
		For i = 1 To rp.COUNT_CONTAMINANT
			Call Database_SavePropertyWithIndex(Rs1, "RP.int_ROOM_NKINI", i, rp.int_ROOM_NKINI(i))
		Next i
		Call Database_SaveProperty(Rs1, "RP.u_ROOM_TKINI", rp.u_ROOM_TKINI)
		Call Database_SaveProperty(Rs1, "RP.u_ROOM_KINI", rp.u_ROOM_KINI)
		''''dbl_ROOM_TKINI() As Double   '(x,y): x=chemical, y=row
		''''dbl_ROOM_KINI() As Double    '(x,y): x=chemical, y=row
		'---- NEW AS OF 1/17/00 ENDS. ---------------------------------------------------------
		
		'END SAVE TO THIS TABLE.
		'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1.Close()
		
		'---- NEW AS OF 11/11/99 BEGINS: ---------------------------------------------------------
		'START SAVE TO THIS TABLE.
		Call Database_DeleteTableContents(Db1, "PSDMInRoomData_CO_Data")
		'UPGRADE_WARNING: Couldn't resolve default property of object Db1.OpenRecordset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1 = Db1.OpenRecordset("PSDMInRoomData_CO_Data")
		'MAIN DATA SET.
		For J = 1 To Number_Component
			For i = 1 To rp.int_ROOM_NCOINI(J)
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.AddNew. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.AddNew()
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("FieldName").Value = "dbl_ROOM_TCOINI"
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("FieldIndex").Value = J
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("FieldIndex2").Value = i
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("dblValue").Value = rp.dbl_ROOM_TCOINI(J, i)
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Update. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.Update()
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.AddNew. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.AddNew()
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("FieldName").Value = "dbl_ROOM_COINI"
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("FieldIndex").Value = J
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("FieldIndex2").Value = i
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("dblValue").Value = rp.dbl_ROOM_COINI(J, i)
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Update. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.Update()
			Next i
		Next J
		'END SAVE TO THIS TABLE.
		'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1.Close()
		'---- NEW AS OF 11/11/99 ENDS. ---------------------------------------------------------
		
		'---- NEW AS OF 11/11/99 BEGINS: ---------------------------------------------------------
		'START SAVE TO THIS TABLE.
		Call Database_DeleteTableContents(Db1, "PSDMInRoomData_WA_Data")
		'UPGRADE_WARNING: Couldn't resolve default property of object Db1.OpenRecordset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1 = Db1.OpenRecordset("PSDMInRoomData_WA_Data")
		'MAIN DATA SET.
		For J = 1 To Number_Component
			For i = 1 To rp.int_ROOM_NEMITINI(J)
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.AddNew. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.AddNew()
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("FieldName").Value = "dbl_ROOM_TEMITINI"
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("FieldIndex").Value = J
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("FieldIndex2").Value = i
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("dblValue").Value = rp.dbl_ROOM_TEMITINI(J, i)
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Update. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.Update()
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.AddNew. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.AddNew()
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("FieldName").Value = "dbl_ROOM_EMITINI"
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("FieldIndex").Value = J
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("FieldIndex2").Value = i
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("dblValue").Value = rp.dbl_ROOM_EMITINI(J, i)
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Update. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.Update()
			Next i
		Next J
		'END SAVE TO THIS TABLE.
		'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1.Close()
		'---- NEW AS OF 11/11/99 ENDS. ---------------------------------------------------------
		
		'---- NEW AS OF 1/17/00 BEGINS: ---------------------------------------------------------
		'START SAVE TO THIS TABLE.
		Call Database_DeleteTableContents(Db1, "PSDMInRoomData_K_Data")
		'UPGRADE_WARNING: Couldn't resolve default property of object Db1.OpenRecordset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1 = Db1.OpenRecordset("PSDMInRoomData_K_Data")
		'MAIN DATA SET.
		For J = 1 To Number_Component
			For i = 1 To rp.int_ROOM_NKINI(J)
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.AddNew. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.AddNew()
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("FieldName").Value = "dbl_ROOM_TKINI"
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("FieldIndex").Value = J
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("FieldIndex2").Value = i
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("dblValue").Value = rp.dbl_ROOM_TKINI(J, i)
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Update. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.Update()
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.AddNew. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.AddNew()
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("FieldName").Value = "dbl_ROOM_KINI"
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("FieldIndex").Value = J
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("FieldIndex2").Value = i
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("dblValue").Value = rp.dbl_ROOM_KINI(J, i)
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Update. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1.Update()
			Next i
		Next J
		'END SAVE TO THIS TABLE.
		'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1.Close()
		'---- NEW AS OF 1/17/00 ENDS. ---------------------------------------------------------
		
		'CLOSE THE DATABASE FILE.
		'UPGRADE_WARNING: Couldn't resolve default property of object Db1.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Db1.Close()
		
		'COMPACT THE DATABASE FILE.
		'TO DO: USE THE DbEngine.CompactDatabase METHOD
		'TO COMPACT THE DATABASE.  PROBLEM TO CONSIDER:
		'THE DB MUST BE COMPACTED TO A TEMPORARY FILE,
		'WHICH THEN SHOULD OVERWRITE THE ORIGINAL FILE.
		
		'RETURN A "SUCCESS" MESSAGE TO CALLER.
		File_Save_Latest_v1_60 = True
		
	End Function


	Sub Units1_Database_SaveProperty(ByRef Rs1 As dao.Recordset, ByRef CboX As ComboBox, ByRef Desc As String)
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
		''''Call ProjectFile_Write(f, OutStr, Desc)
		Call Database_SaveProperty(Rs1, Desc, OutStr)
	End Sub
	Sub Units1_Database_LoadProperty(ByRef Rs1 As dao.Recordset, ByRef CboX As System.Windows.Forms.Control)
		Dim TxtX As System.Windows.Forms.Control
		Dim inline As String
		Dim Dummy1 As String
		Dim NewUnits As String
		Dim H As Short
		Call Database_LoadProperty(Rs1, inline)
		''''Call ProjectFile_Read(f, InLine, Dummy1)
		NewUnits = inline
		'UPGRADE_WARNING: Couldn't resolve default property of object unitsys_lookup_cbox(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		H = unitsys_lookup_cbox(CboX)
		TxtX = unitsys(H).TxtX
		Call unitsys_set_units(TxtX, NewUnits)
	End Sub


	Sub Database_LoadProperty(ByRef Rs1 As dao.Recordset, ByRef LoadedData As Object, Optional ByRef Use_memoValue As Boolean = False)
		'UPGRADE_WARNING: VarType has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		Select Case VarType(LoadedData)
			Case VariantType.Boolean
				LoadedData = Convert.ToBoolean(Rs1("lngValue").Value)
				'UPGRADE_WARNING: Untranslated statement in Database_LoadProperty. Please check source code.
			Case VariantType.Byte
				LoadedData = Convert.ToByte(Rs1("lngValue").Value)
				'UPGRADE_WARNING: Untranslated statement in Database_LoadProperty. Please check source code.
			Case VariantType.Short
				LoadedData = Convert.ToInt16(Rs1("lngValue").Value)
				'UPGRADE_WARNING: Untranslated statement in Database_LoadProperty. Please check source code.
			Case VariantType.Integer
				LoadedData = Convert.ToInt32(Rs1("lngValue").Value)
				'UPGRADE_WARNING: Untranslated statement in Database_LoadProperty. Please check source code.
			Case VariantType.String, VariantType.Date
				If (Use_memoValue) Then
					LoadedData = Convert.ToString(Rs1("memoValue").Value)
					'UPGRADE_WARNING: Untranslated statement in Database_LoadProperty. Please check source code.
				Else
					LoadedData = Convert.ToString(Rs1("strValue").Value)
					'UPGRADE_WARNING: Untranslated statement in Database_LoadProperty. Please check source code.
				End If
			Case VariantType.Double
				LoadedData = Convert.ToDouble(Rs1("dblValue").Value)
				'UPGRADE_WARNING: Untranslated statement in Database_LoadProperty. Please check source code.
			Case VariantType.Single
				LoadedData = Convert.ToSingle(Rs1("dblValue").Value)
				'UPGRADE_WARNING: Untranslated statement in Database_LoadProperty. Please check source code.
		End Select
	End Sub
	Sub Database_SaveProperty(ByRef Rs1 As dao.Recordset, ByRef in_Use_FieldName As String, ByRef SavedData As Object)
		Dim Use_memoValue As Boolean
		Dim Use_FieldName As String
		'NOTE: IF THE FIRST CHARACTER IS AN ASTERISK ("*"), THEN
		'THE FIELD TYPE USED IS THE MEMO TYPE (memoValue).
		Use_memoValue = False
		If (Left(in_Use_FieldName, 1) = "*") Then
			Use_FieldName = Right(in_Use_FieldName, Len(in_Use_FieldName) - 1)
			Use_memoValue = True
		Else
			Use_FieldName = in_Use_FieldName
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.AddNew. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1.AddNew()
		'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1("FieldName").Value = Use_FieldName
		'UPGRADE_WARNING: VarType has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		Select Case VarType(SavedData)
			Case VariantType.Boolean, VariantType.Byte, VariantType.Short, VariantType.Integer
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(lngValue). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object SavedData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("lngValue").Value = CInt(SavedData)
			Case VariantType.String, VariantType.Date
				If (Use_memoValue) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(memoValue). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object SavedData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1("memoValue").Value = CStr(SavedData)
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(strValue). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object SavedData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1("strValue").Value = CStr(SavedData)
				End If
			Case VariantType.Double, VariantType.Single
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(dblValue). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object SavedData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("dblValue").Value = CDbl(SavedData)
		End Select
		'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Update. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1.Update()
	End Sub
	Sub Database_SavePropertyWithIndex(ByRef Rs1 As dao.Recordset, ByRef in_Use_FieldName As String, ByRef Use_FieldIndex As Short, ByRef SavedData As Object)
		Dim Use_memoValue As Boolean
		Dim Use_FieldName As String
		'NOTE: IF THE FIRST CHARACTER IS AN ASTERISK ("*"), THEN
		'THE FIELD TYPE USED IS THE MEMO TYPE (memoValue).
		Use_memoValue = False
		If (Left(in_Use_FieldName, 1) = "*") Then
			Use_FieldName = Right(in_Use_FieldName, Len(in_Use_FieldName) - 1)
			Use_memoValue = True
		Else
			Use_FieldName = in_Use_FieldName
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.AddNew. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1.AddNew()
		'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1("FieldName").Value = Use_FieldName
		'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1("FieldIndex").Value = Use_FieldIndex
		'UPGRADE_WARNING: VarType has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		Select Case VarType(SavedData)
			Case VariantType.Boolean, VariantType.Byte, VariantType.Short, VariantType.Integer
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(lngValue). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object SavedData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("lngValue").Value = CInt(SavedData)
			Case VariantType.String, VariantType.Date
				If (Use_memoValue) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(memoValue). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object SavedData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1("memoValue").Value = CStr(SavedData)
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(strValue). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object SavedData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rs1("strValue").Value = CStr(SavedData)
				End If
			Case VariantType.Double, VariantType.Single
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(dblValue). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object SavedData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object Rs1(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rs1("dblValue").Value = CDbl(SavedData)
		End Select
		'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Update. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1.Update()
	End Sub


	Sub Database_DeleteTableContents(ByRef Db1 As dao.Database, ByRef TableName As String)
		'UPGRADE_ISSUE: Recordset object was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim Rs1 As dao.Recordset
		On Error GoTo err_Database_DeleteTableContents
		'UPGRADE_WARNING: Couldn't resolve default property of object Db1.OpenRecordset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1 = Db1.OpenRecordset(TableName)
		'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveFirst. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1.MoveFirst()
		'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.EOF. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Do Until Rs1.EOF
			'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Delete. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Rs1.Delete()
			'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.MoveNext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Rs1.MoveNext()
		Loop
		'UPGRADE_WARNING: Couldn't resolve default property of object Rs1.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Rs1.Close()
		Exit Sub
exit_err_Database_DeleteTableContents:
		Exit Sub
err_Database_DeleteTableContents:
		Resume exit_err_Database_DeleteTableContents
	End Sub
	Sub Database_CreateMFBTable(ByRef Db1 As dao.Database, ByRef TableName As String, ByRef Include_FieldIndex As Boolean, ByRef Include_FieldIndex2 As Boolean)
		'Dim dbMemo As Object
		'Dim dbDouble As Object
		'	Dim dbText As Object
		'	Dim dbLong As Object
		'UPGRADE_ISSUE: TableDef object was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim Td1 As dao.TableDef
		'UPGRADE_ISSUE: Field object was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim Ff As dao.Field

		'UPGRADE_WARNING: Couldn't resolve default property of object Db1.CreateTableDef. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Td1 = Db1.CreateTableDef(TableName)
		'UPGRADE_WARNING: Couldn't resolve default property of object Td1.CreateField. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Ff = Td1.CreateField("RecordID", dao.DataTypeEnum.dbLong)
		'TODO: ADD AUTONUMBER SETUP FOR THIS FIELD (NOT TOO NEEDED).
		'UPGRADE_WARNING: Couldn't resolve default property of object Td1.Fields. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Td1.Fields.Append(Ff)
		'UPGRADE_WARNING: Couldn't resolve default property of object Td1.CreateField. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Ff = Td1.CreateField("FieldName", dao.DataTypeEnum.dbText, 250)
		'UPGRADE_WARNING: Couldn't resolve default property of object Ff.AllowZeroLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Ff.AllowZeroLength = True
		'UPGRADE_WARNING: Couldn't resolve default property of object Td1.Fields. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Td1.Fields.Append(Ff)
		If (Include_FieldIndex) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Td1.CreateField. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Ff = Td1.CreateField("FieldIndex", dao.DataTypeEnum.dbLong)
			'UPGRADE_WARNING: Couldn't resolve default property of object Td1.Fields. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Td1.Fields.Append(Ff)
		End If
		If (Include_FieldIndex2) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Td1.CreateField. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Ff = Td1.CreateField("FieldIndex2", dao.DataTypeEnum.dbLong)
			'UPGRADE_WARNING: Couldn't resolve default property of object Td1.Fields. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Td1.Fields.Append(Ff)
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object Td1.CreateField. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Ff = Td1.CreateField("strValue", dao.DataTypeEnum.dbText, 250)
		'UPGRADE_WARNING: Couldn't resolve default property of object Ff.AllowZeroLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Ff.AllowZeroLength = True
		'UPGRADE_WARNING: Couldn't resolve default property of object Td1.Fields. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Td1.Fields.Append(Ff)
		'UPGRADE_WARNING: Couldn't resolve default property of object Td1.CreateField. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Ff = Td1.CreateField("dblValue", dao.DataTypeEnum.dbDouble)
		'UPGRADE_WARNING: Couldn't resolve default property of object Td1.Fields. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Td1.Fields.Append(Ff)
		'UPGRADE_WARNING: Couldn't resolve default property of object Td1.CreateField. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Ff = Td1.CreateField("lngValue", dao.DataTypeEnum.dbLong)
		'UPGRADE_WARNING: Couldn't resolve default property of object Td1.Fields. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Td1.Fields.Append(Ff)
		'UPGRADE_WARNING: Couldn't resolve default property of object Td1.CreateField. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Ff = Td1.CreateField("memoValue", dao.DataTypeEnum.dbMemo)
		'UPGRADE_WARNING: Couldn't resolve default property of object Ff.AllowZeroLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Ff.AllowZeroLength = True
		'UPGRADE_WARNING: Couldn't resolve default property of object Td1.Fields. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Td1.Fields.Append(Ff)
		'UPGRADE_WARNING: Couldn't resolve default property of object Td1.CreateField. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Ff = Td1.CreateField("Comments", dao.DataTypeEnum.dbText, 250)
		'UPGRADE_WARNING: Couldn't resolve default property of object Ff.AllowZeroLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Ff.AllowZeroLength = True
		'UPGRADE_WARNING: Couldn't resolve default property of object Td1.Fields. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Td1.Fields.Append(Ff)
		'UPGRADE_WARNING: Couldn't resolve default property of object Db1.TableDefs. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Db1.TableDefs.Append(Td1)
	End Sub

	Sub Database_CreateMFBTable_IfNoExist(Db1 As dao.Database, Use_TableName As String, Include_FieldIndex As Boolean)

		Dim tbl As dao.TableDef
		Dim TableExists As Boolean = False
		For Each tbl In Db1.TableDefs
			If tbl.Name = Use_TableName Then
				TableExists = True
				Exit For
			End If
		Next
		If (TableExists = False) Then
			Call Database_CreateMFBTable(Db1, Use_TableName, Include_FieldIndex, False)
		End If
	End Sub
	Sub Database_CreateMFBTable_IfNoExist_TwoIndices(Db1 As dao.Database, Use_TableName As String)

		Dim tbl As dao.TableDef
		Dim TableExists As Boolean = False
		For Each tbl In Db1.TableDefs
			If tbl.Name = Use_TableName Then
				TableExists = True
				Exit For
			End If
		Next
		If (TableExists = False) Then
			Call Database_CreateMFBTable(Db1, Use_TableName, True, True)
		End If
	End Sub



End Module