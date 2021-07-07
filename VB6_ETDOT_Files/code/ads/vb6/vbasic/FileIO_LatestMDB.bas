Attribute VB_Name = "FileIO_LatestMDB"
Option Explicit





Const FileIO_LatestMDB_declarations_end = True


'RETURNS:
'         TRUE = SUCCEEDED IN LOADING.
'         FALSE = FAILED IN LOADING.
Function File_Open_Latest_v1_60( _
    fn_This As String) As Boolean
Dim Ws1 As Workspace
Dim Db1 As Database
Dim Rs1 As Recordset
Dim Use_FieldIndex As Integer
Dim Use_FieldIndex2 As Integer
Dim ContainsTable_PSDMInRoomData As Boolean
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
  Set Db1 = OpenDatabase(fn_This)

  '=========== INPUT DATA FROM DATABASE TABLES. =================
  
  '------ INPUT DATA FROM TABLE "Version". ------------------------------------------------------------------------------------------------------
  'APPLICABLE DEFAULT VALUES:
  ContainsTable_PSDMInRoomData = False
  If (Database_IsTableExist(Db1, "Version") = False) Then
    'DO NOTHING: USE DEFAULT VALUES.
  Else
    Set Rs1 = Db1.OpenRecordset("Version")
    If (Database_NoRecordsInRecordset(Rs1)) Then
      'DO NOTHING: USE DEFAULT VALUES.
    Else
      Rs1.MoveFirst
      Do Until Rs1.EOF
        ''''Use_FieldIndex = CInt(Database_Get_Long(Rs1, "FieldIndex"))
        Select Case Trim$(UCase$(Database_Get_String(Rs1, "FieldName")))
          'HEADER BLOCK.
          Case Trim$(UCase$("ContainsTable_PSDMInRoomData")): Call Database_LoadProperty(Rs1, ContainsTable_PSDMInRoomData)
        End Select
        Rs1.MoveNext
      Loop
    End If
    Rs1.Close
  End If
  
  '------ INPUT DATA FROM TABLE "Main". ------------------------------------------------------------------------------------------------------
Dim booDoDemoCalc As Boolean
Dim dblDemoChecksum As Double
Dim dblThisVal As Double
Dim lngThisVal As Long
  booDoDemoCalc = False
  dblDemoChecksum = 0#
  If (IsThisADemo() = True) Then
    ' STORE TO VARIABLE TO SAVE A LITTLE TIME.
    booDoDemoCalc = True
  End If
  If (Database_IsTableExist(Db1, "Main") = False) Then
    'DO NOTHING: USE DEFAULT VALUES.
  Else
    Set Rs1 = Db1.OpenRecordset("Main")
    If (Database_NoRecordsInRecordset(Rs1)) Then
      'DO NOTHING: USE DEFAULT VALUES.
    Else
      Rs1.MoveFirst
      Do Until Rs1.EOF
        Use_FieldIndex = CInt(Database_Get_Long(Rs1, "FieldIndex"))
        If (booDoDemoCalc = True) Then
          Call Database_LoadProperty(Rs1, dblThisVal)
          Call Database_LoadProperty(Rs1, lngThisVal)
          If (dblThisVal = 0#) Then dblThisVal = CDbl(lngThisVal)
          If (dblThisVal = 0#) Then dblThisVal = 0.1
          dblDemoChecksum = dblDemoChecksum + Abs(Log(Abs(dblThisVal)))
        End If
        Select Case Trim$(UCase$(Database_Get_String(Rs1, "FieldName")))
          'HEADER BLOCK.
          Case Trim$(UCase$("FileNote")): Call Database_LoadProperty(Rs1, FileNote, True)
          Case Trim$(UCase$("Number_Component")): Call Database_LoadProperty(Rs1, Number_Component)
          'COMPONENT PROPERTIES BLOCK.
          Case Trim$(UCase$("Co.Name")): Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).Name)
          Case Trim$(UCase$("Co.CAS")): Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).CAS)
          Case Trim$(UCase$("Co.MW")): Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).MW)
          Case Trim$(UCase$("Co.InitialConcentration")): Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).InitialConcentration)
          Case Trim$(UCase$("Co.MolarVolume")): Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).MolarVolume)
          Case Trim$(UCase$("Co.BP")): Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).BP)
          Case Trim$(UCase$("Co.Use_K")): Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).Use_K)
          Case Trim$(UCase$("Co.Use_OneOverN")): Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).Use_OneOverN)
          Case Trim$(UCase$("Co.Liquid_Density")): Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).Liquid_Density)
          Case Trim$(UCase$("Co.Aqueous_Solubility")): Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).Aqueous_Solubility)
          Case Trim$(UCase$("Co.Vapor_Pressure")): Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).Vapor_Pressure)
          Case Trim$(UCase$("Co.Refractive_Index")): Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).Refractive_Index)
          Case Trim$(UCase$("Co.SPDFR")): Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).SPDFR)
          Case Trim$(UCase$("Co.SPDFR_Low_Concentration")): Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).SPDFR_Low_Concentration)
          Case Trim$(UCase$("Co.Use_SPDFR_Correlation")): Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).Use_SPDFR_Correlation)
          Case Trim$(UCase$("Co.kf")): Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).kf)
          Case Trim$(UCase$("Co.Ds")): Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).Ds)
          Case Trim$(UCase$("Co.Dp")): Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).Dp)
          Case Trim$(UCase$("Co.Corr(1)")): Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).Corr(1))
          Case Trim$(UCase$("Co.Corr(2)")): Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).Corr(2))
          Case Trim$(UCase$("Co.Corr(3)")): Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).Corr(3))
          Case Trim$(UCase$("Co.KP_User_Input(1)")): Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).KP_User_Input(1))
          Case Trim$(UCase$("Co.KP_User_Input(2)")): Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).KP_User_Input(2))
          Case Trim$(UCase$("Co.KP_User_Input(3)")): Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).KP_User_Input(3))
          Case Trim$(UCase$("Co.K_Reduction")): Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).K_Reduction)
          Case Trim$(UCase$("Co.Correlation.Name")): Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).Correlation.Name)
          Case Trim$(UCase$("Co.Correlation.Coeff(1)")): Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).Correlation.Coeff(1))
          Case Trim$(UCase$("Co.Correlation.Coeff(2)")): Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).Correlation.Coeff(2))
          Case Trim$(UCase$("Co.IsothermDB_Component_Name")): Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).IsothermDB_Component_Name)
          Case Trim$(UCase$("Co.IsothermDB_Range_Num")): Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).IsothermDB_Range_Num)
          Case Trim$(UCase$("Co.IPES_OrderOfMagnitude")): Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).IPES_OrderOfMagnitude)
          Case Trim$(UCase$("Co.IPES_NumRegressionPts")): Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).IPES_NumRegressionPts)
          Case Trim$(UCase$("Co.IPES_RelativeHumidity")): Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).IPES_RelativeHumidity)
          Case Trim$(UCase$("Co.IPES_EstimationMethod")): Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).IPES_EstimationMethod)
          Case Trim$(UCase$("Co.Source_KandOneOverN")): Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).Source_KandOneOverN)
          Case Trim$(UCase$("Co.IsothermDB_K")): Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).IsothermDB_K)
          Case Trim$(UCase$("Co.IsothermDB_OneOverN")): Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).IsothermDB_OneOverN)
          Case Trim$(UCase$("Co.IPESResult_K")): Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).IPESResult_K)
          Case Trim$(UCase$("Co.IPESResult_OneOverN")): Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).IPESResult_OneOverN)
          Case Trim$(UCase$("Co.UserEntered_K")): Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).UserEntered_K)
          Case Trim$(UCase$("Co.UserEntered_OneOverN")): Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).UserEntered_OneOverN)
          Case Trim$(UCase$("Co.Tortuosity")): Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).Tortuosity)
          Case Trim$(UCase$("Co.Use_Tortuosity_Correlation")): Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).Use_Tortuosity_Correlation)
          Case Trim$(UCase$("Co.Constant_Tortuosity")): Call Database_LoadProperty(Rs1, Component(Use_FieldIndex).Constant_Tortuosity)
          'BED PROPERTIES BLOCK.
          Case Trim$(UCase$("Be.length")): Call Database_LoadProperty(Rs1, Bed.length)
          Case Trim$(UCase$("Be.Diameter")): Call Database_LoadProperty(Rs1, Bed.Diameter)
          Case Trim$(UCase$("Be.Weight")): Call Database_LoadProperty(Rs1, Bed.Weight)
          Case Trim$(UCase$("Be.Flowrate")): Call Database_LoadProperty(Rs1, Bed.Flowrate)
          Case Trim$(UCase$("Be.WaterDensity")): Call Database_LoadProperty(Rs1, Bed.WaterDensity)
          Case Trim$(UCase$("Be.WaterViscosity")): Call Database_LoadProperty(Rs1, Bed.WaterViscosity)
          Case Trim$(UCase$("Be.Temperature")): Call Database_LoadProperty(Rs1, Bed.Temperature)
          Case Trim$(UCase$("Be.Pressure")): Call Database_LoadProperty(Rs1, Bed.Pressure)
          Case Trim$(UCase$("Be.Phase")): Call Database_LoadProperty(Rs1, Bed.Phase)
          Case Trim$(UCase$("Be.NumberOfBeds")): Call Database_LoadProperty(Rs1, Bed.NumberOfBeds)
          Case Trim$(UCase$("Be.Water_Correlation.Name")): Call Database_LoadProperty(Rs1, Bed.Water_Correlation.Name)
          Case Trim$(UCase$("Be.Water_Correlation.Coeff(1)")): Call Database_LoadProperty(Rs1, Bed.Water_Correlation.Coeff(1))
          Case Trim$(UCase$("Be.Water_Correlation.Coeff(2)")): Call Database_LoadProperty(Rs1, Bed.Water_Correlation.Coeff(2))
          Case Trim$(UCase$("Be.Water_Correlation.Coeff(3)")): Call Database_LoadProperty(Rs1, Bed.Water_Correlation.Coeff(3))
          Case Trim$(UCase$("Be.Water_Correlation.Coeff(4)")): Call Database_LoadProperty(Rs1, Bed.Water_Correlation.Coeff(4))
          'UNITS BLOCK.
          Case Trim$(UCase$("frmMain.txtBedUnits(0)")): Call Units1_Database_LoadProperty(Rs1, frmMain.txtBedUnits(0))
          Case Trim$(UCase$("frmMain.txtBedUnits(1)")): Call Units1_Database_LoadProperty(Rs1, frmMain.txtBedUnits(1))
          Case Trim$(UCase$("frmMain.txtBedUnits(2)")): Call Units1_Database_LoadProperty(Rs1, frmMain.txtBedUnits(2))
          Case Trim$(UCase$("frmMain.txtBedUnits(3)")): Call Units1_Database_LoadProperty(Rs1, frmMain.txtBedUnits(3))
          Case Trim$(UCase$("frmMain.txtBedUnits(4)")): Call Units1_Database_LoadProperty(Rs1, frmMain.txtBedUnits(4))
          Case Trim$(UCase$("frmMain.txtCarbonUnits(1)")): Call Units1_Database_LoadProperty(Rs1, frmMain.txtCarbonUnits(1))
          Case Trim$(UCase$("frmMain.txtCarbonUnits(2)")): Call Units1_Database_LoadProperty(Rs1, frmMain.txtCarbonUnits(2))
          Case Trim$(UCase$("frmMain.txtTimeUnits(0)")): Call Units1_Database_LoadProperty(Rs1, frmMain.txtTimeUnits(0))
          Case Trim$(UCase$("frmMain.txtTimeUnits(1)")): Call Units1_Database_LoadProperty(Rs1, frmMain.txtTimeUnits(1))
          Case Trim$(UCase$("frmMain.txtTimeUnits(2)")): Call Units1_Database_LoadProperty(Rs1, frmMain.txtTimeUnits(2))
          Case Trim$(UCase$("PropertyUnits.MW")): Call Database_LoadProperty(Rs1, PropertyUnits.MW)
          Case Trim$(UCase$("PropertyUnits.MolarVolume")): Call Database_LoadProperty(Rs1, PropertyUnits.MolarVolume)
          Case Trim$(UCase$("PropertyUnits.BP")): Call Database_LoadProperty(Rs1, PropertyUnits.BP)
          Case Trim$(UCase$("PropertyUnits.InitialConcentration")): Call Database_LoadProperty(Rs1, PropertyUnits.InitialConcentration)
          Case Trim$(UCase$("PropertyUnits.Liquid_Density")): Call Database_LoadProperty(Rs1, PropertyUnits.Liquid_Density)
          Case Trim$(UCase$("PropertyUnits.Aqueous_Solubility")): Call Database_LoadProperty(Rs1, PropertyUnits.Aqueous_Solubility)
          Case Trim$(UCase$("PropertyUnits.Vapor_Pressure")): Call Database_LoadProperty(Rs1, PropertyUnits.Vapor_Pressure)
          Case Trim$(UCase$("PropertyUnits.k")): Call Database_LoadProperty(Rs1, PropertyUnits.k)
          'MISCELLANEOUS BLOCK.
          Case Trim$(UCase$("Carbon.Name")): Call Database_LoadProperty(Rs1, Carbon.Name)
          Case Trim$(UCase$("Carbon.Porosity")): Call Database_LoadProperty(Rs1, Carbon.Porosity)
          Case Trim$(UCase$("Carbon.Density")): Call Database_LoadProperty(Rs1, Carbon.Density)
          Case Trim$(UCase$("Carbon.ParticleRadius")): Call Database_LoadProperty(Rs1, Carbon.ParticleRadius)
          Case Trim$(UCase$("Carbon.Tortuosity")): Call Database_LoadProperty(Rs1, Carbon.Tortuosity)
          Case Trim$(UCase$("Carbon.W0")): Call Database_LoadProperty(Rs1, Carbon.W0)
          Case Trim$(UCase$("Carbon.BB")): Call Database_LoadProperty(Rs1, Carbon.BB)
          Case Trim$(UCase$("Carbon.PolanyiExponent")): Call Database_LoadProperty(Rs1, Carbon.PolanyiExponent)
          Case Trim$(UCase$("State_Check_Water(1)")): Call Database_LoadProperty(Rs1, State_Check_Water(1))
          Case Trim$(UCase$("State_Check_Water(2)")): Call Database_LoadProperty(Rs1, State_Check_Water(2))
          Case Trim$(UCase$("Carbon.ShapeFactor")): Call Database_LoadProperty(Rs1, Carbon.ShapeFactor)
          Case Trim$(UCase$("Constant_Tortuosity")): Call Database_LoadProperty(Rs1, Constant_Tortuosity)
          Case Trim$(UCase$("Carbon.ShapeFactor")): Call Database_LoadProperty(Rs1, Carbon.ShapeFactor)
          Case Trim$(UCase$("NC")): Call Database_LoadProperty(Rs1, NC)
          Case Trim$(UCase$("MC")): Call Database_LoadProperty(Rs1, MC)
          Case Trim$(UCase$("TimeP.Init")): Call Database_LoadProperty(Rs1, TimeP.Init)
          Case Trim$(UCase$("TimeP.End")): Call Database_LoadProperty(Rs1, TimeP.End)
          Case Trim$(UCase$("TimeP.np")): Call Database_LoadProperty(Rs1, TimeP.np)
          Case Trim$(UCase$("TimeP.Step")): Call Database_LoadProperty(Rs1, TimeP.Step)
          'INFLUENT/EFFLUENT POINT COUNTS.
          Case Trim$(UCase$("Number_Influent_Points")): Call Database_LoadProperty(Rs1, Number_Influent_Points)
          Case Trim$(UCase$("NData_Points")): Call Database_LoadProperty(Rs1, NData_Points)
        End Select
        Rs1.MoveNext
      Loop
    End If
    Rs1.Close
  End If
  If (booDoDemoCalc = True) Then
    ''''Call Show_Message("Demo Version value of dblDemoChecksum = " & _
        Trim$(Str$(dblDemoChecksum)) & ".")
    If (Demo_CheckForValidFile(dblDemoChecksum) = False) Then
      Call file_new
      Call Demo_ShowError("In the demonstration version, only the example files may be opened.")
      File_Open_Latest_v1_60 = False
      Exit Function
    End If
  End If

  '------ INPUT DATA FROM TABLE "InfluentPoints". ------------------------------------------------------------------------------------------------------
  If (Database_IsTableExist(Db1, "InfluentPoints") = False) Then
    'DO NOTHING: USE DEFAULT VALUES.
  Else
    Set Rs1 = Db1.OpenRecordset("InfluentPoints")
    If (Database_NoRecordsInRecordset(Rs1)) Then
      'DO NOTHING: USE DEFAULT VALUES.
    Else
      Rs1.MoveFirst
      Do Until Rs1.EOF
        Use_FieldIndex = CInt(Database_Get_Long(Rs1, "FieldIndex"))
        Use_FieldIndex2 = CInt(Database_Get_Long(Rs1, "FieldIndex2"))
        Select Case Trim$(UCase$(Database_Get_String(Rs1, "FieldName")))
          Case Trim$(UCase$("T_Influent")): Call Database_LoadProperty(Rs1, T_Influent(Use_FieldIndex))
          Case Trim$(UCase$("C_Influent")): Call Database_LoadProperty(Rs1, C_Influent(Use_FieldIndex, Use_FieldIndex2))
        End Select
        Rs1.MoveNext
      Loop
    End If
    Rs1.Close
  End If

  '------ INPUT DATA FROM TABLE "EffluentPoints". ------------------------------------------------------------------------------------------------------
  If (Database_IsTableExist(Db1, "EffluentPoints") = False) Then
    'DO NOTHING: USE DEFAULT VALUES.
  Else
    Set Rs1 = Db1.OpenRecordset("EffluentPoints")
    If (Database_NoRecordsInRecordset(Rs1)) Then
      'DO NOTHING: USE DEFAULT VALUES.
    Else
      Rs1.MoveFirst
      Do Until Rs1.EOF
        Use_FieldIndex = CInt(Database_Get_Long(Rs1, "FieldIndex"))
        Use_FieldIndex2 = CInt(Database_Get_Long(Rs1, "FieldIndex2"))
        Select Case Trim$(UCase$(Database_Get_String(Rs1, "FieldName")))
          Case Trim$(UCase$("T_Data_Points")): Call Database_LoadProperty(Rs1, T_Data_Points(Use_FieldIndex))
          Case Trim$(UCase$("C_Data_Points")): Call Database_LoadProperty(Rs1, C_Data_Points(Use_FieldIndex, Use_FieldIndex2))
        End Select
        Rs1.MoveNext
      Loop
    End If
    Rs1.Close
  End If
  
  '------ INPUT DATA FROM TABLE "PSDMInRoomData". ------------------------------------------------------------------------------------------------------
  If (ContainsTable_PSDMInRoomData) Then
    rp = RoomParams
    If (Database_IsTableExist(Db1, "PSDMInRoomData") = False) Then
      'DO NOTHING: USE DEFAULT VALUES.
    Else
      Set Rs1 = Db1.OpenRecordset("PSDMInRoomData")
      If (Database_NoRecordsInRecordset(Rs1)) Then
        'DO NOTHING: USE DEFAULT VALUES.
      Else
        Rs1.MoveFirst
        Do Until Rs1.EOF
          Use_FieldIndex = CInt(Database_Get_Long(Rs1, "FieldIndex"))
          Select Case Trim$(UCase$(Database_Get_String(Rs1, "FieldName")))
            ' INPUT ROOM PARAMETERS.
            Case Trim$(UCase$("RP.COUNT_CONTAMINANT")): Call Database_LoadProperty(Rs1, rp.COUNT_CONTAMINANT)
            Case Trim$(UCase$("RP.ROOM_VOL")): Call Database_LoadProperty(Rs1, rp.ROOM_VOL)
            Case Trim$(UCase$("RP.ROOM_FLOWRATE")): Call Database_LoadProperty(Rs1, rp.ROOM_FLOWRATE)
            Case Trim$(UCase$("RP.ROOM_C0")): Call Database_LoadProperty(Rs1, rp.ROOM_C0(Use_FieldIndex))
            Case Trim$(UCase$("RP.ROOM_EMIT")): Call Database_LoadProperty(Rs1, rp.ROOM_EMIT(Use_FieldIndex))
            ' CALCULATED ROOM PARAMETERS.
            Case Trim$(UCase$("RP.ROOM_CHANGE_RATE")): Call Database_LoadProperty(Rs1, rp.ROOM_CHANGE_RATE)
            Case Trim$(UCase$("RP.ROOM_SS_VALUE")): Call Database_LoadProperty(Rs1, rp.ROOM_SS_VALUE(Use_FieldIndex))
            ' UNITS FOR ALL VARIABLES.
            Case Trim$(UCase$("RP.ROOM_VOL_Units")): Call Database_LoadProperty(Rs1, rp.ROOM_VOL_Units)
            Case Trim$(UCase$("RP.ROOM_FLOWRATE_Units")): Call Database_LoadProperty(Rs1, rp.ROOM_FLOWRATE_Units)
            Case Trim$(UCase$("RP.ROOM_C0_Units")): Call Database_LoadProperty(Rs1, rp.ROOM_C0_Units)
            Case Trim$(UCase$("RP.ROOM_EMIT_Units")): Call Database_LoadProperty(Rs1, rp.ROOM_EMIT_Units)
            Case Trim$(UCase$("RP.INITIAL_ROOM_CONC_Units")): Call Database_LoadProperty(Rs1, rp.INITIAL_ROOM_CONC_Units)
            ' NEW AS OF 9/16/98.
            Case Trim$(UCase$("RP.INITIAL_ROOM_CONC")): Call Database_LoadProperty(Rs1, rp.INITIAL_ROOM_CONC(Use_FieldIndex))
            ' NEW AS OF 9/16/98 ENDS.
            ' NEW AS OF 8/18/99.
            Case Trim$(UCase$("RP.RXN_RATE_CONSTANT")): Call Database_LoadProperty(Rs1, rp.RXN_RATE_CONSTANT(Use_FieldIndex))
            Case Trim$(UCase$("RP.RXN_PRODUCT")): Call Database_LoadProperty(Rs1, rp.RXN_PRODUCT(Use_FieldIndex))
            Case Trim$(UCase$("RP.RXN_RATIO")): Call Database_LoadProperty(Rs1, rp.RXN_RATIO(Use_FieldIndex))
            ' NEW AS OF 8/18/99 ENDS.
            '---- NEW AS OF 11/11/99 BEGINS: ---------------------------------------------------------
            '
            '/////////   TIME-VARIABLE Co   //////////////////////////////////
            Case Trim$(UCase$("RP.bool_ROOM_COINI_ISTIMEVAR")): Call Database_LoadProperty(Rs1, rp.bool_ROOM_COINI_ISTIMEVAR(Use_FieldIndex))
            Case Trim$(UCase$("RP.int_ROOM_NCOINI")): Call Database_LoadProperty(Rs1, rp.int_ROOM_NCOINI(Use_FieldIndex))
            Case Trim$(UCase$("RP.u_ROOM_TCOINI")): Call Database_LoadProperty(Rs1, rp.u_ROOM_TCOINI)
            Case Trim$(UCase$("RP.u_ROOM_COINI")): Call Database_LoadProperty(Rs1, rp.u_ROOM_COINI)
            ''''dbl_ROOM_TCOINI() As Double   '(x,y): x=chemical, y=row
            ''''dbl_ROOM_COINI() As Double    '(x,y): x=chemical, y=row
            '
            '/////////   TIME-VARIABLE w*A   /////////////////////////////////
            Case Trim$(UCase$("RP.bool_ROOM_EMITINI_ISTIMEVAR")): Call Database_LoadProperty(Rs1, rp.bool_ROOM_EMITINI_ISTIMEVAR(Use_FieldIndex))
            Case Trim$(UCase$("RP.int_ROOM_NEMITINI")): Call Database_LoadProperty(Rs1, rp.int_ROOM_NEMITINI(Use_FieldIndex))
            Case Trim$(UCase$("RP.u_ROOM_TEMITINI")): Call Database_LoadProperty(Rs1, rp.u_ROOM_TEMITINI)
            Case Trim$(UCase$("RP.u_ROOM_EMITINI")): Call Database_LoadProperty(Rs1, rp.u_ROOM_EMITINI)
            ''''dbl_ROOM_TEMITINI() As Double   '(x,y): x=chemical, y=row
            ''''dbl_ROOM_EMITINI() As Double    '(x,y): x=chemical, y=row
            '---- NEW AS OF 11/11/99 ENDS. ---------------------------------------------------------
            '---- NEW AS OF 1/17/00 BEGINS: ---------------------------------------------------------
            '
            '/////////   TIME-VARIABLE K   /////////////////////////////////
            Case Trim$(UCase$("RP.bool_ROOM_KINI_ISTIMEVAR")): Call Database_LoadProperty(Rs1, rp.bool_ROOM_KINI_ISTIMEVAR(Use_FieldIndex))
            Case Trim$(UCase$("RP.int_ROOM_NKINI")): Call Database_LoadProperty(Rs1, rp.int_ROOM_NKINI(Use_FieldIndex))
            Case Trim$(UCase$("RP.u_ROOM_TKINI")): Call Database_LoadProperty(Rs1, rp.u_ROOM_TKINI)
            Case Trim$(UCase$("RP.u_ROOM_KINI")): Call Database_LoadProperty(Rs1, rp.u_ROOM_KINI)
            ''''dbl_ROOM_TKINI() As Double   '(x,y): x=chemical, y=row
            ''''dbl_ROOM_KINI() As Double    '(x,y): x=chemical, y=row
            '---- NEW AS OF 1/17/00 ENDS. ---------------------------------------------------------
          End Select
          Rs1.MoveNext
        Loop
      End If
      Rs1.Close
    End If
    RoomParams = rp
    Call RoomParam_Recalculate(RoomParams)
  End If
  
  '---- NEW AS OF 11/11/99 BEGINS: ---------------------------------------------------------
  '------ INPUT DATA FROM TABLE "PSDMInRoomData_CO_Data". ------------------------------------------------------------------------------------------------------
  If (Database_IsTableExist(Db1, "PSDMInRoomData_CO_Data") = False) Then
    'DO NOTHING: USE DEFAULT VALUES.
  Else
    Set Rs1 = Db1.OpenRecordset("PSDMInRoomData_CO_Data")
    If (Database_NoRecordsInRecordset(Rs1)) Then
      'DO NOTHING: USE DEFAULT VALUES.
    Else
      Rs1.MoveFirst
      Do Until Rs1.EOF
        Use_FieldIndex = CInt(Database_Get_Long(Rs1, "FieldIndex"))
        Use_FieldIndex2 = CInt(Database_Get_Long(Rs1, "FieldIndex2"))
        Select Case Trim$(UCase$(Database_Get_String(Rs1, "FieldName")))
          Case Trim$(UCase$("dbl_ROOM_TCOINI")): Call Database_LoadProperty(Rs1, RoomParams.dbl_ROOM_TCOINI(Use_FieldIndex, Use_FieldIndex2))
          Case Trim$(UCase$("dbl_ROOM_COINI")): Call Database_LoadProperty(Rs1, RoomParams.dbl_ROOM_COINI(Use_FieldIndex, Use_FieldIndex2))
        End Select
        Rs1.MoveNext
      Loop
    End If
    Rs1.Close
  End If
  '---- NEW AS OF 11/11/99 ENDS. ---------------------------------------------------------
  
  '---- NEW AS OF 11/11/99 BEGINS: ---------------------------------------------------------
  '------ INPUT DATA FROM TABLE "PSDMInRoomData_WA_Data". ------------------------------------------------------------------------------------------------------
  If (Database_IsTableExist(Db1, "PSDMInRoomData_WA_Data") = False) Then
    'DO NOTHING: USE DEFAULT VALUES.
  Else
    Set Rs1 = Db1.OpenRecordset("PSDMInRoomData_WA_Data")
    If (Database_NoRecordsInRecordset(Rs1)) Then
      'DO NOTHING: USE DEFAULT VALUES.
    Else
      Rs1.MoveFirst
      Do Until Rs1.EOF
        Use_FieldIndex = CInt(Database_Get_Long(Rs1, "FieldIndex"))
        Use_FieldIndex2 = CInt(Database_Get_Long(Rs1, "FieldIndex2"))
        Select Case Trim$(UCase$(Database_Get_String(Rs1, "FieldName")))
          Case Trim$(UCase$("dbl_ROOM_TEMITINI")): Call Database_LoadProperty(Rs1, RoomParams.dbl_ROOM_TEMITINI(Use_FieldIndex, Use_FieldIndex2))
          Case Trim$(UCase$("dbl_ROOM_EMITINI")): Call Database_LoadProperty(Rs1, RoomParams.dbl_ROOM_EMITINI(Use_FieldIndex, Use_FieldIndex2))
        End Select
        Rs1.MoveNext
      Loop
    End If
    Rs1.Close
  End If
  '---- NEW AS OF 11/11/99 ENDS. ---------------------------------------------------------
  
  '---- NEW AS OF 1/17/00 BEGINS: ---------------------------------------------------------
  '------ INPUT DATA FROM TABLE "PSDMInRoomData_K_Data". ------------------------------------------------------------------------------------------------------
  If (Database_IsTableExist(Db1, "PSDMInRoomData_K_Data") = False) Then
    'DO NOTHING: USE DEFAULT VALUES.
  Else
    Set Rs1 = Db1.OpenRecordset("PSDMInRoomData_K_Data")
    If (Database_NoRecordsInRecordset(Rs1)) Then
      'DO NOTHING: USE DEFAULT VALUES.
    Else
      Rs1.MoveFirst
      Do Until Rs1.EOF
        Use_FieldIndex = CInt(Database_Get_Long(Rs1, "FieldIndex"))
        Use_FieldIndex2 = CInt(Database_Get_Long(Rs1, "FieldIndex2"))
        Select Case Trim$(UCase$(Database_Get_String(Rs1, "FieldName")))
          Case Trim$(UCase$("dbl_ROOM_TKINI")): Call Database_LoadProperty(Rs1, RoomParams.dbl_ROOM_TKINI(Use_FieldIndex, Use_FieldIndex2))
          Case Trim$(UCase$("dbl_ROOM_KINI")): Call Database_LoadProperty(Rs1, RoomParams.dbl_ROOM_KINI(Use_FieldIndex, Use_FieldIndex2))
        End Select
        Rs1.MoveNext
      Loop
    End If
    Rs1.Close
  End If
  '---- NEW AS OF 1/17/00 ENDS. ---------------------------------------------------------
  
  'CLOSE THE DATABASE FILE.
  Db1.Close

  'RETURN A "SUCCESS" MESSAGE TO CALLER.
  File_Open_Latest_v1_60 = True

End Function


'RETURNS:
'         TRUE = SUCCEEDED IN SAVING.
'         FALSE = FAILED IN SAVING.
Function File_Save_Latest_v1_60( _
    fn_This As String) As Boolean
Dim Ws1 As Workspace
Dim Db1 As Database
Dim Rs1 As Recordset
Dim i As Integer
Dim J As Integer
Dim Co As ComponentPropertyType
Dim Be As BedPropertyType
Dim IsLegacyVersion As Boolean
Dim NeedToCreateNewDatabase As Boolean
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
    Set Db1 = OpenDatabase(fn_This)
    If (Err.number = 0) Then
      IsLegacyVersion = False
      Db1.Close
    End If
    On Error GoTo 0
    If (IsLegacyVersion) Then
      'DELETE OLD FILE, CREATE NEW DATABASE (SEE BELOW).
      Kill fn_This
      NeedToCreateNewDatabase = True
    Else
      'OPEN DATABASE NORMALLY.
      Set Db1 = OpenDatabase(fn_This)
    End If
  End If
  If (NeedToCreateNewDatabase) Then
    FileCopy MAIN_APP_PATH & "\dbase\template.dat", fn_This
    Set Db1 = OpenDatabase(fn_This)
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
  Set Rs1 = Db1.OpenRecordset("Version")
  Call Database_SaveProperty(Rs1, "DataVersion_Major", CInt(Latest_DataVersion_Major))
  Call Database_SaveProperty(Rs1, "DataVersion_Minor", CInt(Latest_DataVersion_Minor))
  Call Database_SaveProperty(Rs1, "ContainsTable_PSDMInRoomData", True)
  'END SAVE TO THIS TABLE.
  Rs1.Close
  
  '------ OUTPUT DATA TO TABLE "Main". ------------------------------------------------------------------------------------------------------
  'START SAVE TO THIS TABLE.
  Call Database_DeleteTableContents(Db1, "Main")
  Set Rs1 = Db1.OpenRecordset("Main")
  'HEADER BLOCK.
  Call Database_SaveProperty(Rs1, "*FileNote", FileNote)
  Call Database_SaveProperty(Rs1, "Number_Component", Number_Component)
  'COMPONENT PROPERTIES BLOCK.
  For i = 1 To Number_Component
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
  Call Database_SaveProperty(Rs1, "TimeP.End", TimeP.End)
  Call Database_SaveProperty(Rs1, "TimeP.np", TimeP.np)
  Call Database_SaveProperty(Rs1, "TimeP.Step", TimeP.Step)
  'INFLUENT/EFFLUENT POINT COUNTS.
  Call Database_SaveProperty(Rs1, "Number_Influent_Points", Number_Influent_Points)
  Call Database_SaveProperty(Rs1, "NData_Points", NData_Points)
  'END SAVE TO THIS TABLE.
  Rs1.Close
  
  '------ OUTPUT DATA TO TABLE "InfluentPoints". ------------------------------------------------------------------------------------------------------
  'START SAVE TO THIS TABLE.
  Call Database_DeleteTableContents(Db1, "InfluentPoints")
  Set Rs1 = Db1.OpenRecordset("InfluentPoints")
  'MAIN DATA SET.
  If (Number_Influent_Points > 0) Then
    For i = 1 To Number_Influent_Points
      ''''Write #f, T_Influent(i)
      Rs1.AddNew
      Rs1("FieldName") = "T_Influent"
      Rs1("FieldIndex") = i
      Rs1("dblValue") = T_Influent(i)
      Rs1.Update
      For J = 1 To Number_Component
        ''''Write #f, C_Influent(j, i)
        Rs1.AddNew
        Rs1("FieldName") = "C_Influent"
        Rs1("FieldIndex") = J
        Rs1("FieldIndex2") = i
        Rs1("dblValue") = C_Influent(J, i)
        Rs1.Update
      Next J
    Next i
  End If
  'END SAVE TO THIS TABLE.
  Rs1.Close
  
  '------ OUTPUT DATA TO TABLE "EffluentPoints". ------------------------------------------------------------------------------------------------------
  'START SAVE TO THIS TABLE.
  Call Database_DeleteTableContents(Db1, "EffluentPoints")
  Set Rs1 = Db1.OpenRecordset("EffluentPoints")
  'MAIN DATA SET.
  If (NData_Points > 0) Then
    For i = 1 To NData_Points
      ''''Write #f, T_Data_Points(i)
      Rs1.AddNew
      Rs1("FieldName") = "T_Data_Points"
      Rs1("FieldIndex") = i
      Rs1("dblValue") = T_Data_Points(i)
      Rs1.Update
      For J = 1 To Number_Component
        ''''Write #f, C_Data_Points(j, i)
        Rs1.AddNew
        Rs1("FieldName") = "C_Data_Points"
        Rs1("FieldIndex") = J
        Rs1("FieldIndex2") = i
        Rs1("dblValue") = C_Data_Points(J, i)
        Rs1.Update
      Next J
    Next i
  End If
  'END SAVE TO THIS TABLE.
  Rs1.Close
  
  '------ OUTPUT DATA TO TABLE "PSDMInRoomData". ------------------------------------------------------------------------------------------------------
  'START SAVE TO THIS TABLE.
  Call Database_DeleteTableContents(Db1, "PSDMInRoomData")
  Set Rs1 = Db1.OpenRecordset("PSDMInRoomData")
  ' INPUT ROOM PARAMETERS.
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
  Rs1.Close
  
  '---- NEW AS OF 11/11/99 BEGINS: ---------------------------------------------------------
  'START SAVE TO THIS TABLE.
  Call Database_DeleteTableContents(Db1, "PSDMInRoomData_CO_Data")
  Set Rs1 = Db1.OpenRecordset("PSDMInRoomData_CO_Data")
  'MAIN DATA SET.
  For J = 1 To Number_Component
    For i = 1 To rp.int_ROOM_NCOINI(J)
      Rs1.AddNew
      Rs1("FieldName") = "dbl_ROOM_TCOINI"
      Rs1("FieldIndex") = J
      Rs1("FieldIndex2") = i
      Rs1("dblValue") = rp.dbl_ROOM_TCOINI(J, i)
      Rs1.Update
      Rs1.AddNew
      Rs1("FieldName") = "dbl_ROOM_COINI"
      Rs1("FieldIndex") = J
      Rs1("FieldIndex2") = i
      Rs1("dblValue") = rp.dbl_ROOM_COINI(J, i)
      Rs1.Update
    Next i
  Next J
  'END SAVE TO THIS TABLE.
  Rs1.Close
  '---- NEW AS OF 11/11/99 ENDS. ---------------------------------------------------------
  
  '---- NEW AS OF 11/11/99 BEGINS: ---------------------------------------------------------
  'START SAVE TO THIS TABLE.
  Call Database_DeleteTableContents(Db1, "PSDMInRoomData_WA_Data")
  Set Rs1 = Db1.OpenRecordset("PSDMInRoomData_WA_Data")
  'MAIN DATA SET.
  For J = 1 To Number_Component
    For i = 1 To rp.int_ROOM_NEMITINI(J)
      Rs1.AddNew
      Rs1("FieldName") = "dbl_ROOM_TEMITINI"
      Rs1("FieldIndex") = J
      Rs1("FieldIndex2") = i
      Rs1("dblValue") = rp.dbl_ROOM_TEMITINI(J, i)
      Rs1.Update
      Rs1.AddNew
      Rs1("FieldName") = "dbl_ROOM_EMITINI"
      Rs1("FieldIndex") = J
      Rs1("FieldIndex2") = i
      Rs1("dblValue") = rp.dbl_ROOM_EMITINI(J, i)
      Rs1.Update
    Next i
  Next J
  'END SAVE TO THIS TABLE.
  Rs1.Close
  '---- NEW AS OF 11/11/99 ENDS. ---------------------------------------------------------
  
  '---- NEW AS OF 1/17/00 BEGINS: ---------------------------------------------------------
  'START SAVE TO THIS TABLE.
  Call Database_DeleteTableContents(Db1, "PSDMInRoomData_K_Data")
  Set Rs1 = Db1.OpenRecordset("PSDMInRoomData_K_Data")
  'MAIN DATA SET.
  For J = 1 To Number_Component
    For i = 1 To rp.int_ROOM_NKINI(J)
      Rs1.AddNew
      Rs1("FieldName") = "dbl_ROOM_TKINI"
      Rs1("FieldIndex") = J
      Rs1("FieldIndex2") = i
      Rs1("dblValue") = rp.dbl_ROOM_TKINI(J, i)
      Rs1.Update
      Rs1.AddNew
      Rs1("FieldName") = "dbl_ROOM_KINI"
      Rs1("FieldIndex") = J
      Rs1("FieldIndex2") = i
      Rs1("dblValue") = rp.dbl_ROOM_KINI(J, i)
      Rs1.Update
    Next i
  Next J
  'END SAVE TO THIS TABLE.
  Rs1.Close
  '---- NEW AS OF 1/17/00 ENDS. ---------------------------------------------------------
  
  'CLOSE THE DATABASE FILE.
  Db1.Close

  'COMPACT THE DATABASE FILE.
      'TO DO: USE THE DbEngine.CompactDatabase METHOD
      'TO COMPACT THE DATABASE.  PROBLEM TO CONSIDER:
      'THE DB MUST BE COMPACTED TO A TEMPORARY FILE,
      'WHICH THEN SHOULD OVERWRITE THE ORIGINAL FILE.
  
  'RETURN A "SUCCESS" MESSAGE TO CALLER.
  File_Save_Latest_v1_60 = True
  
End Function


Sub Units1_Database_SaveProperty(Rs1 As Recordset, CboX As Control, Desc As String)
Dim OutStr As String
  If (CboX.ListIndex >= 0) Then
    OutStr = CboX.List(CboX.ListIndex)
  Else
    If (CboX.ListCount > 0) Then
      OutStr = CboX.List(0)
    Else
      OutStr = ""     'NOT LIKELY TO GET HERE!
    End If
  End If
  ''''Call ProjectFile_Write(f, OutStr, Desc)
  Call Database_SaveProperty(Rs1, Desc, OutStr)
End Sub
Sub Units1_Database_LoadProperty(Rs1 As Recordset, CboX As Control)
Dim TxtX As Control
Dim inline As String
Dim Dummy1 As String
Dim NewUnits As String
Dim H As Integer
  Call Database_LoadProperty(Rs1, inline)
  ''''Call ProjectFile_Read(f, InLine, Dummy1)
  NewUnits = inline
  H = unitsys_lookup_cbox(CboX)
  Set TxtX = unitsys(H).TxtX
  Call unitsys_set_units(TxtX, NewUnits)
End Sub


Sub Database_LoadProperty( _
    Rs1 As Recordset, _
    LoadedData As Variant, _
    Optional Use_memoValue As Boolean = False)
  Select Case VarType(LoadedData)
    Case vbBoolean:
      LoadedData = CBool(Database_Get_Long(Rs1, "lngValue"))
    Case vbByte:
      LoadedData = CByte(Database_Get_Long(Rs1, "lngValue"))
    Case vbInteger:
      LoadedData = CInt(Database_Get_Long(Rs1, "lngValue"))
    Case vbLong:
      LoadedData = CLng(Database_Get_Long(Rs1, "lngValue"))
    Case vbString, vbDate:
      If (Use_memoValue) Then
        LoadedData = CStr(Database_Get_String(Rs1, "memoValue"))
      Else
        LoadedData = CStr(Database_Get_String(Rs1, "strValue"))
      End If
    Case vbDouble:
      LoadedData = CDbl(Database_Get_Double(Rs1, "dblValue"))
    Case vbSingle:
      LoadedData = CSng(Database_Get_Double(Rs1, "dblValue"))
  End Select
End Sub
Sub Database_SaveProperty( _
    Rs1 As Recordset, _
    in_Use_FieldName As String, _
    SavedData As Variant)
Dim Use_memoValue As Boolean
Dim Use_FieldName As String
  'NOTE: IF THE FIRST CHARACTER IS AN ASTERISK ("*"), THEN
  'THE FIELD TYPE USED IS THE MEMO TYPE (memoValue).
  Use_memoValue = False
  If (Left$(in_Use_FieldName, 1) = "*") Then
    Use_FieldName = Right$(in_Use_FieldName, Len(in_Use_FieldName) - 1)
    Use_memoValue = True
  Else
    Use_FieldName = in_Use_FieldName
  End If
  Rs1.AddNew
  Rs1("FieldName") = Use_FieldName
  Select Case VarType(SavedData)
    Case vbBoolean, vbByte, vbInteger, vbLong:
      Rs1("lngValue") = CLng(SavedData)
    Case vbString, vbDate:
      If (Use_memoValue) Then
        Rs1("memoValue") = CStr(SavedData)
      Else
        Rs1("strValue") = CStr(SavedData)
      End If
    Case vbDouble, vbSingle:
      Rs1("dblValue") = CDbl(SavedData)
  End Select
  Rs1.Update
End Sub
Sub Database_SavePropertyWithIndex( _
    Rs1 As Recordset, _
    in_Use_FieldName As String, _
    Use_FieldIndex As Integer, _
    SavedData As Variant)
Dim Use_memoValue As Boolean
Dim Use_FieldName As String
  'NOTE: IF THE FIRST CHARACTER IS AN ASTERISK ("*"), THEN
  'THE FIELD TYPE USED IS THE MEMO TYPE (memoValue).
  Use_memoValue = False
  If (Left$(in_Use_FieldName, 1) = "*") Then
    Use_FieldName = Right$(in_Use_FieldName, Len(in_Use_FieldName) - 1)
    Use_memoValue = True
  Else
    Use_FieldName = in_Use_FieldName
  End If
  Rs1.AddNew
  Rs1("FieldName") = Use_FieldName
  Rs1("FieldIndex") = Use_FieldIndex
  Select Case VarType(SavedData)
    Case vbBoolean, vbByte, vbInteger, vbLong:
      Rs1("lngValue") = CLng(SavedData)
    Case vbString, vbDate:
      If (Use_memoValue) Then
        Rs1("memoValue") = CStr(SavedData)
      Else
        Rs1("strValue") = CStr(SavedData)
      End If
    Case vbDouble, vbSingle:
      Rs1("dblValue") = CDbl(SavedData)
  End Select
  Rs1.Update
End Sub


Sub Database_DeleteTableContents( _
    Db1 As Database, _
    TableName As String)
Dim Rs1 As Recordset
  On Error GoTo err_Database_DeleteTableContents
  Set Rs1 = Db1.OpenRecordset(TableName)
  Rs1.MoveFirst
  Do Until Rs1.EOF
    Rs1.Delete
    Rs1.MoveNext
  Loop
  Rs1.Close
  Exit Sub
exit_err_Database_DeleteTableContents:
  Exit Sub
err_Database_DeleteTableContents:
  Resume exit_err_Database_DeleteTableContents
End Sub
Sub Database_CreateMFBTable( _
    Db1 As Database, _
    TableName As String, _
    Include_FieldIndex As Boolean, _
    Include_FieldIndex2 As Boolean)
Dim Td1 As TableDef
Dim Ff As Field
    
  Set Td1 = Db1.CreateTableDef(TableName)
  Set Ff = Td1.CreateField("RecordID", dbLong):
  'TODO: ADD AUTONUMBER SETUP FOR THIS FIELD (NOT TOO NEEDED).
  Td1.Fields.Append Ff
  Set Ff = Td1.CreateField("FieldName", dbText, 250):
  Ff.AllowZeroLength = True
  Td1.Fields.Append Ff
  If (Include_FieldIndex) Then
    Set Ff = Td1.CreateField("FieldIndex", dbLong):
    Td1.Fields.Append Ff
  End If
  If (Include_FieldIndex2) Then
    Set Ff = Td1.CreateField("FieldIndex2", dbLong):
    Td1.Fields.Append Ff
  End If
  Set Ff = Td1.CreateField("strValue", dbText, 250):
  Ff.AllowZeroLength = True
  Td1.Fields.Append Ff
  Set Ff = Td1.CreateField("dblValue", dbDouble):
  Td1.Fields.Append Ff
  Set Ff = Td1.CreateField("lngValue", dbLong):
  Td1.Fields.Append Ff
  Set Ff = Td1.CreateField("memoValue", dbMemo):
  Ff.AllowZeroLength = True
  Td1.Fields.Append Ff
  Set Ff = Td1.CreateField("Comments", dbText, 250):
  Ff.AllowZeroLength = True
  Td1.Fields.Append Ff
  Db1.TableDefs.Append Td1
End Sub
Sub Database_CreateMFBTable_IfNoExist( _
    Db1 As Database, _
    Use_TableName As String, _
    Include_FieldIndex As Boolean)
  If (Database_IsTableExist(Db1, Use_TableName) = False) Then
    Call Database_CreateMFBTable(Db1, Use_TableName, Include_FieldIndex, False)
  End If
End Sub
Sub Database_CreateMFBTable_IfNoExist_TwoIndices( _
    Db1 As Database, _
    Use_TableName As String)
  If (Database_IsTableExist(Db1, Use_TableName) = False) Then
    Call Database_CreateMFBTable(Db1, Use_TableName, True, True)
  End If
End Sub




