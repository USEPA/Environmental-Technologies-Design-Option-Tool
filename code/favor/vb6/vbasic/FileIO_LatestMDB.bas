Attribute VB_Name = "FileIO_LatestMDB"
Option Explicit





Const FileIO_LatestMDB_declarations_end = True


'RETURNS:
'         TRUE = SUCCEEDED IN LOADING.
'         FALSE = FAILED IN LOADING.
Function File_Open_Latest_v1_00( _
    fn_This As String) As Boolean
Dim Ws1 As Workspace
Dim Db1 As Database
Dim Rs1 As Recordset
Dim Use_FieldIndex As Integer
Dim Use_FieldIndex2 As Integer
Dim ContainsTable_PSDMInRoomData As Boolean
Dim Prj As Project_Type
Dim Pp As TYPE_PlantDiagram
Dim Iw As TYPE_Weir
Dim Gc As TYPE_GritChamber
Dim Pc As TYPE_Clarifier
Dim Pw As TYPE_Weir
Dim Ab As TYPE_AerationBasin
Dim Sc As TYPE_Clarifier
Dim Sw As TYPE_Weir
Dim Cd As TYPE_PhysicoChemicalData
Dim Cs As TYPE_CSTRModeling
Dim Bt As TYPE_BioTreatmentModeling
Dim Ds As TYPE_DataSource

'>>>>>>>>>>>>>>>>>>>> *TODO* UPDATE THIS ENTIRE FUNCTION <<<<<<<<<<<<<<<<<<<<<<<<<

  '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
  '>>>>>>>>>>>>>>>>>>>>>>>>>>>  INPUT FROM MAIN DATABASE  <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
  '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
  If (Not FileExists(fn_This)) Then
    'ERROR: UNABLE TO FIND THE FILE!
    File_Open_Latest_v1_00 = False
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
  If (Database_IsTableExist(Db1, "Main") = False) Then
    'DO NOTHING: USE DEFAULT VALUES.
  Else
    Set Rs1 = Db1.OpenRecordset("Main")
    If (Database_NoRecordsInRecordset(Rs1)) Then
      'DO NOTHING: USE DEFAULT VALUES.
    Else
      '
      ' SET DEFAULT PROJECT DATA IN TEMPORARY VARIABLE.
      '
      Call Project_SetDefaults(Prj)
      '
      ' READ IN THE PROJECT DATA TO TEMPORARY VARIABLE.
      '
      Rs1.MoveFirst
      Do Until Rs1.EOF
        Use_FieldIndex = CInt(Database_Get_Long(Rs1, "FieldIndex"))
        Select Case Trim$(UCase$(Database_Get_String(Rs1, "FieldName")))
          '
          ' HANDLE Pp = Prj.Plant = NowProj.Plant.
          '
          Case Trim$(UCase$("Pp.en_InfluentWeir")): Call Database_LoadProperty(Rs1, Pp.en_InfluentWeir)
          Case Trim$(UCase$("Pp.en_GritChamber")): Call Database_LoadProperty(Rs1, Pp.en_GritChamber)
          Case Trim$(UCase$("Pp.en_PrimaryWeir")): Call Database_LoadProperty(Rs1, Pp.en_PrimaryWeir)
          Case Trim$(UCase$("Pp.en_SecondaryWeir")): Call Database_LoadProperty(Rs1, Pp.en_SecondaryWeir)
          Case Trim$(UCase$("Pp.Flow")): Call Database_LoadProperty(Rs1, Pp.Flow)
          Case Trim$(UCase$("Pp.SolidsConc")): Call Database_LoadProperty(Rs1, Pp.SolidsConc)
          '
          ' HANDLE Iw = Pp.InfluentWeir.
          '
          Case Trim$(UCase$("Iw.ModelingMechanism")): Call Database_LoadProperty(Rs1, Iw.ModelingMechanism)
          Case Trim$(UCase$("Iw.Width")): Call Database_LoadProperty(Rs1, Iw.Width)
          Case Trim$(UCase$("Iw.WaterLevelDiff")): Call Database_LoadProperty(Rs1, Iw.WaterLevelDiff)
          Case Trim$(UCase$("Iw.GasFlow")): Call Database_LoadProperty(Rs1, Iw.GasFlow)
          Case Trim$(UCase$("Iw.UnitsOfDisplay")): Call Database_LoadProperty(Rs1, Iw.UnitsOfDisplay(Use_FieldIndex))
          '
          ' HANDLE Gc = Pp.GritChamber.
          '
          Case Trim$(UCase$("Gc.IsCovered")): Call Database_LoadProperty(Rs1, Gc.IsCovered)
          Case Trim$(UCase$("Gc.Count")): Call Database_LoadProperty(Rs1, Gc.Count)
          Case Trim$(UCase$("Gc.VentilationRate")): Call Database_LoadProperty(Rs1, Gc.VentilationRate)
          Case Trim$(UCase$("Gc.Depth")): Call Database_LoadProperty(Rs1, Gc.Depth)
          Case Trim$(UCase$("Gc.Volume")): Call Database_LoadProperty(Rs1, Gc.Volume)
          Case Trim$(UCase$("Gc.GasFlow")): Call Database_LoadProperty(Rs1, Gc.GasFlow)
          Case Trim$(UCase$("Gc.SOTR")): Call Database_LoadProperty(Rs1, Gc.SOTR)
          Case Trim$(UCase$("Gc.UnitsOfDisplay")): Call Database_LoadProperty(Rs1, Gc.UnitsOfDisplay(Use_FieldIndex))
          '
          ' HANDLE Pc = Pp.PrimaryClarifier.
          '
          Case Trim$(UCase$("Pc.IsCovered")): Call Database_LoadProperty(Rs1, Pc.IsCovered)
          Case Trim$(UCase$("Pc.Count")): Call Database_LoadProperty(Rs1, Pc.Count)
          Case Trim$(UCase$("Pc.SorptionRemovalMethod")): Call Database_LoadProperty(Rs1, Pc.SorptionRemovalMethod)
          Case Trim$(UCase$("Pc.VolatilizationRemovalMechanism")): Call Database_LoadProperty(Rs1, Pc.VolatilizationRemovalMechanism)
          Case Trim$(UCase$("Pc.VentilationRate")): Call Database_LoadProperty(Rs1, Pc.VentilationRate)
          Case Trim$(UCase$("Pc.Depth")): Call Database_LoadProperty(Rs1, Pc.Depth)
          Case Trim$(UCase$("Pc.Volume")): Call Database_LoadProperty(Rs1, Pc.Volume)
          Case Trim$(UCase$("Pc.WastageFlow")): Call Database_LoadProperty(Rs1, Pc.WastageFlow)
          Case Trim$(UCase$("Pc.PercentageRemoval")): Call Database_LoadProperty(Rs1, Pc.PercentageRemoval)
          Case Trim$(UCase$("Pc.EffluentSolidsConc")): Call Database_LoadProperty(Rs1, Pc.EffluentSolidsConc)
          Case Trim$(UCase$("Pc.UnitsOfDisplay")): Call Database_LoadProperty(Rs1, Pc.UnitsOfDisplay(Use_FieldIndex))
          '
          ' HANDLE Pw = Pp.PrimaryWeir.
          '
          Case Trim$(UCase$("Pw.ModelingMechanism")): Call Database_LoadProperty(Rs1, Pw.ModelingMechanism)
          Case Trim$(UCase$("Pw.Width")): Call Database_LoadProperty(Rs1, Pw.Width)
          Case Trim$(UCase$("Pw.WaterLevelDiff")): Call Database_LoadProperty(Rs1, Pw.WaterLevelDiff)
          Case Trim$(UCase$("Pw.GasFlow")): Call Database_LoadProperty(Rs1, Pw.GasFlow)
          Case Trim$(UCase$("Pw.UnitsOfDisplay")): Call Database_LoadProperty(Rs1, Pw.UnitsOfDisplay(Use_FieldIndex))
          '
          ' HANDLE Ab = Pp.AerationBasin.
          '
          Case Trim$(UCase$("Ab.IsCovered")): Call Database_LoadProperty(Rs1, Ab.IsCovered)
          Case Trim$(UCase$("Ab.Count")): Call Database_LoadProperty(Rs1, Ab.Count)
          Case Trim$(UCase$("Ab.ModelingMechanism")): Call Database_LoadProperty(Rs1, Ab.ModelingMechanism)
          Case Trim$(UCase$("Ab.AutoCalcBioMass")): Call Database_LoadProperty(Rs1, Ab.AutoCalcBioMass)
          Case Trim$(UCase$("Ab.VentilationRate")): Call Database_LoadProperty(Rs1, Ab.VentilationRate)
          Case Trim$(UCase$("Ab.Depth")): Call Database_LoadProperty(Rs1, Ab.Depth)
          Case Trim$(UCase$("Ab.WastageFlow")): Call Database_LoadProperty(Rs1, Ab.WastageFlow)
          Case Trim$(UCase$("Ab.RecycleFlow")): Call Database_LoadProperty(Rs1, Ab.RecycleFlow)
          Case Trim$(UCase$("Ab.SolidsConcInRecycle")): Call Database_LoadProperty(Rs1, Ab.SolidsConcInRecycle)
          Case Trim$(UCase$("Ab.SOTR")): Call Database_LoadProperty(Rs1, Ab.SOTR)
          Case Trim$(UCase$("Ab.Volume")): Call Database_LoadProperty(Rs1, Ab.Volume)
          Case Trim$(UCase$("Ab.GasFlow")): Call Database_LoadProperty(Rs1, Ab.GasFlow)
          Case Trim$(UCase$("Ab.BioMass")): Call Database_LoadProperty(Rs1, Ab.BioMass)
          Case Trim$(UCase$("Ab.UnitsOfDisplay")): Call Database_LoadProperty(Rs1, Ab.UnitsOfDisplay(Use_FieldIndex))
          Case Trim$(UCase$("Cs.Count")): Call Database_LoadProperty(Rs1, Cs.Count)
          Case Trim$(UCase$("Cs.UseStepFeed")): Call Database_LoadProperty(Rs1, Cs.UseStepFeed)
          Case Trim$(UCase$("Cs.UniformFeed")): Call Database_LoadProperty(Rs1, Cs.UniformFeed)
          Case Trim$(UCase$("Cs.Feed")): Call Database_LoadProperty(Rs1, Cs.Feed(Use_FieldIndex))
          Case Trim$(UCase$("Cs.UniformVolume")): Call Database_LoadProperty(Rs1, Cs.UniformVolume)
          Case Trim$(UCase$("Cs.Volume")): Call Database_LoadProperty(Rs1, Cs.Volume(Use_FieldIndex))
          Case Trim$(UCase$("Cs.UniformGasFlow")): Call Database_LoadProperty(Rs1, Cs.UniformGasFlow)
          Case Trim$(UCase$("Cs.GasFlow")): Call Database_LoadProperty(Rs1, Cs.GasFlow(Use_FieldIndex))
          Case Trim$(UCase$("Cs.UniformBioMass")): Call Database_LoadProperty(Rs1, Cs.UniformBioMass)
          Case Trim$(UCase$("Cs.BioMass")): Call Database_LoadProperty(Rs1, Cs.BioMass(Use_FieldIndex))
          Case Trim$(UCase$("Bt.MaxGrowthRate")): Call Database_LoadProperty(Rs1, Bt.MaxGrowthRate)
          Case Trim$(UCase$("Bt.HalfVelocityConst")): Call Database_LoadProperty(Rs1, Bt.HalfVelocityConst)
          Case Trim$(UCase$("Bt.BacterialDecay")): Call Database_LoadProperty(Rs1, Bt.BacterialDecay)
          Case Trim$(UCase$("Bt.YieldCoeff")): Call Database_LoadProperty(Rs1, Bt.YieldCoeff)
          Case Trim$(UCase$("Bt.BOD5Conc")): Call Database_LoadProperty(Rs1, Bt.BOD5Conc)
          Case Trim$(UCase$("Bt.UnitsOfDisplay")): Call Database_LoadProperty(Rs1, Bt.UnitsOfDisplay(Use_FieldIndex))
          '
          ' HANDLE Sc = Pp.SecondaryClarifier.
          '
          Case Trim$(UCase$("Sc.IsCovered")): Call Database_LoadProperty(Rs1, Sc.IsCovered)
          Case Trim$(UCase$("Sc.Count")): Call Database_LoadProperty(Rs1, Sc.Count)
          Case Trim$(UCase$("Sc.SorptionRemovalMethod")): Call Database_LoadProperty(Rs1, Sc.SorptionRemovalMethod)
          Case Trim$(UCase$("Sc.VolatilizationRemovalMechanism")): Call Database_LoadProperty(Rs1, Sc.VolatilizationRemovalMechanism)
          Case Trim$(UCase$("Sc.VentilationRate")): Call Database_LoadProperty(Rs1, Sc.VentilationRate)
          Case Trim$(UCase$("Sc.Depth")): Call Database_LoadProperty(Rs1, Sc.Depth)
          Case Trim$(UCase$("Sc.Volume")): Call Database_LoadProperty(Rs1, Sc.Volume)
          Case Trim$(UCase$("Sc.WastageFlow")): Call Database_LoadProperty(Rs1, Sc.WastageFlow)
          Case Trim$(UCase$("Sc.PercentageRemoval")): Call Database_LoadProperty(Rs1, Sc.PercentageRemoval)
          Case Trim$(UCase$("Sc.EffluentSolidsConc")): Call Database_LoadProperty(Rs1, Sc.EffluentSolidsConc)
          Case Trim$(UCase$("Sc.UnitsOfDisplay")): Call Database_LoadProperty(Rs1, Sc.UnitsOfDisplay(Use_FieldIndex))
          '
          ' HANDLE Sw = Pp.SecondaryWeir.
          '
          Case Trim$(UCase$("Sw.ModelingMechanism")): Call Database_LoadProperty(Rs1, Sw.ModelingMechanism)
          Case Trim$(UCase$("Sw.Width")): Call Database_LoadProperty(Rs1, Sw.Width)
          Case Trim$(UCase$("Sw.WaterLevelDiff")): Call Database_LoadProperty(Rs1, Sw.WaterLevelDiff)
          Case Trim$(UCase$("Sw.GasFlow")): Call Database_LoadProperty(Rs1, Sw.GasFlow)
          Case Trim$(UCase$("Sw.UnitsOfDisplay")): Call Database_LoadProperty(Rs1, Sw.UnitsOfDisplay(Use_FieldIndex))
          '
          ' HANDLE Cd = Pp.ChemicalData.
          '
          Case Trim$(UCase$("Cd.env_Pressure")): Call Database_LoadProperty(Rs1, Cd.env_Pressure)
          Case Trim$(UCase$("Cd.env_Temperature")): Call Database_LoadProperty(Rs1, Cd.env_Temperature)
          Case Trim$(UCase$("Cd.env_WindVelocity")): Call Database_LoadProperty(Rs1, Cd.env_WindVelocity)
          Case Trim$(UCase$("Cd.ContaminantName")): Call Database_LoadProperty(Rs1, Cd.ContaminantName)
          Case Trim$(UCase$("Cd.InfluentConc")): Call Database_LoadProperty(Rs1, Cd.InfluentConc)
          Case Trim$(UCase$("Cd.BiodegredationRate")): Call Database_LoadProperty(Rs1, Cd.BiodegredationRate)
          Case Trim$(UCase$("Cd.LogKow")): Call Database_LoadProperty(Rs1, Cd.LogKow)
          Case Trim$(UCase$("Cd.VOC_HenrysConstant")): Call Database_LoadProperty(Rs1, Cd.VOC_HenrysConstant)
          Case Trim$(UCase$("Cd.VOC_MolecularWeight")): Call Database_LoadProperty(Rs1, Cd.VOC_MolecularWeight)
          Case Trim$(UCase$("Cd.VOC_DiffusivityInH2O")): Call Database_LoadProperty(Rs1, Cd.VOC_DiffusivityInH2O)
          Case Trim$(UCase$("Cd.VOC_DiffusivityInGas")): Call Database_LoadProperty(Rs1, Cd.VOC_DiffusivityInGas)
          Case Trim$(UCase$("Cd.O2_SaturationConc")): Call Database_LoadProperty(Rs1, Cd.O2_SaturationConc)
          Case Trim$(UCase$("Cd.O2_HenrysConstant")): Call Database_LoadProperty(Rs1, Cd.O2_HenrysConstant)
          Case Trim$(UCase$("Cd.O2_Diffusivity")): Call Database_LoadProperty(Rs1, Cd.O2_Diffusivity)
          Case Trim$(UCase$("Cd.H2O_Density")): Call Database_LoadProperty(Rs1, Cd.H2O_Density)
          Case Trim$(UCase$("Cd.H2O_Viscosity")): Call Database_LoadProperty(Rs1, Cd.H2O_Viscosity)
          Case Trim$(UCase$("Cd.H2O_VaporPressure")): Call Database_LoadProperty(Rs1, Cd.H2O_VaporPressure)
          Case Trim$(UCase$("Cd.H2O_Alpha")): Call Database_LoadProperty(Rs1, Cd.H2O_Alpha)
          Case Trim$(UCase$("Cd.AIR_Density")): Call Database_LoadProperty(Rs1, Cd.AIR_Density)
          Case Trim$(UCase$("Cd.AIR_Viscosity")): Call Database_LoadProperty(Rs1, Cd.AIR_Viscosity)
          Case Trim$(UCase$("Ds.SourceType")): Call Database_LoadProperty(Rs1, Cd.DataSources(Use_FieldIndex).SourceType)
          Case Trim$(UCase$("Ds.Val_UserInput")): Call Database_LoadProperty(Rs1, Cd.DataSources(Use_FieldIndex).Val_UserInput)
          Case Trim$(UCase$("Ds.Val_StEPP")): Call Database_LoadProperty(Rs1, Cd.DataSources(Use_FieldIndex).Val_StEPP)
          Case Trim$(UCase$("Ds.Val_Corr")): Call Database_LoadProperty(Rs1, Cd.DataSources(Use_FieldIndex).Val_Corr)
          Case Trim$(UCase$("Cd.O2_CInfinity")): Call Database_LoadProperty(Rs1, Cd.O2_CInfinity)
          Case Trim$(UCase$("Cd.UnitsOfDisplay")): Call Database_LoadProperty(Rs1, Cd.UnitsOfDisplay(Use_FieldIndex))
          '
          ' MISCELLANEOUS.
          '
          Case Trim$(UCase$("UnitType")): Call Database_LoadProperty(Rs1, Prj.UnitType)
        End Select
        Rs1.MoveNext
      Loop
      '
      ' TRANSFER PROJECT DATA TO MEMORY.
      '
      Pp.InfluentWeir = Iw
      Pp.GritChamber = Gc
      Pp.PrimaryClarifier = Pc
      Pp.PrimaryWeir = Pw
      Ab.CSTR = Cs
      Ab.BioTreat = Bt
      Pp.AerationBasin = Ab
      Pp.SecondaryClarifier = Sc
      Pp.SecondaryWeir = Sw
      Pp.ChemicalData = Cd
      Prj.Plant = Pp
      NowProj = Prj
    End If
    Rs1.Close
  End If

  'CLOSE THE DATABASE FILE.
  Db1.Close
  'add call to calculate
  Call ModelFAVOR_Go
  'RETURN A "SUCCESS" MESSAGE TO CALLER.
  File_Open_Latest_v1_00 = True

End Function


'RETURNS:
'         TRUE = SUCCEEDED IN SAVING.
'         FALSE = FAILED IN SAVING.
Function File_Save_Latest_v1_00( _
    fn_This As String) As Boolean
Dim Ws1 As Workspace
Dim Db1 As Database
Dim Rs1 As Recordset
Dim i As Integer
Dim j As Integer
Dim IsInvalidFormat As Boolean
Dim NeedToCreateNewDatabase As Boolean
Dim Prj As Project_Type
Dim Pp As TYPE_PlantDiagram
Dim Iw As TYPE_Weir
Dim Gc As TYPE_GritChamber
Dim Pc As TYPE_Clarifier
Dim Pw As TYPE_Weir
Dim Ab As TYPE_AerationBasin
Dim Sc As TYPE_Clarifier
Dim Sw As TYPE_Weir
Dim Cd As TYPE_PhysicoChemicalData
Dim Cs As TYPE_CSTRModeling
Dim Bt As TYPE_BioTreatmentModeling
Dim Ds As TYPE_DataSource

'>>>>>>>>>>>>>>>>>>>> *TODO* UPDATE THIS ENTIRE FUNCTION <<<<<<<<<<<<<<<<<<<<<<<<<

  '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
  '>>>>>>>>>>>>>>>>>>>>>>>>>>>  SAVE TO MAIN DATABASE  <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
  '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
  
  'IF FILE DOES NOT EXIST, CREATE IT.
  'FOR EACH TABLE, IF IT EXISTS, DELETE IT.
  If (Not FileExists(fn_This)) Then
    'CREATE NEW DATABASE.
    NeedToCreateNewDatabase = True
  Else
    'DETERMINE WHETHER OLD FILE IS AN INVALID VERSION (i.e. A NON-MDB FILE).
    IsInvalidFormat = True
    On Error Resume Next
    Set Db1 = OpenDatabase(fn_This)
    If (Err.Number = 0) Then
      IsInvalidFormat = False
      Db1.Close
    End If
    On Error GoTo 0
    If (IsInvalidFormat) Then
      'DELETE OLD FILE, CREATE NEW DATABASE (SEE BELOW).
      Kill fn_This
      NeedToCreateNewDatabase = True
    Else
      'OPEN DATABASE NORMALLY.
      Set Db1 = OpenDatabase(fn_This)
    End If
  End If
  If (NeedToCreateNewDatabase) Then
    FileCopy MAIN_APP_PATH & "\dbase\template.fvr", fn_This
    Set Db1 = OpenDatabase(fn_This)
    ''''Set Db1 = CreateDatabase(fn_This, dbLangGeneral)
  End If
  'CREATE NEW TABLES WITHIN DATABASE, IF NECESSARY.
  Call Database_CreateMFBTable_IfNoExist(Db1, "Version", True)
  Call Database_CreateMFBTable_IfNoExist(Db1, "Main", True)
  
  '=========== OUTPUT DATA TO DATABASE TABLES. =================
  
  '------ OUTPUT DATA TO TABLE "Version". ------------------------------------------------------------------------------------------------------
  'START SAVE TO THIS TABLE.
  Call Database_DeleteTableContents(Db1, "Version")
  Set Rs1 = Db1.OpenRecordset("Version")
  Call Database_SaveProperty(Rs1, "DataVersion_Major", CInt(Latest_DataVersion_Major))
  Call Database_SaveProperty(Rs1, "DataVersion_Minor", CInt(Latest_DataVersion_Minor))
  ''''Call Database_SaveProperty(Rs1, "ContainsTable_PSDMInRoomData", True)
  'END SAVE TO THIS TABLE.
  Rs1.Close
  
  '------ OUTPUT DATA TO TABLE "Main". ------------------------------------------------------------------------------------------------------
  'START SAVE TO THIS TABLE.
  Call Database_DeleteTableContents(Db1, "Main")
  Set Rs1 = Db1.OpenRecordset("Main")
  'MAIN BLOCK.
  Prj = NowProj
  Pp = Prj.Plant
  Iw = Pp.InfluentWeir
  Gc = Pp.GritChamber
  Pc = Pp.PrimaryClarifier
  Pw = Pp.PrimaryWeir
  Ab = Pp.AerationBasin
  Sc = Pp.SecondaryClarifier
  Sw = Pp.SecondaryWeir
  Cd = Pp.ChemicalData
  Cs = Ab.CSTR
  Bt = Ab.BioTreat
  '
  ' HANDLE Pp = Prj.Plant = NowProj.Plant.
  '
  Call Database_SaveProperty(Rs1, "Pp.en_InfluentWeir", Pp.en_InfluentWeir)
  Call Database_SaveProperty(Rs1, "Pp.en_GritChamber", Pp.en_GritChamber)
  Call Database_SaveProperty(Rs1, "Pp.en_PrimaryWeir", Pp.en_PrimaryWeir)
  Call Database_SaveProperty(Rs1, "Pp.en_SecondaryWeir", Pp.en_SecondaryWeir)
  Call Database_SaveProperty(Rs1, "Pp.Flow", Pp.Flow)
  Call Database_SaveProperty(Rs1, "Pp.SolidsConc", Pp.SolidsConc)
  '
  ' HANDLE Iw = Pp.InfluentWeir.
  '
  Call Database_SaveProperty(Rs1, "Iw.ModelingMechanism", Iw.ModelingMechanism)
  Call Database_SaveProperty(Rs1, "Iw.Width", Iw.Width)
  Call Database_SaveProperty(Rs1, "Iw.WaterLevelDiff", Iw.WaterLevelDiff)
  Call Database_SaveProperty(Rs1, "Iw.GasFlow", Iw.GasFlow)
  For i = 0 To 2
    Call Database_SavePropertyWithIndex(Rs1, "Iw.UnitsOfDisplay", i, Iw.UnitsOfDisplay(i))
  Next i
  '
  ' HANDLE Gc = Pp.GritChamber.
  '
  Call Database_SaveProperty(Rs1, "Gc.IsCovered", Gc.IsCovered)
  Call Database_SaveProperty(Rs1, "Gc.Count", Gc.Count)
  Call Database_SaveProperty(Rs1, "Gc.VentilationRate", Gc.VentilationRate)
  Call Database_SaveProperty(Rs1, "Gc.Depth", Gc.Depth)
  Call Database_SaveProperty(Rs1, "Gc.Volume", Gc.Volume)
  Call Database_SaveProperty(Rs1, "Gc.GasFlow", Gc.GasFlow)
  Call Database_SaveProperty(Rs1, "Gc.SOTR", Gc.SOTR)
  For i = 1 To 5
    Call Database_SavePropertyWithIndex(Rs1, "Gc.UnitsOfDisplay", i, Gc.UnitsOfDisplay(i))
  Next i
  '
  ' HANDLE Pc = Pp.PrimaryClarifier.
  '
  Call Database_SaveProperty(Rs1, "Pc.IsCovered", Pc.IsCovered)
  Call Database_SaveProperty(Rs1, "Pc.Count", Pc.Count)
  Call Database_SaveProperty(Rs1, "Pc.SorptionRemovalMethod", Pc.SorptionRemovalMethod)
  Call Database_SaveProperty(Rs1, "Pc.VolatilizationRemovalMechanism", Pc.VolatilizationRemovalMechanism)
  Call Database_SaveProperty(Rs1, "Pc.VentilationRate", Pc.VentilationRate)
  Call Database_SaveProperty(Rs1, "Pc.Depth", Pc.Depth)
  Call Database_SaveProperty(Rs1, "Pc.Volume", Pc.Volume)
  Call Database_SaveProperty(Rs1, "Pc.WastageFlow", Pc.WastageFlow)
  Call Database_SaveProperty(Rs1, "Pc.PercentageRemoval", Pc.PercentageRemoval)
  ''''Call Database_SaveProperty(Rs1, "Pc.EffluentSolidsConc", Pc.EffluentSolidsConc)
  For i = 1 To 5
    Call Database_SavePropertyWithIndex(Rs1, "Pc.UnitsOfDisplay", i, Pc.UnitsOfDisplay(i))
  Next i
  '
  ' HANDLE Pw = Pp.PrimaryWeir.
  '
  Call Database_SaveProperty(Rs1, "Pw.ModelingMechanism", Pw.ModelingMechanism)
  Call Database_SaveProperty(Rs1, "Pw.Width", Pw.Width)
  Call Database_SaveProperty(Rs1, "Pw.WaterLevelDiff", Pw.WaterLevelDiff)
  Call Database_SaveProperty(Rs1, "Pw.GasFlow", Pw.GasFlow)
  For i = 0 To 2
    Call Database_SavePropertyWithIndex(Rs1, "Pw.UnitsOfDisplay", i, Pw.UnitsOfDisplay(i))
  Next i
  '
  ' HANDLE Ab = Pp.AerationBasin.
  '
  Call Database_SaveProperty(Rs1, "Ab.IsCovered", Ab.IsCovered)
  Call Database_SaveProperty(Rs1, "Ab.Count", Ab.Count)
  Call Database_SaveProperty(Rs1, "Ab.ModelingMechanism", Ab.ModelingMechanism)
  Call Database_SaveProperty(Rs1, "Ab.AutoCalcBioMass", Ab.AutoCalcBioMass)
  Call Database_SaveProperty(Rs1, "Ab.VentilationRate", Ab.VentilationRate)
  Call Database_SaveProperty(Rs1, "Ab.Depth", Ab.Depth)
  Call Database_SaveProperty(Rs1, "Ab.WastageFlow", Ab.WastageFlow)
  Call Database_SaveProperty(Rs1, "Ab.RecycleFlow", Ab.RecycleFlow)
  Call Database_SaveProperty(Rs1, "Ab.SolidsConcInRecycle", Ab.SolidsConcInRecycle)
  Call Database_SaveProperty(Rs1, "Ab.SOTR", Ab.SOTR)
  Call Database_SaveProperty(Rs1, "Ab.Volume", Ab.Volume)
  Call Database_SaveProperty(Rs1, "Ab.GasFlow", Ab.GasFlow)
  Call Database_SaveProperty(Rs1, "Ab.BioMass", Ab.BioMass)
  For i = 1 To 5
    Call Database_SavePropertyWithIndex(Rs1, "Ab.UnitsOfDisplay", i, Ab.UnitsOfDisplay(i))
  Next i
  Call Database_SaveProperty(Rs1, "Cs.Count", Cs.Count)
  Call Database_SaveProperty(Rs1, "Cs.UseStepFeed", Cs.UseStepFeed)
  Call Database_SaveProperty(Rs1, "Cs.UniformFeed", Cs.UniformFeed)
  For i = 0 To 8
    Call Database_SavePropertyWithIndex(Rs1, "Cs.Feed", i, Cs.Feed(i))
  Next i
  Call Database_SaveProperty(Rs1, "Cs.UniformVolume", Cs.UniformVolume)
  For i = 0 To 8
    Call Database_SavePropertyWithIndex(Rs1, "Cs.Volume", i, Cs.Volume(i))
  Next i
  Call Database_SaveProperty(Rs1, "Cs.UniformGasFlow", Cs.UniformGasFlow)
  For i = 0 To 8
    Call Database_SavePropertyWithIndex(Rs1, "Cs.GasFlow", i, Cs.GasFlow(i))
  Next i
  Call Database_SaveProperty(Rs1, "Cs.UniformBioMass", Cs.UniformBioMass)
  For i = 0 To 8
    Call Database_SavePropertyWithIndex(Rs1, "Cs.BioMass", i, Cs.BioMass(i))
  Next i
  Call Database_SaveProperty(Rs1, "Bt.MaxGrowthRate", Bt.MaxGrowthRate)
  Call Database_SaveProperty(Rs1, "Bt.HalfVelocityConst", Bt.HalfVelocityConst)
  Call Database_SaveProperty(Rs1, "Bt.BacterialDecay", Bt.BacterialDecay)
  Call Database_SaveProperty(Rs1, "Bt.YieldCoeff", Bt.YieldCoeff)
  Call Database_SaveProperty(Rs1, "Bt.BOD5Conc", Bt.BOD5Conc)
  For i = 0 To 4
    Call Database_SavePropertyWithIndex(Rs1, "Bt.UnitsOfDisplay", i, Bt.UnitsOfDisplay(i))
  Next i
  '
  ' HANDLE Sc = Pp.SecondaryClarifier.
  '
  Call Database_SaveProperty(Rs1, "Sc.IsCovered", Sc.IsCovered)
  Call Database_SaveProperty(Rs1, "Sc.Count", Sc.Count)
  Call Database_SaveProperty(Rs1, "Sc.SorptionRemovalMethod", Sc.SorptionRemovalMethod)
  Call Database_SaveProperty(Rs1, "Sc.VolatilizationRemovalMechanism", Sc.VolatilizationRemovalMechanism)
  Call Database_SaveProperty(Rs1, "Sc.VentilationRate", Sc.VentilationRate)
  Call Database_SaveProperty(Rs1, "Sc.Depth", Sc.Depth)
  Call Database_SaveProperty(Rs1, "Sc.Volume", Sc.Volume)
  Call Database_SaveProperty(Rs1, "Sc.WastageFlow", Sc.WastageFlow)
  ''''Call Database_SaveProperty(Rs1, "Sc.PercentageRemoval", Sc.PercentageRemoval)
  Call Database_SaveProperty(Rs1, "Sc.EffluentSolidsConc", Sc.EffluentSolidsConc)
  For i = 1 To 5
    Call Database_SavePropertyWithIndex(Rs1, "Sc.UnitsOfDisplay", i, Sc.UnitsOfDisplay(i))
  Next i
  '
  ' HANDLE Sw = Pp.SecondaryWeir.
  '
  Call Database_SaveProperty(Rs1, "Sw.ModelingMechanism", Sw.ModelingMechanism)
  Call Database_SaveProperty(Rs1, "Sw.Width", Sw.Width)
  Call Database_SaveProperty(Rs1, "Sw.WaterLevelDiff", Sw.WaterLevelDiff)
  Call Database_SaveProperty(Rs1, "Sw.GasFlow", Sw.GasFlow)
  For i = 0 To 2
    Call Database_SavePropertyWithIndex(Rs1, "Sw.UnitsOfDisplay", i, Sw.UnitsOfDisplay(i))
  Next i
  '
  ' HANDLE Cd = Pp.ChemicalData.
  '
  Call Database_SaveProperty(Rs1, "Cd.env_Pressure", Cd.env_Pressure)
  Call Database_SaveProperty(Rs1, "Cd.env_Temperature", Cd.env_Temperature)
  Call Database_SaveProperty(Rs1, "Cd.env_WindVelocity", Cd.env_WindVelocity)
  Call Database_SaveProperty(Rs1, "Cd.ContaminantName", Cd.ContaminantName)
  Call Database_SaveProperty(Rs1, "Cd.InfluentConc", Cd.InfluentConc)
  Call Database_SaveProperty(Rs1, "Cd.BiodegredationRate", Cd.BiodegredationRate)
  Call Database_SaveProperty(Rs1, "Cd.LogKow", Cd.LogKow)
  Call Database_SaveProperty(Rs1, "Cd.VOC_HenrysConstant", Cd.VOC_HenrysConstant)
  Call Database_SaveProperty(Rs1, "Cd.VOC_MolecularWeight", Cd.VOC_MolecularWeight)
  Call Database_SaveProperty(Rs1, "Cd.VOC_DiffusivityInH2O", Cd.VOC_DiffusivityInH2O)
  Call Database_SaveProperty(Rs1, "Cd.VOC_DiffusivityInGas", Cd.VOC_DiffusivityInGas)
  Call Database_SaveProperty(Rs1, "Cd.O2_SaturationConc", Cd.O2_SaturationConc)
  Call Database_SaveProperty(Rs1, "Cd.O2_HenrysConstant", Cd.O2_HenrysConstant)
  Call Database_SaveProperty(Rs1, "Cd.O2_Diffusivity", Cd.O2_Diffusivity)
  Call Database_SaveProperty(Rs1, "Cd.H2O_Density", Cd.H2O_Density)
  Call Database_SaveProperty(Rs1, "Cd.H2O_Viscosity", Cd.H2O_Viscosity)
  Call Database_SaveProperty(Rs1, "Cd.H2O_VaporPressure", Cd.H2O_VaporPressure)
  Call Database_SaveProperty(Rs1, "Cd.H2O_Alpha", Cd.H2O_Alpha)
  Call Database_SaveProperty(Rs1, "Cd.AIR_Density", Cd.AIR_Density)
  Call Database_SaveProperty(Rs1, "Cd.AIR_Viscosity", Cd.AIR_Viscosity)
  For i = 0 To 18
    Ds = Cd.DataSources(i)
    Call Database_SavePropertyWithIndex(Rs1, "Ds.SourceType", i, Ds.SourceType)
    Call Database_SavePropertyWithIndex(Rs1, "Ds.Val_UserInput", i, Ds.Val_UserInput)
    Call Database_SavePropertyWithIndex(Rs1, "Ds.Val_StEPP", i, Ds.Val_StEPP)
    Call Database_SavePropertyWithIndex(Rs1, "Ds.Val_Corr", i, Ds.Val_Corr)
  Next i
  Call Database_SaveProperty(Rs1, "Cd.O2_CInfinity", Cd.O2_CInfinity)
  For i = 0 To 18
    Call Database_SavePropertyWithIndex(Rs1, "Cd.UnitsOfDisplay", i, Cd.UnitsOfDisplay(i))
  Next i
  '
  ' MISCELLANEOUS.
  '
  Call Database_SaveProperty(Rs1, "UnitType", Prj.UnitType)
  'END SAVE TO THIS TABLE.
  Rs1.Close
  
  'CLOSE THE DATABASE FILE.
  Db1.Close

  'COMPACT THE DATABASE FILE.
      'TO DO: USE THE DbEngine.CompactDatabase METHOD
      'TO COMPACT THE DATABASE.  PROBLEM TO CONSIDER:
      'THE DB MUST BE COMPACTED TO A TEMPORARY FILE,
      'WHICH THEN SHOULD OVERWRITE THE ORIGINAL FILE.
  
  'RETURN A "SUCCESS" MESSAGE TO CALLER.
  File_Save_Latest_v1_00 = True
  
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
Dim InLine As String
Dim Dummy1 As String
Dim NewUnits As String
Dim H As Integer
  Call Database_LoadProperty(Rs1, InLine)
  ''''Call ProjectFile_Read(f, InLine, Dummy1)
  NewUnits = InLine
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




