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
Dim FileID As String
Dim FoundResinInList As Boolean
Dim ListIndexOfResin As Integer
Dim i As Integer


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
      'Call Project_SetDefaults(NowProj)
      '
      ' READ IN THE PROJECT DATA TO TEMPORARY VARIABLE.
      frmIonExchangeMain!cboIons(2).Clear
      Rs1.MoveFirst
      Prj.FileID = Trim$(UCase$(Database_Get_String(Rs1, "FieldName")))
      Rs1.MoveNext
      Do Until Rs1.EOF
        Use_FieldIndex = CInt(Database_Get_Long(Rs1, "FieldIndex"))
        Select Case Trim$(UCase$(Database_Get_String(Rs1, "FieldName")))
          'MAIN BLOCK.
          Case Trim$(UCase$("Pressure")): _
            Call Database_LoadProperty(Rs1, Prj.Operating.Pressure)
          Case Trim$(UCase$("Temperature")): _
            Call Database_LoadProperty(Rs1, Prj.Operating.Temperature)
          Case Trim$(UCase$("Bed Length")): _
            Call Database_LoadProperty(Rs1, Prj.Bed.length)
          Case Trim$(UCase$("Bed Diameter")): _
            Call Database_LoadProperty(Rs1, Prj.Bed.Diameter)
          Case Trim$(UCase$("Bed Weight")): _
            Call Database_LoadProperty(Rs1, Prj.Bed.Weight)
          Case Trim$(UCase$("Bed Flowrate Value")): _
            Call Database_LoadProperty(Rs1, Prj.Bed.Flowrate.Value)
          Case Trim$(UCase$("Bed Flowrate UserInput")): _
            Call Database_LoadProperty(Rs1, Prj.Bed.Flowrate.UserInput)
          Case Trim$(UCase$("Bed EBCT Value")): _
            Call Database_LoadProperty(Rs1, Prj.Bed.EBCT.Value)
          Case Trim$(UCase$("Bed EBCT UserInput")): _
            Call Database_LoadProperty(Rs1, Prj.Bed.EBCT.UserInput)
          Case Trim$(UCase$("Number of Beds (in series)")): _
            Call Database_LoadProperty(Rs1, Prj.Bed.NumberOfBeds)
          Case Trim$(UCase$("Resin Name")): _
            Call Database_LoadProperty(Rs1, Prj.Resin.Name)
            FoundResinInList = False
            For i = 0 To frmIonExchangeMain!cboAdsorbents.ListCount - 1
              If Trim$(frmIonExchangeMain!cboAdsorbents.List(i)) = _
                Trim$(Prj.Resin.Name) Then
                FoundResinInList = True
                ListIndexOfResin = i
                Exit For
              End If
            Next i
            If Not FoundResinInList Then
              frmIonExchangeMain!cboAdsorbents.AddItem Trim(Prj.Resin.Name)
              ListIndexOfResin = frmIonExchangeMain!cboAdsorbents.ListCount - 1
            End If
          Case Trim$(UCase$("Apparent Density")): _
            Call Database_LoadProperty(Rs1, Prj.Resin.ApparentDensity)
          Case Trim$(UCase$("Particle Radius")): _
            Call Database_LoadProperty(Rs1, Prj.Resin.ParticleRadius)
          Case Trim$(UCase$("Particle Porosity (-)")): _
            Call Database_LoadProperty(Rs1, Prj.Resin.ParticlePorosity)
          Case Trim$(UCase$("Tortuosity (-)")): _
            Call Database_LoadProperty(Rs1, Prj.Resin.Tortuosity)
          Case Trim$(UCase$("Total Resin Capacity")): _
            Call Database_LoadProperty(Rs1, Prj.Resin.TotalCapacity)
          Case Trim$(UCase$("Time Parameters - Total Run Time")): _
            Call Database_LoadProperty(Rs1, Prj.TimeParameters.FinalTime)
          Case Trim$(UCase$("Time Parameters - InitialTime")): _
            Call Database_LoadProperty(Rs1, Prj.TimeParameters.InitialTime)
          Case Trim$(UCase$("Time Parameters - Time Step")): _
            Call Database_LoadProperty(Rs1, Prj.TimeParameters.TimeStep)
'          If FileID <> "Ion Exchange Model l - Input File" Then
            Case Trim$(UCase$("EPS-ErrorCriteriaForDGEARIntegrator")): _
              Call Database_LoadProperty(Rs1, Prj.EPS_ErrorCriteriaForDGEARIntegrator)
            Case Trim$(UCase$("DH0_InitialTimeStepForDGEARIntegrator")): _
              Call Database_LoadProperty(Rs1, Prj.DH0_InitialTimeStepForDGEARIntegrator)
'          End If
          
          Case Trim$(UCase$("Number of Axial Collocation Points")): _
            Call Database_LoadProperty(Rs1, Prj.NumAxialCollocationPoints)
          Case Trim$(UCase$("Number of Radial Collocation Points")): _
            Call Database_LoadProperty(Rs1, Prj.NumRadialCollocationPoints)
          Case Trim$(UCase$("Correlation for Ionic Transport Coeff, kf")): _
            Call Database_LoadProperty(Rs1, Prj.IonicTransportCoeffCorrName)
          Case Trim$(UCase$("Number of Cations")): _
            Call Database_LoadProperty(Rs1, Prj.NumberOfCations)
          Case Trim$(UCase$("Presaturant Cation")): _
            Call Database_LoadProperty(Rs1, Prj.PresaturantCation)
          Case Trim$(UCase$("Sum of Cation Time-Averaged Initial Influent Concs.")): _
            Call Database_LoadProperty(Rs1, Prj.SumCationInitialEquivalents)
            Rs1.MoveNext
            Call Database_LoadProperty(Rs1, Prj.OKToGetCationDimensionless)
          Case Trim$(UCase$("Cation Separation Factor Info. Row")): _
            Call Database_LoadProperty(Rs1, Prj.CationSeparationFactorInput.Row)
          Case Trim$(UCase$("Cation Separation Factor Info. Value")): _
            Call Database_LoadProperty(Rs1, Prj.CationSeparationFactorInput.Value)
           kmtest = True
          Case Trim$(UCase$("Name")):
            If (Prj.NumberOfCations > 0) Then
                Cations.Available = True
                
                kmtest = True
                
                ReDim Prj.Cation(1 To Prj.NumberOfCations)
                For i = 1 To Prj.NumberOfCations
                  Do While i = Trim$(UCase$(Database_Get_Integer(Rs1, "FieldIndex")))
                    Select Case Trim$(UCase$(Database_Get_String(Rs1, "FieldName")))
                      Case Trim$(UCase$("Name")): _
                        Call Database_LoadProperty(Rs1, Prj.Cation(i).Name)
                      Case Trim$(UCase$("MolecularWeight")): _
                         Call Database_LoadProperty(Rs1, Prj.Cation(i).MolecularWeight)
                      Case Trim$(UCase$("InitialConcentration")): _
                         Call Database_LoadProperty(Rs1, Prj.Cation(i).InitialConcentration)
                      Case Trim$(UCase$("EquivalentInitialConcentration")): _
                         Call Database_LoadProperty(Rs1, Prj.Cation(i).EquivalentInitialConcentration)
                      Case Trim$(UCase$("Valence")): _
                         Call Database_LoadProperty(Rs1, Prj.Cation(i).Valence)
                      Case Trim$(UCase$("SeparationFactor")): _
                         Call Database_LoadProperty(Rs1, Prj.Cation(i).SeparationFactor)
                      Case Trim$(UCase$("Kinetic.LiquidDiffusivity.Value")): _
                         Call Database_LoadProperty(Rs1, Prj.Cation(i).Kinetic.LiquidDiffusivity.Value)
                      Case Trim$(UCase$("Kinetic.LiquidDiffusivity.UserInput")): _
                         Call Database_LoadProperty(Rs1, Prj.Cation(i).Kinetic.LiquidDiffusivity.UserInput)
                      Case Trim$(UCase$("Kinetic.LiquidDiffusivityCorrelation")): _
                         Call Database_LoadProperty(Rs1, Prj.Cation(i).Kinetic.LiquidDiffusivityCorrelation)
                      Case Trim$(UCase$("Kinetic.LiquidDiffusivityUserInput")): _
                         Call Database_LoadProperty(Rs1, Prj.Cation(i).Kinetic.LiquidDiffusivityUserInput)
                      Case Trim$(UCase$("Kinetic.IonicTransportCoefficient.Value")): _
                         Call Database_LoadProperty(Rs1, Prj.Cation(i).Kinetic.IonicTransportCoefficient.Value)
                      Case Trim$(UCase$("Kinetic.IonTransportCoefficient.UserInput")): _
                         Call Database_LoadProperty(Rs1, Prj.Cation(i).Kinetic.IonicTransportCoefficient.UserInput)
                      Case Trim$(UCase$("Kinetic.IonicTransportCoeffCorrelation")): _
                         Call Database_LoadProperty(Rs1, Prj.Cation(i).Kinetic.IonicTransportCoeffCorrelation)
                      Case Trim$(UCase$("Kinetic.IonicTransportCoeffUserInput")): _
                         Call Database_LoadProperty(Rs1, Prj.Cation(i).Kinetic.IonicTransportCoeffUserInput)
                      Case Trim$(UCase$("Kinetic.PoreDiffusivity.Value")): _
                         Call Database_LoadProperty(Rs1, Prj.Cation(i).Kinetic.PoreDiffusivity.Value)
                      Case Trim$(UCase$("Kinetic.PoreDiffusivity.UserInput")): _
                         Call Database_LoadProperty(Rs1, Prj.Cation(i).Kinetic.PoreDiffusivity.UserInput)
                      Case Trim$(UCase$("Kinetic.PoreDiffusivityCorrelation")): _
                         Call Database_LoadProperty(Rs1, Prj.Cation(i).Kinetic.PoreDiffusivityCorrelation)
                      Case Trim$(UCase$("Kinetic.PoreDiffusivityUserInput")): _
                         Call Database_LoadProperty(Rs1, Prj.Cation(i).Kinetic.PoreDiffusivityUserInput)
                      Case Trim$(UCase$("Kinetic.NernstHaskellCation.Ion_Name")): _
                         Call Database_LoadProperty(Rs1, Prj.Cation(i).Kinetic.NernstHaskellCation.Ion_Name)
                      Case Trim$(UCase$("Kinetic.NernstHaskellCation.Valence")): _
                         Call Database_LoadProperty(Rs1, Prj.Cation(i).Kinetic.NernstHaskellCation.Valence)
                      Case Trim$(UCase$("Kinetic.NernstHaskellCation.LimitingIonicConductance")): _
                         Call Database_LoadProperty(Rs1, Prj.Cation(i).Kinetic.NernstHaskellCation.LimitingIonicConductance)
                      Case Trim$(UCase$("Kinetic.NernstHaskellAnion.Ion_Name")): _
                         Call Database_LoadProperty(Rs1, Prj.Cation(i).Kinetic.NernstHaskellAnion.Ion_Name)
                      Case Trim$(UCase$("Kinetic.NernstHaskellAnion.Valence")): _
                         Call Database_LoadProperty(Rs1, Prj.Cation(i).Kinetic.NernstHaskellAnion.Valence)
                      Case Trim$(UCase$("Kinetic.NernstHaskellAnion.LimitingIonicConductance")): _
                         Call Database_LoadProperty(Rs1, Prj.Cation(i).Kinetic.NernstHaskellAnion.LimitingIonicConductance)
                    End Select
                    Rs1.MoveNext
                  Loop
                Next i
            End If

          Case Trim$(UCase$("Number of Anions")): _
            Call Database_LoadProperty(Rs1, Prj.NumberOfAnions)
           Case Trim$(UCase$("Presaturant Anion")): _
            Call Database_LoadProperty(Rs1, Prj.PresaturantAnion)
          Case Trim$(UCase$("Sum of Anion Time-Averaged Initial Influent Concs.")): _
            Call Database_LoadProperty(Rs1, Prj.SumAnionInitialEquivalents)
            Rs1.MoveNext
            Call Database_LoadProperty(Rs1, Prj.OKToGetAnionDimensionless)
          Case Trim$(UCase$("Anion Separation Factor Info. Row,")): _
            Call Database_LoadProperty(Rs1, Prj.AnionSeparationFactorInput.Row)
          Case Trim$(UCase$("Anion Separation Factor Info. Value")): _
            Call Database_LoadProperty(Rs1, Prj.AnionSeparationFactorInput.Value)
          
          Case Trim$(UCase$("Name")):
            If (Prj.NumberOfAnions > 0) Then
                Anions.Available = True
                ReDim Preserve Prj.Anion(1 To Prj.NumberOfAnions)
                For i = 1 To Prj.NumberOfAnions
                  Do While i = Trim$(UCase$(Database_Get_Integer(Rs1, "FieldIndex")))
                    Select Case Trim$(UCase$(Database_Get_String(Rs1, "FieldName")))
                      Case Trim$(UCase$("Name")): _
                        Call Database_LoadProperty(Rs1, Prj.Anion(i).Name)
                      Case Trim$(UCase$("MolecularWeight")): _
                         Call Database_LoadProperty(Rs1, Prj.Anion(i).MolecularWeight)
                      Case Trim$(UCase$("InitialConcentration")): _
                         Call Database_LoadProperty(Rs1, Prj.Anion(i).InitialConcentration)
                      Case Trim$(UCase$("EquivalentInitialConcentration")): _
                         Call Database_LoadProperty(Rs1, Prj.Anion(i).EquivalentInitialConcentration)
                      Case Trim$(UCase$("Valence")): _
                         Call Database_LoadProperty(Rs1, Prj.Anion(i).Valence)
                      Case Trim$(UCase$("SeparationFactor")): _
                         Call Database_LoadProperty(Rs1, Prj.Anion(i).SeparationFactor)
                      Case Trim$(UCase$("Kinetic.LiquidDiffusivity.Value")): _
                         Call Database_LoadProperty(Rs1, Prj.Anion(i).Kinetic.LiquidDiffusivity.Value)
                      Case Trim$(UCase$("Kinetic.LiquidDiffusivity.UserInput")): _
                         Call Database_LoadProperty(Rs1, Prj.Anion(i).Kinetic.LiquidDiffusivity.UserInput)
                      Case Trim$(UCase$("Kinetic.LiquidDiffusivityCorrelation")): _
                         Call Database_LoadProperty(Rs1, Prj.Anion(i).Kinetic.LiquidDiffusivityCorrelation)
                      Case Trim$(UCase$("Kinetic.LiquidDiffusivityUserInput")): _
                         Call Database_LoadProperty(Rs1, Prj.Anion(i).Kinetic.LiquidDiffusivityUserInput)
                      Case Trim$(UCase$("Kinetic.IonicTransportCoefficient.Value")): _
                         Call Database_LoadProperty(Rs1, Prj.Anion(i).Kinetic.IonicTransportCoefficient.Value)
                      Case Trim$(UCase$("Kinetic.IonTransportCoefficient.UserInput")): _
                         Call Database_LoadProperty(Rs1, Prj.Anion(i).Kinetic.IonicTransportCoefficient.UserInput)
                      Case Trim$(UCase$("Kinetic.IonicTransportCoeffCorrelation")): _
                         Call Database_LoadProperty(Rs1, Prj.Anion(i).Kinetic.IonicTransportCoeffCorrelation)
                      Case Trim$(UCase$("Kinetic.IonicTransportCoeffUserInput")): _
                         Call Database_LoadProperty(Rs1, Prj.Anion(i).Kinetic.IonicTransportCoeffUserInput)
                      Case Trim$(UCase$("Kinetic.PoreDiffusivity.Value")): _
                         Call Database_LoadProperty(Rs1, Prj.Anion(i).Kinetic.PoreDiffusivity.Value)
                      Case Trim$(UCase$("Kinetic.PoreDiffusivity.UserInput")): _
                         Call Database_LoadProperty(Rs1, Prj.Anion(i).Kinetic.PoreDiffusivity.UserInput)
                      Case Trim$(UCase$("Kinetic.PoreDiffusivityCorrelation")): _
                         Call Database_LoadProperty(Rs1, Prj.Anion(i).Kinetic.PoreDiffusivityCorrelation)
                      Case Trim$(UCase$("Kinetic.PoreDiffusivityUserInput")): _
                         Call Database_LoadProperty(Rs1, Prj.Anion(i).Kinetic.PoreDiffusivityUserInput)
                      Case Trim$(UCase$("Kinetic.NernstHaskellCation.Ion_Name")): _
                         Call Database_LoadProperty(Rs1, Prj.Anion(i).Kinetic.NernstHaskellCation.Ion_Name)
                      Case Trim$(UCase$("Kinetic.NernstHaskellCation.Valence")): _
                         Call Database_LoadProperty(Rs1, Prj.Anion(i).Kinetic.NernstHaskellCation.Valence)
                      Case Trim$(UCase$("Kinetic.NernstHaskellCation.LimitingIonicConductance")): _
                         Call Database_LoadProperty(Rs1, Prj.Anion(i).Kinetic.NernstHaskellCation.LimitingIonicConductance)
                      Case Trim$(UCase$("Kinetic.NernstHaskellAnion.Ion_Name")): _
                         Call Database_LoadProperty(Rs1, Prj.Anion(i).Kinetic.NernstHaskellAnion.Ion_Name)
                      Case Trim$(UCase$("Kinetic.NernstHaskellAnion.Valence")): _
                         Call Database_LoadProperty(Rs1, Prj.Anion(i).Kinetic.NernstHaskellAnion.Valence)
                      Case Trim$(UCase$("Kinetic.NernstHaskellAnion.LimitingIonicConductance")): _
                         Call Database_LoadProperty(Rs1, Prj.Anion(i).Kinetic.NernstHaskellAnion.LimitingIonicConductance)
                    End Select
                    Rs1.MoveNext
                  Loop
                Next i
            End If
          
          Case Trim$(UCase$("Name of File for Cation Variable Influent")): _
            Call Database_LoadProperty(Rs1, Prj.VarInfluentFileCation)
          Case Trim$(UCase$("Name of File for Anion Variable Influent")): _
            Call Database_LoadProperty(Rs1, Prj.VarInfluentFileAnion)
        End Select
        Rs1.MoveNext
      Loop
      '
      ' TRANSFER PROJECT DATA TO MEMORY.
      '
      NowProj = Prj
    End If
    Rs1.Close
  End If

  'CLOSE THE DATABASE FILE.
  Db1.Close

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
'  If (NeedToCreateNewDatabase) Then
'    Set Db1 = CreateDatabase(fn_This, dbLangGeneral)
'  End If
  If (NeedToCreateNewDatabase) Then
    FileCopy MAIN_APP_PATH & "\dbase\template.iex", fn_This
'    Set Db1 = CreateDatabase(fn_this, dbLangGeneral)
    Set Db1 = OpenDatabase(fn_This)
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
  Call Database_SaveProperty(Rs1, "Ion Exchange Model - Input File", NowProj.FileID)
  Call Database_SaveProperty(Rs1, "Pressure", NowProj.Operating.Pressure)
  Call Database_SaveProperty(Rs1, "Temperature", NowProj.Operating.Temperature)
  Call Database_SaveProperty(Rs1, "Bed Length", NowProj.Bed.length)
  Call Database_SaveProperty(Rs1, "Bed Diameter", NowProj.Bed.Diameter)
  Call Database_SaveProperty(Rs1, "Bed Weight", NowProj.Bed.Weight)
  Call Database_SaveProperty(Rs1, "Bed Flowrate Value", NowProj.Bed.Flowrate.Value)
  Call Database_SaveProperty(Rs1, "Bed Flowrate UserInput", NowProj.Bed.Flowrate.UserInput)
  Call Database_SaveProperty(Rs1, "Bed EBCT Value", NowProj.Bed.EBCT.Value)
  Call Database_SaveProperty(Rs1, "Bed EBCT UserInput", NowProj.Bed.EBCT.UserInput)
  Call Database_SaveProperty(Rs1, "Number of Beds (in series)", NowProj.Bed.NumberOfBeds)
  
  Call Database_SaveProperty(Rs1, "Resin Name", NowProj.Resin.Name) '?????
  
  Call Database_SaveProperty(Rs1, "Apparent Density", NowProj.Resin.ApparentDensity)
  Call Database_SaveProperty(Rs1, "Particle Radius", NowProj.Resin.ParticleRadius)
  Call Database_SaveProperty(Rs1, "Particle Porosity (-)", NowProj.Resin.ParticlePorosity)
  Call Database_SaveProperty(Rs1, "Tortuosity (-)", NowProj.Resin.Tortuosity)
  Call Database_SaveProperty(Rs1, "Total Resin Capacity", NowProj.Resin.TotalCapacity)
  Call Database_SaveProperty(Rs1, "Time Parameters - Total Run Time", NowProj.TimeParameters.FinalTime)
  Call Database_SaveProperty(Rs1, "Time Parameters - InitialTime", NowProj.TimeParameters.InitialTime)
  Call Database_SaveProperty(Rs1, "Time Parameters - Time Step", NowProj.TimeParameters.TimeStep)
  Call Database_SaveProperty(Rs1, "EPS-ErrorCriteriaForDGEARIntegrator", NowProj.EPS_ErrorCriteriaForDGEARIntegrator)
  Call Database_SaveProperty(Rs1, "DH0_InitialTimeStepForDGEARIntegrator", NowProj.DH0_InitialTimeStepForDGEARIntegrator)
  Call Database_SaveProperty(Rs1, "Number of Axial Collocation Points", NowProj.NumAxialCollocationPoints)
  Call Database_SaveProperty(Rs1, "Number of Radial Collocation Points", NowProj.NumRadialCollocationPoints)
  Call Database_SaveProperty(Rs1, "Correlation for Ionic Transport Coeff.", NowProj.IonicTransportCoeffCorrName)
  Call Database_SaveProperty(Rs1, "Number of Cations", NowProj.NumberOfCations)
  Call Database_SaveProperty(Rs1, "Presaturant Cation", NowProj.PresaturantCation)
  Call Database_SaveProperty(Rs1, "Sum of Cation Time-Averaged Initial Influent Concs.", _
    NowProj.SumCationInitialEquivalents)
  Call Database_SaveProperty(Rs1, "OKToGetCationDimensionless", NowProj.OKToGetCationDimensionless)
  Call Database_SaveProperty(Rs1, "Cation Separation Factor Info. Row", NowProj.CationSeparationFactorInput.Row)
  Call Database_SaveProperty(Rs1, "Cation Separation Factor Info. Value", NowProj.CationSeparationFactorInput.Value)
  For i = 1 To NowProj.NumberOfCations
    Call Database_SavePropertyWithIndex(Rs1, "Name", i, NowProj.Cation(i).Name)
    Call Database_SavePropertyWithIndex(Rs1, "MolecularWeight", i, NowProj.Cation(i).MolecularWeight)
    Call Database_SavePropertyWithIndex(Rs1, "InitialConcentration", i, NowProj.Cation(i).InitialConcentration)
    Call Database_SavePropertyWithIndex(Rs1, "EquivalentInitialConcentration", i, NowProj.Cation(i).EquivalentInitialConcentration)
    Call Database_SavePropertyWithIndex(Rs1, "Valence", i, NowProj.Cation(i).Valence)
    Call Database_SavePropertyWithIndex(Rs1, "SeparationFactor", i, NowProj.Cation(i).SeparationFactor)
    Call Database_SavePropertyWithIndex(Rs1, "Kinetic.LiquidDiffusivity.Value", i, NowProj.Cation(i).Kinetic.LiquidDiffusivity.Value)
    Call Database_SavePropertyWithIndex(Rs1, "Kinetic.LiquidDiffusivity.UserInput", i, NowProj.Cation(i).Kinetic.LiquidDiffusivity.UserInput)
    Call Database_SavePropertyWithIndex(Rs1, "Kinetic.LiquidDiffusivityCorrelation", i, NowProj.Cation(i).Kinetic.LiquidDiffusivityCorrelation)
    Call Database_SavePropertyWithIndex(Rs1, "Kinetic.LiquidDiffusivityUserInput", i, NowProj.Cation(i).Kinetic.LiquidDiffusivityUserInput)
    Call Database_SavePropertyWithIndex(Rs1, "Kinetic.IonicTransportCoefficient.Value", i, NowProj.Cation(i).Kinetic.IonicTransportCoefficient.Value)
    Call Database_SavePropertyWithIndex(Rs1, "Kinetic.IonicTransportCoefficient.UserInput", i, NowProj.Cation(i).Kinetic.IonicTransportCoefficient.UserInput)
    Call Database_SavePropertyWithIndex(Rs1, "Kinetic.IonicTransportCoeffCorrelation", i, NowProj.Cation(i).Kinetic.IonicTransportCoeffCorrelation)
    Call Database_SavePropertyWithIndex(Rs1, "Kinetic.IonicTransportCoeffUserInput", i, NowProj.Cation(i).Kinetic.IonicTransportCoeffUserInput)
    Call Database_SavePropertyWithIndex(Rs1, "Kinetic.PoreDiffusivity.Value", i, NowProj.Cation(i).Kinetic.PoreDiffusivity.Value)
    Call Database_SavePropertyWithIndex(Rs1, "Kinetic.PoreDiffusivity.UserInput", i, NowProj.Cation(i).Kinetic.PoreDiffusivity.UserInput)
    Call Database_SavePropertyWithIndex(Rs1, "Kinetic.PoreDiffusivityCorrelation", i, NowProj.Cation(i).Kinetic.PoreDiffusivityCorrelation)
    Call Database_SavePropertyWithIndex(Rs1, "Kinetic.PoreDiffusivityUserInput", i, NowProj.Cation(i).Kinetic.PoreDiffusivityUserInput)
    Call Database_SavePropertyWithIndex(Rs1, "Kinetic.NernstHaskellCation.Ion_Name", i, NowProj.Cation(i).Kinetic.NernstHaskellCation.Ion_Name)
    Call Database_SavePropertyWithIndex(Rs1, "Kinetic.NernstHaskellCation.Valence", i, NowProj.Cation(i).Kinetic.NernstHaskellCation.Valence)
    Call Database_SavePropertyWithIndex(Rs1, "Kinetic.NernstHaskellCation.LimitingIonicConductance", i, NowProj.Cation(i).Kinetic.NernstHaskellCation.LimitingIonicConductance)
    Call Database_SavePropertyWithIndex(Rs1, "Kinetic.NernstHaskellAnion.Ion_Name", i, NowProj.Cation(i).Kinetic.NernstHaskellAnion.Ion_Name)
    Call Database_SavePropertyWithIndex(Rs1, "Kinetic.NernstHaskellAnion.Valence", i, NowProj.Cation(i).Kinetic.NernstHaskellAnion.Valence)
    Call Database_SavePropertyWithIndex(Rs1, "Kinetic.NernstHaskellAnion.LimitingIonicConductance", i, NowProj.Cation(i).Kinetic.NernstHaskellAnion.LimitingIonicConductance)
  Next i
  Call Database_SaveProperty(Rs1, "Number of Anions", NowProj.NumberOfAnions)
  Call Database_SaveProperty(Rs1, "Presaturant Anion", NowProj.PresaturantAnion)
  Call Database_SaveProperty(Rs1, "Sum of Anion Time-Averaged Initial Influent Concs.", _
    NowProj.SumAnionInitialEquivalents)
  Call Database_SaveProperty(Rs1, "OKToGetAnionDimensionless", NowProj.OKToGetAnionDimensionless)
  Call Database_SaveProperty(Rs1, "Anion Separation Factor Info. Row", NowProj.CationSeparationFactorInput.Row)
  Call Database_SaveProperty(Rs1, "Anion Separation Factor Info. Value", NowProj.CationSeparationFactorInput.Value)
  For i = 1 To NowProj.NumberOfAnions
    Call Database_SavePropertyWithIndex(Rs1, "Name", i, NowProj.Anion(i).Name)
    Call Database_SavePropertyWithIndex(Rs1, "MolecularWeight", i, NowProj.Anion(i).MolecularWeight)
    Call Database_SavePropertyWithIndex(Rs1, "InitialConcentration", i, NowProj.Anion(i).InitialConcentration)
    Call Database_SavePropertyWithIndex(Rs1, "EquivalentInitialConcentration", i, NowProj.Anion(i).EquivalentInitialConcentration)
    Call Database_SavePropertyWithIndex(Rs1, "Valence", i, NowProj.Anion(i).Valence)
    Call Database_SavePropertyWithIndex(Rs1, "SeparationFactor", i, NowProj.Anion(i).SeparationFactor)
    Call Database_SavePropertyWithIndex(Rs1, "Kinetic.LiquidDiffusivity.Value", i, NowProj.Anion(i).Kinetic.LiquidDiffusivity.Value)
    Call Database_SavePropertyWithIndex(Rs1, "Kinetic.LiquidDiffusivity.UserInput", i, NowProj.Anion(i).Kinetic.LiquidDiffusivity.UserInput)
    Call Database_SavePropertyWithIndex(Rs1, "Kinetic.LiquidDiffusivityCorrelation", i, NowProj.Anion(i).Kinetic.LiquidDiffusivityCorrelation)
    Call Database_SavePropertyWithIndex(Rs1, "Kinetic.LiquidDiffusivityUserInput", i, NowProj.Anion(i).Kinetic.LiquidDiffusivityUserInput)
    Call Database_SavePropertyWithIndex(Rs1, "Kinetic.IonicTransportCoefficient.Value", i, NowProj.Anion(i).Kinetic.IonicTransportCoefficient.Value)
    Call Database_SavePropertyWithIndex(Rs1, "Kinetic.IonTransportCoefficient.UserInput", i, NowProj.Anion(i).Kinetic.IonicTransportCoefficient.UserInput)
    Call Database_SavePropertyWithIndex(Rs1, "Kinetic.IonicTransportCoeffCorrelation", i, NowProj.Anion(i).Kinetic.IonicTransportCoeffCorrelation)
    Call Database_SavePropertyWithIndex(Rs1, "Kinetic.IonicTransportCoeffUserInput", i, NowProj.Anion(i).Kinetic.IonicTransportCoeffUserInput)
    Call Database_SavePropertyWithIndex(Rs1, "Kinetic.PoreDiffusivity.Value", i, NowProj.Anion(i).Kinetic.PoreDiffusivity.Value)
    Call Database_SavePropertyWithIndex(Rs1, "Kinetic.PoreDiffusivity.UserInput", i, NowProj.Anion(i).Kinetic.PoreDiffusivity.UserInput)
    Call Database_SavePropertyWithIndex(Rs1, "Kinetic.PoreDiffusivityCorrelation", i, NowProj.Anion(i).Kinetic.PoreDiffusivityCorrelation)
    Call Database_SavePropertyWithIndex(Rs1, "Kinetic.PoreDiffusivityUserInput", i, NowProj.Anion(i).Kinetic.PoreDiffusivityUserInput)
    Call Database_SavePropertyWithIndex(Rs1, "Kinetic.NernstHaskellCation.Ion_Name", i, NowProj.Anion(i).Kinetic.NernstHaskellCation.Ion_Name)
    Call Database_SavePropertyWithIndex(Rs1, "Kinetic.NernstHaskellCation.Valence", i, NowProj.Anion(i).Kinetic.NernstHaskellCation.Valence)
    Call Database_SavePropertyWithIndex(Rs1, "Kinetic.NernstHaskellCation.LimitingIonicConductance", i, NowProj.Anion(i).Kinetic.NernstHaskellCation.LimitingIonicConductance)
    Call Database_SavePropertyWithIndex(Rs1, "Kinetic.NernstHaskellAnion.Ion_Name", i, NowProj.Anion(i).Kinetic.NernstHaskellAnion.Ion_Name)
    Call Database_SavePropertyWithIndex(Rs1, "Kinetic.NernstHaskellAnion.Valence", i, NowProj.Anion(i).Kinetic.NernstHaskellAnion.Valence)
    Call Database_SavePropertyWithIndex(Rs1, "Kinetic.NernstHaskellAnion.LimitingIonicConductance", i, NowProj.Anion(i).Kinetic.NernstHaskellAnion.LimitingIonicConductance)
  Next i
   
  Call Database_SaveProperty(Rs1, "Name of File for Cation Variable Influent", NowProj.VarInfluentFileCation)
  Call Database_SaveProperty(Rs1, "Name of File for Anion Variable Influent", NowProj.VarInfluentFileAnion)

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
  If (left$(in_Use_FieldName, 1) = "*") Then
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
  If (left$(in_Use_FieldName, 1) = "*") Then
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




