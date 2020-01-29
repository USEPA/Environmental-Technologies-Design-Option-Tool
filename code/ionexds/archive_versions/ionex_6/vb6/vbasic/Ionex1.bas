Attribute VB_Name = "ionex1"
Option Explicit

Global Const SCREEN_WIDTH_STANDARD = 9600
Global Const SCREEN_HEIGHT_STANDARD = 7200

Global Const TOLERANCE = 0.000000000000001

Global Const NERNST_HASKELL_ANION_DB_IDENTIFIER = "Nernst_Haskell_Anion_Database"
Global Const NERNST_HASKELL_CATION_DB_IDENTIFIER = "Nernst_Haskell_Cation_Database"

Global filename As String
Global OldFileName As String

'------Begin Modification Hokanson: 11-Aug2000
'Global Const ION_EXCHANGE_FILEID = "Ion Exchange Model - Input File"
Global Const ION_EXCHANGE_FILEID = "Ion Exchange Model - Input File - I/O v2000aug11"
'------Begin Modification Hokanson: 11-Aug2000

'Path variables
Global IonExchangePath As String
Global SaveAndLoadPath As String

'Tabs used by Printing Routines
Global Const TAB_BED_DATA = 25
Global Const TAB_RESIN_PROPERTIES = TAB_BED_DATA
Global Const TAB_OPERATING_CONDITIONS = TAB_BED_DATA
Global Const TAB_COMPONENT_PROPERTIES_1 = 30
Global Const TAB_COMPONENT_PROPERTIES_2 = 45
Global Const TAB_COMPONENT_PROPERTIES_3 = 60

Global Const TAB_SEPARATION_FACTORS_1 = 6
Global Const TAB_SEPARATION_FACTORS_INTERVAL = 10

Global Const TAB_KINETIC_PARAMETERS_1 = 30
Global Const TAB_KINETIC_PARAMETERS_2 = 47
Global Const TAB_KINETIC_PARAMETERS_3 = 64

Global Const TAB_DIMENSIONLESS_GROUPS_1 = 20
Global Const TAB_DIMENSIONLESS_GROUPS_2 = 30
Global Const TAB_DIMENSIONLESS_GROUPS_3 = 40
Global Const TAB_DIMENSIONLESS_GROUPS_4 = 50
Global Const TAB_DIMENSIONLESS_GROUPS_5 = 60
Global Const TAB_DIMENSIONLESS_GROUPS_6 = 70

Global Const PRINT_FONTSIZE_DATA = 10
Global Const PRINTER_FONT = "Courier New"

Function EBCT(Bed_Length As Double, Bed_Diameter As Double, Bed_Flowrate As Double) As Double
   'Given:
   '   Bed_Length in m
   '   Bed_Diameter in m
   '   Bed_Flowrate in m3/s
   'Returns EBCT in min

   EBCT = Bed_Length * Pi * Bed_Diameter * Bed_Diameter / 4# / Bed_Flowrate / 60#

End Function

Function File_Get_Rid_Of_Path(filename As String) As String
Dim L As Integer, i As Integer, temp As String, M  As Integer
  L = Len(filename)
  For i = L To 1 Step -1
    temp = Mid$(filename, i, 1)
    If temp = "\" Then
     M = i
     Exit For
    End If
  Next i
  File_Get_Rid_Of_Path = Mid$(filename, M + 1)
End Function

Function FileNameIsValid(Name_Of_File As String, Box As CommonDialog) As Integer
Dim TemporaryName As String, response As Integer

    TemporaryName = Box.filename
    If IsValidPath(TemporaryName, "C:") Then
      TemporaryName = Mid$(TemporaryName, 1, Len(TemporaryName) - 1)
      Name_Of_File = TemporaryName
    Else
      Box.filename = ""
      MsgBox "This file name is not valid.", 48, App.title
      FileNameIsValid = False
      Exit Function
    End If

    If Dir(Name_Of_File) <> "" Then         ' File already exists, so ask if overwriting is desired.
      response = MsgBox("Overwrite existing file " & Name_Of_File & " ?", MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2, App.title)
      If response = IDNO Then
        FileNameIsValid = False
        Exit Function
      End If
    End If
    FileNameIsValid = True
End Function

Function Flowrate(Bed_Length As Double, Bed_Diameter As Double, Bed_EBCT As Double) As Double
   'Given:
   '   Bed_Length in m
   '   Bed_Diameter in m
   '   EBCT in min
   'Returns Flowrate in m3/s

   Flowrate = Bed_Length * Pi * Bed_Diameter * Bed_Diameter / 4# / Bed_EBCT / 60#

End Function

Function GetTheFormat(Value As Double) As String
   Dim AbsValue As Double

   AbsValue = Abs(Value)

   If AbsValue < 0.001 Then
      GetTheFormat = "0.00E+00"
   ElseIf AbsValue < 0.01 Then
      GetTheFormat = "0.00E+00"
   ElseIf AbsValue < 0.1 Then
      GetTheFormat = "0.0000"
   ElseIf AbsValue < 1 Then
      GetTheFormat = "0.000"
   ElseIf AbsValue < 10 Then
      GetTheFormat = "0.00"
   ElseIf AbsValue < 100 Then
      GetTheFormat = "0.0"
   ElseIf AbsValue < 1000 Then
      GetTheFormat = "0"
   Else
      GetTheFormat = "0.00E+00"
   End If

End Function

Function HaveValue(Value As Double) As Integer

    If Value > 0# Then HaveValue = True Else HaveValue = False

End Function

Sub InitializeAvailableIons()

    'Load Presaturant Combo Boxes on Main Form
    frmIonExchangeMain!cboIons(0).AddItem "H+"
    frmIonExchangeMain!cboIons(0).AddItem "Na+"
    frmIonExchangeMain!cboIons(0).AddItem "K+"
    frmIonExchangeMain!cboIons(0).AddItem "Ca+2"
    NowProj.Cation(1).Name = "H+"
    NowProj.Cation(2).Name = "Na+"
    NowProj.Cation(3).Name = "K+"
    NowProj.Cation(4).Name = "Ca+2"
     NowProj.NumberOfCations = 4
    frmIonExchangeMain!cboIons(0).ListIndex = 0
    frmIonExchangeMain!cboIons(1).AddItem "OH-"
    frmIonExchangeMain!cboIons(1).AddItem "Cl-"
    NowProj.Anion(1).Name = "OH-"
    NowProj.Anion(2).Name = "Cl-"
    NowProj.NumberOfAnions = 2
    frmIonExchangeMain!cboIons(1).ListIndex = 0
    

End Sub

Sub InitializeDefaultIonProperties()

    'initialize default anion properties
     DefaultAnion.MolecularWeight = 35.45
     DefaultAnion.InitialConcentration = 100#
     DefaultAnion.SeparationFactor = 1#
     DefaultAnion.Valence = 1#
     DefaultAnion.EquivalentInitialConcentration = DefaultAnion.InitialConcentration * ConcentrationConversionFactor(CONCENTRATION_MEQ_per_L, DefaultAnion.Valence, DefaultAnion.MolecularWeight)
     DefaultAnion.Kinetic.NernstHaskellAnion = NernstHaskell.DefaultAnion
     DefaultAnion.Kinetic.NernstHaskellCation = NernstHaskell.DefaultCation
     DefaultAnion.Kinetic.LiquidDiffusivity.Value = 0.0000161
     DefaultAnion.Kinetic.IonicTransportCoefficient.Value = 0.00415
   
    'initialize default cation properties
     DefaultCation.MolecularWeight = 22.99
     DefaultCation.InitialConcentration = 100#
     DefaultCation.SeparationFactor = 1#
     DefaultCation.Valence = 1#
     DefaultCation.EquivalentInitialConcentration = DefaultCation.InitialConcentration * ConcentrationConversionFactor(CONCENTRATION_MEQ_per_L, DefaultCation.Valence, DefaultCation.MolecularWeight)
     DefaultCation.Kinetic.NernstHaskellCation = NernstHaskell.DefaultCation
     DefaultCation.Kinetic.NernstHaskellAnion = NernstHaskell.DefaultAnion
     DefaultCation.Kinetic.LiquidDiffusivity.Value = 0.0000161
     DefaultCation.Kinetic.IonicTransportCoefficient.Value = 0.00415

    'Initialize ChangedIon so Program doesn't crash
    ChangedIon = DefaultCation


End Sub

Sub InitializeIonExchangeParameters()

    'Operating Conditions

     frmIonExchangeMain!txtOperatingConditions(0).Text = "101325"
      NowProj.Operating.Pressure = CDbl(frmIonExchangeMain!txtOperatingConditions(0).Text)

     frmIonExchangeMain!txtOperatingConditions(1).Text = "294.15"
      NowProj.Operating.Temperature = CDbl(frmIonExchangeMain!txtOperatingConditions(1).Text)


    'Bed Data

    frmIonExchangeMain!txtBedData(0).Text = "0.114"
    NowProj.Bed.length = CDbl(frmIonExchangeMain!txtBedData(0).Text)

    frmIonExchangeMain!txtBedData(1).Text = "0.026856"
    NowProj.Bed.Diameter = CDbl(frmIonExchangeMain!txtBedData(1).Text)

    frmIonExchangeMain!txtBedData(2).Text = "0.056136"
    NowProj.Bed.Weight = CDbl(frmIonExchangeMain!txtBedData(2).Text)

    frmIonExchangeMain!txtBedData(3).Text = "8.279E-07"
    NowProj.Bed.Flowrate.Value = CDbl(frmIonExchangeMain!txtBedData(3).Text)
    NowProj.Bed.Flowrate.UserInput = True

    NowProj.Bed.EBCT.UserInput = False
    NowProj.Bed.EBCT.Value = EBCT(NowProj.Bed.length, NowProj.Bed.Diameter, NowProj.Bed.Flowrate.Value)
    frmIonExchangeMain!txtBedData(4).Text = Format$(NowProj.Bed.EBCT.Value, "0.00E+00")

    NowProj.Bed.NumberOfBeds = 1


    'Adsorbent Properties

    frmIonExchangeMain!cboAdsorbents.AddItem "IRN-77"
    frmIonExchangeMain!cboAdsorbents.AddItem "IRN-78"
    frmIonExchangeMain!cboAdsorbents.AddItem "IRA-68"
    frmIonExchangeMain!cboAdsorbents.ListIndex = 0
    NowProj.Resin.Name = Trim$(frmIonExchangeMain!cboAdsorbents.Text)

    frmIonExchangeMain!txtAdsorbentProperties(1).Text = "1.22"
    NowProj.Resin.ApparentDensity = CDbl(frmIonExchangeMain!txtAdsorbentProperties(1).Text)

    frmIonExchangeMain!txtAdsorbentProperties(2).Text = "0.0002975"
    NowProj.Resin.ParticleRadius = CDbl(frmIonExchangeMain!txtAdsorbentProperties(2).Text)

    frmIonExchangeMain!txtAdsorbentProperties(3).Text = "0.6289"
    NowProj.Resin.ParticlePorosity = CDbl(frmIonExchangeMain!txtAdsorbentProperties(3).Text)

    frmIonExchangeMain!txtAdsorbentProperties(4).Text = "1.00"
    NowProj.Resin.Tortuosity = CDbl(frmIonExchangeMain!txtAdsorbentProperties(4).Text)
    Results.Use_Tortuosity_Correlation = False
    Results.Constant_Tortuosity = True

    frmIonExchangeMain!txtAdsorbentProperties(5).Text = "2.1311"
    NowProj.Resin.TotalCapacity = CDbl(frmIonExchangeMain!txtAdsorbentProperties(5).Text)

    Call CalculateLiquidDensity
    Call CalculateLiquidViscosity
    Call CalculateParticleDiameter
    Call CalculateBedArea
    Call CalculateBedVolume
    Call CalculateBedDensity
    Call CalculateBedPorosity
    Call CalculateEffectiveContactTime
    Call CalculateSuperficialVelocity
    Call CalculateInterstitialVelocity
    

End Sub

Sub InitializeSeparationFactorInfo()

    NowProj.CationSeparationFactorInput.Row = True
    NowProj.CationSeparationFactorInput.Value = 11

     NowProj.AnionSeparationFactorInput.Row = True
     NowProj.AnionSeparationFactorInput.Value = 11

End Sub

Sub InitializeTimeAndCollocationInfo()

    NowProj.TimeParameters.InitialTime = 1#
    NowProj.TimeParameters.FinalTime = 399#
    NowProj.TimeParameters.TimeStep = 1#

    NowProj.NumAxialCollocationPoints = 15
    NowProj.NumRadialCollocationPoints = 6

'------Begin Modification Hokanson: 11-Aug2000
    EPS_ErrorCriteriaForDGEARIntegrator = 0.00001
    DH0_InitialTimeStepForDGEARIntegrator = 0.000001
'------End Modification Hokanson: 11-Aug2000

End Sub

'-------------------------------------------------------'
' Function:   IsValidPath as integer
' arguments:  DestPath$         a string that is a full path
'             DefaultDrive$     the default drive.  eg.  "C:"
'
'  If DestPath$ does not include a drive specification,
'  IsValidPath uses Default Drive
'
'  When IsValidPath is finished, DestPath$ is reformated
'  to the format "X:\dir\dir\dir\"
'
' Result:  True (-1) if path is valid.
'          False (0) if path is invalid
'-------------------------------------------------------
'
Function IsValidPath(DestPath$, ByVal DefaultDrive$) As Integer

    Dim Drive As String, legalChar As String, BackPos As Variant, forePos As Variant
    Dim temp As String, i As Integer, periodPos As Variant, length As Variant
    '----------------------------
    ' Remove left and right spaces
    '----------------------------
    DestPath$ = RTrim$(LTrim$(DestPath$))
    

    '-----------------------------
    ' Check Default Drive Parameter
    '-----------------------------
    If Right$(DefaultDrive$, 1) <> ":" Or Len(DefaultDrive$) <> 2 Then
        MsgBox "Bad default drive parameter specified in IsValidPath Function.  You passed,  """ + DefaultDrive$ + """.  Must be one drive letter and "":"".  For example, ""C:"", ""D:""...", 64, "Setup Kit Error"
        GoTo parseErr
    End If
    

    '-------------------------------------------------------
    ' Insert default drive if path begins with root backslash
    '-------------------------------------------------------
    If left$(DestPath$, 1) = "\" Then
        DestPath$ = DefaultDrive + DestPath$
    End If
    
    '-----------------------------
    ' check for invalid characters
    '-----------------------------
    On Error Resume Next
    Dim tmp As String
    tmp = Dir$(DestPath$)
    If Err <> 0 Then
        GoTo parseErr
    End If
    

    '-----------------------------------------
    ' Check for wildcard characters and spaces
    '-----------------------------------------
    If (InStr(DestPath$, "*") <> 0) Then GoTo parseErr
    If (InStr(DestPath$, "?") <> 0) Then GoTo parseErr
    If (InStr(DestPath$, " ") <> 0) Then GoTo parseErr
         
    
    '------------------------------------------
    ' Make Sure colon is in second char position
    '------------------------------------------
    If Mid$(DestPath$, 2, 1) <> Chr$(58) Then GoTo parseErr
    

    '-------------------------------
    ' Insert root backslash if needed
    '-------------------------------
    If Len(DestPath$) > 2 Then
      If Right$(left$(DestPath$, 3), 1) <> "\" Then
        DestPath$ = left$(DestPath$, 2) + "\" + Right$(DestPath$, Len(DestPath$) - 2)
      End If
    End If

    '-------------------------
    ' Check drive to install on
    '-------------------------
    
    Drive$ = left$(DestPath$, 1)
    ChDrive (Drive$)                                                        ' Try to change to the dest drive
    If Err <> 0 Then GoTo parseErr
    
    '-----------
    ' Add final \
    '-----------
    If Right$(DestPath$, 1) <> "\" Then
        DestPath$ = DestPath$ + "\"
    End If
    

    '-------------------------------------
    ' Root dir is a valid dir
    '-------------------------------------
    If Len(DestPath$) = 3 Then
        If Right$(DestPath$, 2) = ":\" Then
            GoTo ParseOK
        End If
    End If
    

    '------------------------
    ' Check for repeated Slash
    '------------------------
    If InStr(DestPath$, "\\") <> 0 Then GoTo parseErr
        
    '--------------------------------------
    ' Check for illegal directory names
    '--------------------------------------
    legalChar$ = "!#$%&'()-0123456789@ABCDEFGHIJKLMNOPQRSTUVWXYZ^_`{}~.üäöÄÖÜß"
    BackPos = 3
    forePos = InStr(4, DestPath$, "\")
    Do
        temp$ = Mid$(DestPath$, BackPos + 1, forePos - BackPos - 1)
        
        '----------------------------
        ' Test for illegal characters
        '----------------------------
        For i = 1 To Len(temp$)
            If InStr(legalChar$, UCase$(Mid$(temp$, i, 1))) = 0 Then GoTo parseErr
        Next i

        '-------------------------------------------
        ' Check combinations of periods and lengths
        '-------------------------------------------
        periodPos = InStr(temp$, ".")
        length = Len(temp$)
        If periodPos = 0 Then
            If length > 8 Then GoTo parseErr                         ' Base too long
        Else
            If periodPos > 9 Then GoTo parseErr                      ' Base too long
            If length > periodPos + 3 Then GoTo parseErr             ' Extension too long
            If InStr(periodPos + 1, temp$, ".") <> 0 Then GoTo parseErr ' Two periods not allowed
        End If

        BackPos = forePos
        forePos = InStr(BackPos + 1, DestPath$, "\")
    Loop Until forePos = 0

ParseOK:
    IsValidPath = True
    Exit Function

parseErr:
    IsValidPath = False
End Function

Sub LoadFileIonExchange(filename As String)

    On Error Resume Next
    frmIonExchangeMain!CMDialog1.DefaultExt = "iex"
    frmIonExchangeMain!CMDialog1.Filter = "Ion Exchange Files (*.iex)|*.iex"
    frmIonExchangeMain!CMDialog1.DialogTitle = "Load Ion Exchange File"
    frmIonExchangeMain!CMDialog1.flags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST
'------Begin Modification Hokanson: 12-Aug2000
    frmIonExchangeMain!CMDialog1.CancelError = True
'------End Modification Hokanson: 11-Aug2000
    frmIonExchangeMain!CMDialog1.Action = 1
    filename$ = frmIonExchangeMain!CMDialog1.filename
    If Err = 32755 Then   'Cancel selected by user
       filename$ = ""
    End If

End Sub

Sub LoadIonExchange(strFilename)
    Dim FileID As String, msg As String
    Dim i As Integer
    Dim NamePlusCAS As String
    Dim DummyString As String
    Dim StringLength As Integer, MainString As String
    Dim UnitsOfStorage As String, UnitsOfDisplay As String
    Dim FoundResinInList As Integer
    Dim FileIon As String, FileCation As String, FileAnion As String
    Dim ListIndexOfResin As Integer
    Dim ListIndex As Integer

    If strFilename = "" Then
      Call LoadFileIonExchange(filename)
    Else
      filename$ = strFilename
    End If
    
    If filename$ <> "" Then
       FileID = ""
       Open filename$ For Input As #1
       
       Input #1, FileID
      
       'Input Pressure and Temperature
       Input #1, DummyString, NowProj.Operating.Pressure, UnitsOfStorage, UnitsOfDisplay
'------BEGIN MODIFICATIONS HOKANSON 12-AUG-2000
''''       frmIonExchangeMain!txtOperatingConditions(0).Text = Format$( nowproj.Operating.Pressure, GetTheFormat( nowproj.Operating.Pressure))
       frmIonExchangeMain!cboOperatingConditionsUnits(0).ListIndex = -1
'------END MODIFICATIONS HOKANSON 12-AUG-2000
       For i = 0 To frmIonExchangeMain!cboOperatingConditionsUnits(0).ListCount - 1
           If Trim$(frmIonExchangeMain!cboOperatingConditionsUnits(0).List(i)) = Trim$(UnitsOfDisplay) Then
              frmIonExchangeMain!cboOperatingConditionsUnits(0).ListIndex = i
              Exit For
           End If
       Next i

       Input #1, DummyString, NowProj.Operating.Temperature, UnitsOfStorage, UnitsOfDisplay
'------BEGIN MODIFICATIONS HOKANSON 12-AUG-2000
''''       frmIonExchangeMain!txtOperatingConditions(1).Text = Format$( nowproj.Operating.Temperature, GetTheFormat( nowproj.Operating.Temperature))
       frmIonExchangeMain!cboOperatingConditionsUnits(1).ListIndex = -1
'------END MODIFICATIONS HOKANSON 12-AUG-2000
       For i = 0 To frmIonExchangeMain!cboOperatingConditionsUnits(1).ListCount - 1
           If Trim$(frmIonExchangeMain!cboOperatingConditionsUnits(1).List(i)) = Trim$(UnitsOfDisplay) Then
              frmIonExchangeMain!cboOperatingConditionsUnits(1).ListIndex = i
              Exit For
           End If
       Next i


       'Input Bed Data

       Input #1, DummyString, NowProj.Bed.length, UnitsOfStorage, UnitsOfDisplay
'------BEGIN MODIFICATIONS HOKANSON 12-AUG-2000
''''       frmIonExchangeMain!txtBedData(0).Text = Format$(nowproj.bed.Length, GetTheFormat(nowproj.bed.Length))
       frmIonExchangeMain!cboBedDataUnits(0).ListIndex = -1
'------END MODIFICATIONS HOKANSON 12-AUG-2000
       For i = 0 To frmIonExchangeMain!cboBedDataUnits(0).ListCount - 1
           If Trim$(frmIonExchangeMain!cboBedDataUnits(0).List(i)) = Trim$(UnitsOfDisplay) Then
              frmIonExchangeMain!cboBedDataUnits(0).ListIndex = i
              Exit For
           End If
       Next i

       Input #1, DummyString, NowProj.Bed.Diameter, UnitsOfStorage, UnitsOfDisplay
'------BEGIN MODIFICATIONS HOKANSON 12-AUG-2000
''''       frmIonExchangeMain!txtBedData(1).Text = Format$(nowproj.bed.Diameter, GetTheFormat(nowproj.bed.Diameter))
       frmIonExchangeMain!cboBedDataUnits(1).ListIndex = -1
'------END MODIFICATIONS HOKANSON 12-AUG-2000
       For i = 0 To frmIonExchangeMain!cboBedDataUnits(1).ListCount - 1
           If Trim$(frmIonExchangeMain!cboBedDataUnits(1).List(i)) = Trim$(UnitsOfDisplay) Then
              frmIonExchangeMain!cboBedDataUnits(1).ListIndex = i
              Exit For
           End If
       Next i

       Input #1, DummyString, NowProj.Bed.Weight, UnitsOfStorage, UnitsOfDisplay
'------BEGIN MODIFICATIONS HOKANSON 12-AUG-2000
''''       frmIonExchangeMain!txtBedData(2).Text = Format$(nowproj.bed.Weight, GetTheFormat(nowproj.bed.Weight))
       frmIonExchangeMain!cboBedDataUnits(2).ListIndex = -1
'------END MODIFICATIONS HOKANSON 12-AUG-2000
       For i = 0 To frmIonExchangeMain!cboBedDataUnits(2).ListCount - 1
           If Trim$(frmIonExchangeMain!cboBedDataUnits(2).List(i)) = Trim$(UnitsOfDisplay) Then
              frmIonExchangeMain!cboBedDataUnits(2).ListIndex = i
              Exit For
           End If
       Next i

       Input #1, DummyString, NowProj.Bed.Flowrate.Value, UnitsOfStorage, UnitsOfDisplay, NowProj.Bed.Flowrate.UserInput
'------BEGIN MODIFICATIONS HOKANSON 12-AUG-2000
''''       frmIonExchangeMain!txtBedData(3).Text = Format$(nowproj.bed.Flowrate.Value, GetTheFormat(nowproj.bed.Flowrate.Value))
       frmIonExchangeMain!cboBedDataUnits(3).ListIndex = -1
'------END MODIFICATIONS HOKANSON 12-AUG-2000
       For i = 0 To frmIonExchangeMain!cboBedDataUnits(3).ListCount - 1
           If Trim$(frmIonExchangeMain!cboBedDataUnits(3).List(i)) = Trim$(UnitsOfDisplay) Then
              frmIonExchangeMain!cboBedDataUnits(3).ListIndex = i
              Exit For
           End If
       Next i

       Input #1, DummyString, NowProj.Bed.EBCT.Value, UnitsOfStorage, UnitsOfDisplay, NowProj.Bed.EBCT.UserInput
'------BEGIN MODIFICATIONS HOKANSON 12-AUG-2000
''''       frmIonExchangeMain!txtBedData(4).Text = Format$(nowproj.bed.EBCT.Value, GetTheFormat(nowproj.bed.EBCT.Value))
       frmIonExchangeMain!cboBedDataUnits(4).ListIndex = -1
'------END MODIFICATIONS HOKANSON 12-AUG-2000
       For i = 0 To frmIonExchangeMain!cboBedDataUnits(4).ListCount - 1
           If Trim$(frmIonExchangeMain!cboBedDataUnits(4).List(i)) = Trim$(UnitsOfDisplay) Then
              frmIonExchangeMain!cboBedDataUnits(4).ListIndex = i
              Exit For
           End If
       Next i

       Input #1, DummyString, NowProj.Bed.NumberOfBeds


       'Input Adsorbent Properties

       Input #1, DummyString, NowProj.Resin.Name
       FoundResinInList = False
       For i = 0 To frmIonExchangeMain!cboAdsorbents.ListCount - 1
           If Trim$(frmIonExchangeMain!cboAdsorbents.List(i)) = Trim$(NowProj.Resin.Name) Then
              FoundResinInList = True
              ListIndexOfResin = i
              Exit For
           End If
       Next i
       If Not FoundResinInList Then
          frmIonExchangeMain!cboAdsorbents.AddItem Trim$(NowProj.Resin.Name)
          ListIndexOfResin = frmIonExchangeMain!cboAdsorbents.ListCount - 1
       End If

       Input #1, DummyString, NowProj.Resin.ApparentDensity, UnitsOfStorage, UnitsOfDisplay
'------BEGIN MODIFICATIONS HOKANSON 12-AUG-2000
''''       frmIonExchangeMain!txtAdsorbentProperties(1).Text = Format$(nowproj.resin.ApparentDensity, GetTheFormat(nowproj.resin.ApparentDensity))
       frmIonExchangeMain!cboAdsorbentPropertyUnits(1).ListIndex = -1
'------END MODIFICATIONS HOKANSON 12-AUG-2000
       For i = 0 To frmIonExchangeMain!cboAdsorbentPropertyUnits(1).ListCount - 1
           If Trim$(frmIonExchangeMain!cboAdsorbentPropertyUnits(1).List(i)) = Trim$(UnitsOfDisplay) Then
              frmIonExchangeMain!cboAdsorbentPropertyUnits(1).ListIndex = i
              Exit For
           End If
       Next i

       Input #1, DummyString, NowProj.Resin.ParticleRadius, UnitsOfStorage, UnitsOfDisplay
'------BEGIN MODIFICATIONS HOKANSON 12-AUG-2000
''''       frmIonExchangeMain!txtAdsorbentProperties(2).Text = Format$(nowproj.resin.ParticleRadius, GetTheFormat(nowproj.resin.ParticleRadius))
       frmIonExchangeMain!cboAdsorbentPropertyUnits(2).ListIndex = -1
'------END MODIFICATIONS HOKANSON 12-AUG-2000
       For i = 0 To frmIonExchangeMain!cboAdsorbentPropertyUnits(2).ListCount - 1
           If Trim$(frmIonExchangeMain!cboAdsorbentPropertyUnits(2).List(i)) = Trim$(UnitsOfDisplay) Then
              frmIonExchangeMain!cboAdsorbentPropertyUnits(2).ListIndex = i
              Exit For
           End If
       Next i

       Input #1, DummyString, NowProj.Resin.ParticlePorosity
       frmIonExchangeMain!txtAdsorbentProperties(3).Text = Format$(NowProj.Resin.ParticlePorosity, GetTheFormat(NowProj.Resin.ParticlePorosity))

       Input #1, DummyString, NowProj.Resin.Tortuosity
       frmIonExchangeMain!txtAdsorbentProperties(4).Text = Format$(NowProj.Resin.Tortuosity, GetTheFormat(NowProj.Resin.Tortuosity))

       Input #1, DummyString, NowProj.Resin.TotalCapacity, UnitsOfStorage, UnitsOfDisplay
'------BEGIN MODIFICATIONS HOKANSON 12-AUG-2000
''''       frmIonExchangeMain!txtAdsorbentProperties(5).Text = Format$(nowproj.resin.TotalCapacity, GetTheFormat(nowproj.resin.TotalCapacity))
       frmIonExchangeMain!cboAdsorbentPropertyUnits(5).ListIndex = -1
'------END MODIFICATIONS HOKANSON 12-AUG-2000
       For i = 0 To frmIonExchangeMain!cboAdsorbentPropertyUnits(5).ListCount - 1
           If Trim$(frmIonExchangeMain!cboAdsorbentPropertyUnits(5).List(i)) = Trim$(UnitsOfDisplay) Then
              frmIonExchangeMain!cboAdsorbentPropertyUnits(5).ListIndex = i
              Exit For
           End If
       Next i

       'Input Time Parameters
       Input #1, DummyString, NowProj.TimeParameters.FinalTime, UnitsOfStorage, UnitsOfDisplay
'------BEGIN MODIFICATIONS HOKANSON 12-AUG-2000
       frmOptionsInputParameters!cboTimeParametersUnits(0).ListIndex = -1
'------END MODIFICATIONS HOKANSON 12-AUG-2000
       For i = 0 To frmOptionsInputParameters!cboTimeParametersUnits(0).ListCount - 1
           If Trim$(frmOptionsInputParameters!cboTimeParametersUnits(0).List(i)) = Trim$(UnitsOfDisplay) Then
              frmOptionsInputParameters!cboTimeParametersUnits(0).ListIndex = i
              Exit For
           End If
       Next i

       Input #1, DummyString, NowProj.TimeParameters.InitialTime, UnitsOfStorage, UnitsOfDisplay
'------BEGIN MODIFICATIONS HOKANSON 12-AUG-2000
       frmOptionsInputParameters!cboTimeParametersUnits(1).ListIndex = -1
'------END MODIFICATIONS HOKANSON 12-AUG-2000
       For i = 0 To frmOptionsInputParameters!cboTimeParametersUnits(1).ListCount - 1
           If Trim$(frmOptionsInputParameters!cboTimeParametersUnits(1).List(i)) = Trim$(UnitsOfDisplay) Then
              frmOptionsInputParameters!cboTimeParametersUnits(1).ListIndex = i
              Exit For
           End If
       Next i

       Input #1, DummyString, NowProj.TimeParameters.TimeStep, UnitsOfStorage, UnitsOfDisplay
'------BEGIN MODIFICATIONS HOKANSON 12-AUG-2000
       frmOptionsInputParameters!cboTimeParametersUnits(2).ListIndex = -1
'------END MODIFICATIONS HOKANSON 12-AUG-2000
       For i = 0 To frmOptionsInputParameters!cboTimeParametersUnits(2).ListCount - 1
           If Trim$(frmOptionsInputParameters!cboTimeParametersUnits(2).List(i)) = Trim$(UnitsOfDisplay) Then
              frmOptionsInputParameters!cboTimeParametersUnits(2).ListIndex = i
              Exit For
           End If
       Next i

'------Begin Modification Hokanson: 11-Aug2000
       'Input EPS_ErrorCriteria and DH0_Initial Time step for DGEAR integrator
       If FileID <> "Ion Exchange Model - Input File" Then
          Input #1, DummyString, EPS_ErrorCriteriaForDGEARIntegrator
       End If

       If FileID <> "Ion Exchange Model - Input File" Then
          Input #1, DummyString, DH0_InitialTimeStepForDGEARIntegrator, UnitsOfStorage, UnitsOfDisplay
'---------BEGIN MODIFICATIONS HOKANSON 12-AUG-2000
          frmOptionsInputParameters!cboTimeParametersUnits(4).ListIndex = -1
'---------END MODIFICATIONS HOKANSON 12-AUG-2000
          For i = 0 To frmOptionsInputParameters!cboTimeParametersUnits(4).ListCount - 1
              If Trim$(frmOptionsInputParameters!cboTimeParametersUnits(4).List(i)) = Trim$(UnitsOfDisplay) Then
                 frmOptionsInputParameters!cboTimeParametersUnits(4).ListIndex = i
                 Exit For
              End If
          Next i
       End If
'------End Modification Hokanson: 11-Aug2000

       'Input Collocation Points
       Input #1, DummyString, NowProj.NumAxialCollocationPoints
       Input #1, DummyString, NowProj.NumRadialCollocationPoints

       'Correlation Used to Calculate Ionic Transport Coefficient
       Input #1, DummyString, NowProj.IonicTransportCoeffCorrName

       'Input Component Properties
       Input #1, DummyString
       Input #1, DummyString, NowProj.NumberOfCations
       Input #1, DummyString, NowProj.PresaturantCation
       Input #1, DummyString, NowProj.SumCationInitialEquivalents, NowProj.OKToGetCationDimensionless
       Input #1, DummyString, NowProj.CationSeparationFactorInput.Row, NowProj.CationSeparationFactorInput.Value
       For i = 1 To NowProj.NumberOfCations
           Input #1, NowProj.Cation(i).Name
           Input #1, NowProj.Cation(i).MolecularWeight, NowProj.Cation(i).InitialConcentration, NowProj.Cation(i).EquivalentInitialConcentration, NowProj.Cation(i).Valence, NowProj.Cation(i).SeparationFactor
           Input #1, NowProj.Cation(i).Kinetic.LiquidDiffusivity.Value, NowProj.Cation(i).Kinetic.LiquidDiffusivity.UserInput
           Input #1, NowProj.Cation(i).Kinetic.LiquidDiffusivityCorrelation, UnitsOfStorage
           Input #1, NowProj.Cation(i).Kinetic.LiquidDiffusivityUserInput, UnitsOfStorage
           Input #1, NowProj.Cation(i).Kinetic.IonicTransportCoefficient.Value, NowProj.Cation(i).Kinetic.IonicTransportCoefficient.UserInput
           Input #1, NowProj.Cation(i).Kinetic.IonicTransportCoeffCorrelation, UnitsOfStorage
           Input #1, NowProj.Cation(i).Kinetic.IonicTransportCoeffUserInput, UnitsOfStorage
           Input #1, NowProj.Cation(i).Kinetic.PoreDiffusivity.Value, NowProj.Cation(i).Kinetic.PoreDiffusivity.UserInput
           Input #1, NowProj.Cation(i).Kinetic.PoreDiffusivityCorrelation, UnitsOfStorage
           Input #1, NowProj.Cation(i).Kinetic.PoreDiffusivityUserInput, UnitsOfStorage
           Input #1, NowProj.Cation(i).Kinetic.NernstHaskellCation.Ion_Name, NowProj.Cation(i).Kinetic.NernstHaskellCation.Valence, NowProj.Cation(i).Kinetic.NernstHaskellCation.LimitingIonicConductance
           Input #1, NowProj.Cation(i).Kinetic.NernstHaskellAnion.Ion_Name, NowProj.Cation(i).Kinetic.NernstHaskellAnion.Valence, NowProj.Cation(i).Kinetic.NernstHaskellAnion.LimitingIonicConductance
       Next i

       Input #1, DummyString
       Input #1, DummyString, NowProj.NumberOfAnions
       Input #1, DummyString, NowProj.PresaturantAnion
       Input #1, DummyString, NowProj.SumAnionInitialEquivalents, NowProj.OKToGetAnionDimensionless
       Input #1, DummyString, NowProj.AnionSeparationFactorInput.Row, NowProj.AnionSeparationFactorInput.Value
       For i = 1 To NowProj.NumberOfAnions
           Input #1, NowProj.Anion(i).Name
           Input #1, NowProj.Anion(i).MolecularWeight, NowProj.Anion(i).InitialConcentration, NowProj.Anion(i).EquivalentInitialConcentration, NowProj.Anion(i).Valence, NowProj.Anion(i).SeparationFactor
           Input #1, NowProj.Anion(i).Kinetic.LiquidDiffusivity.Value, NowProj.Anion(i).Kinetic.LiquidDiffusivity.UserInput
           Input #1, NowProj.Anion(i).Kinetic.LiquidDiffusivityCorrelation, UnitsOfStorage
           Input #1, NowProj.Anion(i).Kinetic.LiquidDiffusivityUserInput, UnitsOfStorage
           Input #1, NowProj.Anion(i).Kinetic.IonicTransportCoefficient.Value, NowProj.Anion(i).Kinetic.IonicTransportCoefficient.UserInput
           Input #1, NowProj.Anion(i).Kinetic.IonicTransportCoeffCorrelation, UnitsOfStorage
           Input #1, NowProj.Anion(i).Kinetic.IonicTransportCoeffUserInput, UnitsOfStorage
           Input #1, NowProj.Anion(i).Kinetic.PoreDiffusivity.Value, NowProj.Anion(i).Kinetic.PoreDiffusivity.UserInput
           Input #1, NowProj.Anion(i).Kinetic.PoreDiffusivityCorrelation, UnitsOfStorage
           Input #1, NowProj.Anion(i).Kinetic.PoreDiffusivityUserInput, UnitsOfStorage
           Input #1, NowProj.Anion(i).Kinetic.NernstHaskellCation.Ion_Name, NowProj.Anion(i).Kinetic.NernstHaskellCation.Valence, NowProj.Anion(i).Kinetic.NernstHaskellCation.LimitingIonicConductance
           Input #1, NowProj.Anion(i).Kinetic.NernstHaskellAnion.Ion_Name, NowProj.Anion(i).Kinetic.NernstHaskellAnion.Valence, NowProj.Anion(i).Kinetic.NernstHaskellAnion.LimitingIonicConductance
       Next i


       'Place information on anions and cations in appropriate locations

       'In cboIons on main form

       'Cations
       frmIonExchangeMain!cboIons(0).Clear
       frmIonExchangeMain!lstIons(0).Clear
       For i = 1 To NowProj.NumberOfCations
           frmIonExchangeMain!cboIons(0).AddItem NowProj.Cation(i).Name
       Next i
'------BEGIN MODIFICATIONS HOKANSON 12-AUG-2000
       frmIonExchangeMain!cboIons(0).ListIndex = -1
'------END MODIFICATIONS HOKANSON 12-AUG-2000
       frmIonExchangeMain!cboIons(0).ListIndex = NowProj.PresaturantCation - 1

       'Anions
       frmIonExchangeMain!cboIons(1).Clear
       frmIonExchangeMain!lstIons(1).Clear
       For i = 1 To NowProj.NumberOfAnions
           frmIonExchangeMain!cboIons(1).AddItem NowProj.Anion(i).Name
       Next i
'------BEGIN MODIFICATIONS HOKANSON 12-AUG-2000
       frmIonExchangeMain!cboIons(1).ListIndex = -1
'------END MODIFICATIONS HOKANSON 12-AUG-2000
       frmIonExchangeMain!cboIons(1).ListIndex = NowProj.PresaturantAnion - 1

       'Generate click event on appropriate resin in cboAdsorbents
       frmIonExchangeMain!cboAdsorbents.ListIndex = -1
       frmIonExchangeMain!cboAdsorbents.ListIndex = ListIndexOfResin
       frmIonExchangeMain!fraKineticDimensionless.Enabled = True

       'Calculate Needed Properties
       Call CalculateLiquidDensity
       Call CalculateLiquidViscosity
       Call CalculateBedArea
       Call CalculateBedVolume
       Call CalculateBedDensity
       Call CalculateBedPorosity
       Call CalculateSuperficialVelocity
       Call CalculateInterstitialVelocity
       Call CalculateEffectiveContactTime
       Call CalculateParticleDiameter

       Call UpdateKineticParametersAllIons
'       Call UpdateDimensionlessGroupAllIons

       'Input units of display of properties related to components
       'Units of Display of Properties Related to Components

       'Molecular Weight
       Input #1, DummyString, UnitsOfDisplay
'------BEGIN MODIFICATIONS HOKANSON 12-AUG-2000
       frmAddComponent!cboAddIonUnits(0).ListIndex = -1
'------END MODIFICATIONS HOKANSON 12-AUG-2000
       For i = 0 To frmAddComponent!cboAddIonUnits(0).ListCount - 1
           If Trim$(frmAddComponent!cboAddIonUnits(0).List(i)) = Trim$(UnitsOfDisplay) Then
              frmAddComponent!cboAddIonUnits(0).ListIndex = i
              Exit For
           End If
       Next i

       'Initial Concentration
       Input #1, DummyString, UnitsOfDisplay
'------BEGIN MODIFICATIONS HOKANSON 12-AUG-2000
       frmAddComponent!cboAddIonUnits(1).ListIndex = -1
'------END MODIFICATIONS HOKANSON 12-AUG-2000
       For i = 0 To frmAddComponent!cboAddIonUnits(1).ListCount - 1
           If Trim$(frmAddComponent!cboAddIonUnits(1).List(i)) = Trim$(UnitsOfDisplay) Then
              frmAddComponent!cboAddIonUnits(1).ListIndex = i
              Exit For
           End If
       Next i

       'Liquid Diffusivity Correlation
       Input #1, DummyString, UnitsOfDisplay
'------BEGIN MODIFICATIONS HOKANSON 12-AUG-2000
       frmInputKineticParameters!cboLiquidDiffusivityUnits(0).ListIndex = -1
'------END MODIFICATIONS HOKANSON 12-AUG-2000
       For i = 0 To frmInputKineticParameters!cboLiquidDiffusivityUnits(0).ListCount - 1
           If Trim$(frmInputKineticParameters!cboLiquidDiffusivityUnits(0).List(i)) = Trim$(UnitsOfDisplay) Then
              frmInputKineticParameters!cboLiquidDiffusivityUnits(0).ListIndex = i
              Exit For
           End If
       Next i

       'Liquid Diffusivity User Input
       Input #1, DummyString, UnitsOfDisplay
'------BEGIN MODIFICATIONS HOKANSON 12-AUG-2000
       frmInputKineticParameters!cboLiquidDiffusivityUnits(1).ListIndex = -1
'------END MODIFICATIONS HOKANSON 12-AUG-2000
       For i = 0 To frmInputKineticParameters!cboLiquidDiffusivityUnits(1).ListCount - 1
           If Trim$(frmInputKineticParameters!cboLiquidDiffusivityUnits(1).List(i)) = Trim$(UnitsOfDisplay) Then
              frmInputKineticParameters!cboLiquidDiffusivityUnits(1).ListIndex = i
              Exit For
           End If
       Next i

       'Ionic Transport Coeff Correlation
       Input #1, DummyString, UnitsOfDisplay
'------BEGIN MODIFICATIONS HOKANSON 12-AUG-2000
       frmInputKineticParameters!cboIonicTransportUnits(0).ListIndex = -1
'------END MODIFICATIONS HOKANSON 12-AUG-2000
       For i = 0 To frmInputKineticParameters!cboIonicTransportUnits(0).ListCount - 1
           If Trim$(frmInputKineticParameters!cboIonicTransportUnits(0).List(i)) = Trim$(UnitsOfDisplay) Then
              frmInputKineticParameters!cboIonicTransportUnits(0).ListIndex = i
              Exit For
           End If
       Next i

       'Ionic Transport Coeff User Input
       Input #1, DummyString, UnitsOfDisplay
'------BEGIN MODIFICATIONS HOKANSON 12-AUG-2000
       frmInputKineticParameters!cboIonicTransportUnits(1).ListIndex = -1
'------END MODIFICATIONS HOKANSON 12-AUG-2000
       For i = 0 To frmInputKineticParameters!cboIonicTransportUnits(1).ListCount - 1
           If Trim$(frmInputKineticParameters!cboIonicTransportUnits(1).List(i)) = Trim$(UnitsOfDisplay) Then
              frmInputKineticParameters!cboIonicTransportUnits(1).ListIndex = i
              Exit For
           End If
       Next i

       'Pore Diffusivity Correlation
       Input #1, DummyString, UnitsOfDisplay
       For i = 0 To frmInputKineticParameters!cboPoreDiffusivityUnits(0).ListCount - 1
           If Trim$(frmInputKineticParameters!cboPoreDiffusivityUnits(0).List(i)) = Trim$(UnitsOfDisplay) Then
              frmInputKineticParameters!cboPoreDiffusivityUnits(0).ListIndex = i
              Exit For
           End If
       Next i

       'Pore Diffusivity User Input
       Input #1, DummyString, UnitsOfDisplay
'------BEGIN MODIFICATIONS HOKANSON 12-AUG-2000
       frmInputKineticParameters!cboPoreDiffusivityUnits(1).ListIndex = -1
'------END MODIFICATIONS HOKANSON 12-AUG-2000
       For i = 0 To frmInputKineticParameters!cboPoreDiffusivityUnits(1).ListCount - 1
           If Trim$(frmInputKineticParameters!cboPoreDiffusivityUnits(1).List(i)) = Trim$(UnitsOfDisplay) Then
              frmInputKineticParameters!cboPoreDiffusivityUnits(1).ListIndex = i
              Exit For
           End If
       Next i

       Input #1, DummyString, NowProj.VarInfluentFileCation
       Input #1, DummyString, NowProj.VarInfluentFileAnion

'       'Generate Click Event On frmIonExchangeMain!cboKinDimComponent
'       ListIndex = frmIonExchangeMain!cboKinDimComponent.ListIndex
'       frmIonExchangeMain!cboKinDimComponent.ListIndex = -1
'       frmIonExchangeMain!cboKinDimComponent.ListIndex = ListIndex

       If NowProj.IonicTransportCoeffCorrName = IONIC_TRANSPORT_COEFFICIENT_1 Then
          If frmInputKineticParameters!cboIonicTransport.ListIndex = 1 Then
             frmInputKineticParameters!cboIonicTransport.ListIndex = 0
          Else
             'Generate Click Event On frmInputKineticParameters!cboIon
             ListIndex = frmInputKineticParameters!cboIon.ListIndex
             frmInputKineticParameters!cboIon.ListIndex = -1
             frmInputKineticParameters!cboIon.ListIndex = ListIndex
          End If
       ElseIf NowProj.IonicTransportCoeffCorrName = IONIC_TRANSPORT_COEFFICIENT_2 Then
          If frmInputKineticParameters!cboIonicTransport.ListIndex = 0 Then
             frmInputKineticParameters!cboIonicTransport.ListIndex = 1
          Else
             'Generate Click Event On frmInputKineticParameters!cboIon
             ListIndex = frmInputKineticParameters!cboIon.ListIndex
             frmInputKineticParameters!cboIon.ListIndex = -1
             frmInputKineticParameters!cboIon.ListIndex = ListIndex
          End If
       End If

       Close #1

       Call ReadVarInfluentConcs

       frmIonExchangeMain.Caption = "Ion Exchange Design Software - " & filename
       
      Current_Filename = filename
      Call OldFileList_Promote( _
      Current_Filename, _
      1, _
      frmIonExchangeMain.mnuFile(13), _
      frmIonExchangeMain.mnuFile(8), _
      frmIonExchangeMain.mnuFile(9), _
      frmIonExchangeMain.mnuFile(10), _
      frmIonExchangeMain.mnuFile(11))
    
    End If

End Sub

Sub LoadNernstHaskellDatabases()
    Dim DB_Identifier As String
    Dim Ion_Name As String
    Dim Valence As Integer
    Dim LimitingIonicConductance As Double
    Dim NumberOfIons As Integer
    Dim msg As String
    Dim i As Integer

    'Load Anion Database
    Open IonExchangePath & "\exes\NHANION.TXT" For Input As #1
       Input #1, DB_Identifier
       If DB_Identifier <> NERNST_HASKELL_ANION_DB_IDENTIFIER Then
          msg = "There is a problem with the Nernst-Haskell anion database 'NHANION.TXT'.  "
          msg = "Check to make sure this file is included in the application directory."
          MsgBox msg, MB_ICONINFORMATION
          Close #1
          GoTo CationDatabase
       End If
       
       Input #1, NumberOfIons
       If NumberOfIons > MAX_NERNST_HASKELL_DB_IONS Then
          msg = "The number of ions currently in the Nernst-Haskell anion database (" & NumberOfIons & ") "
          msg = msg & "exceeds the maximum number of ions that can be stored in the "
          msg = msg & "program (" & MAX_NERNST_HASKELL_DB_IONS & ").  Only the first " & MAX_NERNST_HASKELL_DB_IONS & " "
          msg = msg & "ions will be loaded into the program."
          MsgBox msg, MB_ICONINFORMATION
          NumberOfIons = MAX_NERNST_HASKELL_DB_IONS
       End If
       If NumberOfIons <= 0 Then
          MsgBox "There are no ions currently in the Nernst-Haskell anion database."
          GoTo CationDatabase
       End If

       NernstHaskell.NumberOfAnionsInDB = NumberOfIons
       ReDim NernstHaskell.Anion(NumberOfIons)

       For i = 1 To NumberOfIons
''''           ReDim NernstHaskell.Anion(NumberOfIons)
           Input #1, Ion_Name, Valence, LimitingIonicConductance
           NernstHaskell.Anion(i).Ion_Name = Ion_Name
           NernstHaskell.Anion(i).Valence = Valence
           NernstHaskell.Anion(i).LimitingIonicConductance = LimitingIonicConductance
           frmAddComponent!cboAnion.List(i - 1) = Trim$(Ion_Name)
       Next i
    Close #1
    NumberOfIons = NumberOfIons
    
CationDatabase:

    'Load Cation Database
    Open IonExchangePath & "\exes\NHCATION.TXT" For Input As #1
       Input #1, DB_Identifier
       If DB_Identifier <> NERNST_HASKELL_CATION_DB_IDENTIFIER Then
          msg = "There is a problem with the Nernst-Haskell cation database 'NHCATION.TXT'.  "
          msg = "Check to make sure this file is included in the application directory."
          MsgBox msg, MB_ICONINFORMATION
          Close #1
          GoTo EndOfSub
       End If

       Input #1, NumberOfIons
       If NumberOfIons > MAX_NERNST_HASKELL_DB_IONS Then
          msg = "The number of ions currently in the Nernst-Haskell cation database (" & NumberOfIons & ") "
          msg = msg & "exceeds the maximum number of ions that can be stored in the "
          msg = msg & "program (" & MAX_NERNST_HASKELL_DB_IONS & ").  Only the first " & MAX_NERNST_HASKELL_DB_IONS & " "
          msg = msg & "ions will be loaded into the program."
          MsgBox msg, MB_ICONINFORMATION
          NumberOfIons = MAX_NERNST_HASKELL_DB_IONS
       End If
       If NumberOfIons <= 0 Then
          MsgBox "There are no ions currently in the Nernst-Haskell anion database."
          GoTo EndOfSub
       End If

       NernstHaskell.NumberOfCationsInDB = NumberOfIons
       ReDim NernstHaskell.Cation(NumberOfIons)
       For i = 1 To NumberOfIons
''''            ReDim NernstHaskell.Cation(NumberOfIons)
           Input #1, Ion_Name, Valence, LimitingIonicConductance
           NernstHaskell.Cation(i).Ion_Name = Ion_Name
           NernstHaskell.Cation(i).Valence = Valence
           NernstHaskell.Cation(i).LimitingIonicConductance = LimitingIonicConductance
           frmAddComponent!cboCation.List(i - 1) = Trim$(Ion_Name)
       Next i
       NernstHaskell.DefaultAnion = NernstHaskell.Anion(2)
       NernstHaskell.DefaultCation = NernstHaskell.Cation(3)

    Close #1
    NumberOfIons = NumberOfIons
    
EndOfSub:

End Sub

Sub LoadUnitsAddIon()

    'Load Possible Choices into the Add Ion Units Combo Boxes
    'on frmAddComponent

    'Molecular Weight Units
    frmAddComponent!cboAddIonUnits(0).List(0) = "mg/mmol"
    frmAddComponent!cboAddIonUnits(0).List(1) = Chr$(181) & "g/" & Chr$(181) & "mol"
    frmAddComponent!cboAddIonUnits(0).List(2) = "g/gmol"
    frmAddComponent!cboAddIonUnits(0).List(3) = "kg/kmol"
    frmAddComponent!cboAddIonUnits(0).ListIndex = 0

    'Initial Concentration Units
    frmAddComponent!cboAddIonUnits(1).List(0) = "mg/L"
    frmAddComponent!cboAddIonUnits(1).List(1) = Chr$(181) & "g/L"
    frmAddComponent!cboAddIonUnits(1).List(2) = "g/L"
    frmAddComponent!cboAddIonUnits(1).List(3) = "meq/L"
    frmAddComponent!cboAddIonUnits(1).List(4) = "eq/L"
    frmAddComponent!cboAddIonUnits(1).List(5) = "mmol/L"
    frmAddComponent!cboAddIonUnits(1).List(6) = Chr$(181) & "mol/L"
    frmAddComponent!cboAddIonUnits(1).List(7) = "gmol/L"
    frmAddComponent!cboAddIonUnits(1).ListIndex = 0

End Sub

Sub LoadUnitsAdsorbentProperties()

    'Load Possible Choices into the Adsorbent Properties Units Combo Boxes

    frmIonExchangeMain!cboAdsorbentPropertyUnits(1).List(0) = "g/ml"
    frmIonExchangeMain!cboAdsorbentPropertyUnits(1).List(1) = "kg/m" & Chr$(179)
    frmIonExchangeMain!cboAdsorbentPropertyUnits(1).List(2) = "lb/ft" & Chr$(179)
    frmIonExchangeMain!cboAdsorbentPropertyUnits(1).List(3) = "lb/gal"
    frmIonExchangeMain!cboAdsorbentPropertyUnits(1).ListIndex = 0

    frmIonExchangeMain!cboAdsorbentPropertyUnits(2).List(0) = "m"
    frmIonExchangeMain!cboAdsorbentPropertyUnits(2).List(1) = "cm"
    frmIonExchangeMain!cboAdsorbentPropertyUnits(2).List(2) = "ft"
    frmIonExchangeMain!cboAdsorbentPropertyUnits(2).List(3) = "in"
    frmIonExchangeMain!cboAdsorbentPropertyUnits(2).ListIndex = 0

    frmIonExchangeMain!cboAdsorbentPropertyUnits(5).List(0) = "meq/g"
    frmIonExchangeMain!cboAdsorbentPropertyUnits(5).List(1) = "meq/ml bed"
    frmIonExchangeMain!cboAdsorbentPropertyUnits(5).List(2) = "meq/ml resin"
    frmIonExchangeMain!cboAdsorbentPropertyUnits(5).ListIndex = 0
    

End Sub

Sub LoadUnitsBedData()

    'Load Possible Choices into the Bed Value Units Combo Boxes
    frmIonExchangeMain!cboBedDataUnits(0).List(0) = "m"
    frmIonExchangeMain!cboBedDataUnits(0).List(1) = "cm"
    frmIonExchangeMain!cboBedDataUnits(0).List(2) = "ft"
    frmIonExchangeMain!cboBedDataUnits(0).List(3) = "in"
    frmIonExchangeMain!cboBedDataUnits(0).ListIndex = 0

    frmIonExchangeMain!cboBedDataUnits(1).List(0) = "m"
    frmIonExchangeMain!cboBedDataUnits(1).List(1) = "cm"
    frmIonExchangeMain!cboBedDataUnits(1).List(2) = "ft"
    frmIonExchangeMain!cboBedDataUnits(1).List(3) = "in"
    frmIonExchangeMain!cboBedDataUnits(1).ListIndex = 0

    frmIonExchangeMain!cboBedDataUnits(2).List(0) = "kg"
    frmIonExchangeMain!cboBedDataUnits(2).List(1) = "g"
    frmIonExchangeMain!cboBedDataUnits(2).List(2) = "lb"
    frmIonExchangeMain!cboBedDataUnits(2).ListIndex = 0

    frmIonExchangeMain!cboBedDataUnits(3).List(0) = "m" & Chr$(179) & "/s"
    frmIonExchangeMain!cboBedDataUnits(3).List(1) = "m" & Chr$(179) & "/d"
    frmIonExchangeMain!cboBedDataUnits(3).List(2) = "cm" & Chr$(179) & "/s"
    frmIonExchangeMain!cboBedDataUnits(3).List(3) = "ml/min"
    frmIonExchangeMain!cboBedDataUnits(3).List(4) = "ft" & Chr$(179) & "/s"
    frmIonExchangeMain!cboBedDataUnits(3).List(5) = "ft" & Chr$(179) & "/d"
    frmIonExchangeMain!cboBedDataUnits(3).List(6) = "gpm"
    frmIonExchangeMain!cboBedDataUnits(3).List(7) = "gpd"
    frmIonExchangeMain!cboBedDataUnits(3).List(8) = "MGD"
    frmIonExchangeMain!cboBedDataUnits(3).ListIndex = 0

    frmIonExchangeMain!cboBedDataUnits(4).List(0) = "s"
    frmIonExchangeMain!cboBedDataUnits(4).List(1) = "min"
    frmIonExchangeMain!cboBedDataUnits(4).List(2) = "hr"
    frmIonExchangeMain!cboBedDataUnits(4).List(3) = "d"
    frmIonExchangeMain!cboBedDataUnits(4).ListIndex = 0

End Sub

Sub LoadUnitsKineticParameters()
    Dim i As Integer

    NowProj.IonicTransportCoeffCorrName = IONIC_TRANSPORT_COEFFICIENT_1

    'Load units on frmInputKineticParameters

    'First Liquid Diffusivity Combo Box
    frmInputKineticParameters!cboLiquidDiffusivityUnits(0).AddItem "cm" & Chr$(178) & "/s"
    frmInputKineticParameters!cboLiquidDiffusivityUnits(0).AddItem "cm" & Chr$(178) & "/min"
    frmInputKineticParameters!cboLiquidDiffusivityUnits(0).AddItem "m" & Chr$(178) & "/s"
    frmInputKineticParameters!cboLiquidDiffusivityUnits(0).AddItem "m" & Chr$(178) & "/min"
    frmInputKineticParameters!cboLiquidDiffusivityUnits(0).AddItem "m" & Chr$(178) & "/hr"
    frmInputKineticParameters!cboLiquidDiffusivityUnits(0).AddItem "m" & Chr$(178) & "/d"
    frmInputKineticParameters!cboLiquidDiffusivityUnits(0).AddItem "ft" & Chr$(178) & "/s"
    frmInputKineticParameters!cboLiquidDiffusivityUnits(0).AddItem "ft" & Chr$(178) & "/min"
    frmInputKineticParameters!cboLiquidDiffusivityUnits(0).AddItem "ft" & Chr$(178) & "/hr"
    frmInputKineticParameters!cboLiquidDiffusivityUnits(0).AddItem "ft" & Chr$(178) & "/d"


    'Second Liquid Diffusivity Combo Box and Pore Diffusivity Combo Boxes
    For i = 0 To frmInputKineticParameters!cboLiquidDiffusivityUnits(0).ListCount - 1
        frmInputKineticParameters!cboLiquidDiffusivityUnits(1).AddItem frmInputKineticParameters!cboLiquidDiffusivityUnits(0).List(i)
        frmInputKineticParameters!cboPoreDiffusivityUnits(0).AddItem frmInputKineticParameters!cboLiquidDiffusivityUnits(0).List(i)
        frmInputKineticParameters!cboPoreDiffusivityUnits(1).AddItem frmInputKineticParameters!cboLiquidDiffusivityUnits(0).List(i)
    Next i

    'First Ionic Transport Coefficient Combo Box
    frmInputKineticParameters!cboIonicTransportUnits(0).AddItem "cm/s"
    frmInputKineticParameters!cboIonicTransportUnits(0).AddItem "cm/min"
    frmInputKineticParameters!cboIonicTransportUnits(0).AddItem "m/s"
    frmInputKineticParameters!cboIonicTransportUnits(0).AddItem "m/min"
    frmInputKineticParameters!cboIonicTransportUnits(0).AddItem "m/hr"
    frmInputKineticParameters!cboIonicTransportUnits(0).AddItem "m/d"
    frmInputKineticParameters!cboIonicTransportUnits(0).AddItem "ft/s"
    frmInputKineticParameters!cboIonicTransportUnits(0).AddItem "ft/min"
    frmInputKineticParameters!cboIonicTransportUnits(0).AddItem "ft/hr"
    frmInputKineticParameters!cboIonicTransportUnits(0).AddItem "ft/d"

    'Second Ionic Transport Coefficient Combo Box
    For i = 0 To frmInputKineticParameters!cboLiquidDiffusivityUnits(0).ListCount - 1
        frmInputKineticParameters!cboIonicTransportUnits(1).AddItem frmInputKineticParameters!cboIonicTransportUnits(0).List(i)
    Next i

    frmInputKineticParameters!cboLiquidDiffusivityUnits(0).ListIndex = 0
    frmInputKineticParameters!cboLiquidDiffusivityUnits(1).ListIndex = 0

    frmInputKineticParameters!cboIonicTransportUnits(0).ListIndex = 0
    frmInputKineticParameters!cboIonicTransportUnits(1).ListIndex = 0

    frmInputKineticParameters!cboPoreDiffusivityUnits(0).ListIndex = 0
    frmInputKineticParameters!cboPoreDiffusivityUnits(1).ListIndex = 0

    'kf, Dl, and Dp units on main form
    frmIonExchangeMain!lblKineticDimensionlessUnits(0).Caption = "cm/s"
    frmIonExchangeMain!lblKineticDimensionlessUnits(1).Caption = "cm" & Chr$(178) & "/s"
    frmIonExchangeMain!lblKineticDimensionlessUnits(2).Caption = "cm" & Chr$(178) & "/s"

End Sub

Sub LoadUnitsOperatingConditions()

    'Load Possible Choices into the Operating Condition Units Combo Boxes
    frmIonExchangeMain!cboOperatingConditionsUnits(0).List(0) = "Pa"
    frmIonExchangeMain!cboOperatingConditionsUnits(0).List(1) = "kPa"
    frmIonExchangeMain!cboOperatingConditionsUnits(0).List(2) = "bars"
    frmIonExchangeMain!cboOperatingConditionsUnits(0).List(3) = "atm"
    frmIonExchangeMain!cboOperatingConditionsUnits(0).List(4) = "psi"
    frmIonExchangeMain!cboOperatingConditionsUnits(0).List(5) = "mm Hg"
    frmIonExchangeMain!cboOperatingConditionsUnits(0).List(6) = "m H2O"
    frmIonExchangeMain!cboOperatingConditionsUnits(0).List(7) = "ft H2O"
    frmIonExchangeMain!cboOperatingConditionsUnits(0).List(8) = "in H2O"
    frmIonExchangeMain!cboOperatingConditionsUnits(0).ListIndex = 0

    frmIonExchangeMain!cboOperatingConditionsUnits(1).List(0) = "K"
    frmIonExchangeMain!cboOperatingConditionsUnits(1).List(1) = "C"
    frmIonExchangeMain!cboOperatingConditionsUnits(1).List(2) = "R"
    frmIonExchangeMain!cboOperatingConditionsUnits(1).List(3) = "F"
    frmIonExchangeMain!cboOperatingConditionsUnits(1).ListIndex = 0

End Sub

Sub LoadUnitsTimeParameters()

    frmOptionsInputParameters!cboTimeParametersUnits(0).List(0) = "min"
    frmOptionsInputParameters!cboTimeParametersUnits(0).List(1) = "s"
    frmOptionsInputParameters!cboTimeParametersUnits(0).List(2) = "hr"
    frmOptionsInputParameters!cboTimeParametersUnits(0).List(3) = "d"
    frmOptionsInputParameters!cboTimeParametersUnits(0).ListIndex = 0

    frmOptionsInputParameters!cboTimeParametersUnits(1).List(0) = "min"
    frmOptionsInputParameters!cboTimeParametersUnits(1).List(1) = "s"
    frmOptionsInputParameters!cboTimeParametersUnits(1).List(2) = "hr"
    frmOptionsInputParameters!cboTimeParametersUnits(1).List(3) = "d"
    frmOptionsInputParameters!cboTimeParametersUnits(1).ListIndex = 0

    frmOptionsInputParameters!cboTimeParametersUnits(2).List(0) = "min"
    frmOptionsInputParameters!cboTimeParametersUnits(2).List(1) = "s"
    frmOptionsInputParameters!cboTimeParametersUnits(2).List(2) = "hr"
    frmOptionsInputParameters!cboTimeParametersUnits(2).List(3) = "d"
    frmOptionsInputParameters!cboTimeParametersUnits(2).ListIndex = 0

'------Begin Modification Hokanson: 11-Aug2000
    frmOptionsInputParameters!cboTimeParametersUnits(4).List(0) = "min"
    frmOptionsInputParameters!cboTimeParametersUnits(4).List(1) = "s"
    frmOptionsInputParameters!cboTimeParametersUnits(4).List(2) = "hr"
    frmOptionsInputParameters!cboTimeParametersUnits(4).List(3) = "d"
    frmOptionsInputParameters!cboTimeParametersUnits(4).ListIndex = 0
'------End Modification Hokanson: 11-Aug2000

End Sub

Sub NumberChanged(ValueChanged As Integer, OldValue As Double, NewValue As Double)
    Dim Dummy1 As Double, Dummy2 As Double

    ValueChanged = True
    If Abs(OldValue - NewValue) < TOLERANCE Then ValueChanged = False

End Sub

Sub NumberCheck(KeyAscii As Integer)

    If (KeyAscii > Asc("9") Or KeyAscii < Asc("0")) And KeyAscii <> Asc(".") And KeyAscii <> 8 And KeyAscii <> Asc("E") And KeyAscii <> Asc("e") And KeyAscii <> Asc("-") Then
       KeyAscii = 0
       Beep
    End If

End Sub

Sub PrintIonExchange()
    Dim msg As String
    Dim i As Integer, j As Integer
    Dim Ion_Name As String
    Dim MolWt As Double, MolWtToDisplay As String, MolWtUnits As String
    Dim Valence As Double, ValenceToDisplay As String
    Dim InitialConc As Double, InitialConcToDisplay As String, InitialConcUnits As String
    Dim CurrentUnits As Integer
    Dim LiqDiff As Double, LiqDiffToDisplay As String, LiqDiffUnits As Integer, LiqDiffUnitsOfDisplay As String
    Dim IonicTrans As Double, IonicTransToDisplay As String, IonicTransUnits As Integer, IonicTransUnitsOfDisplay As String
    Dim PoreDiff As Double, PoreDiffToDisplay As String, PoreDiffUnits As Integer, PoreDiffUnitsOfDisplay As String
    Dim Dgs As Double, DgsToDisplay As String
    Dim Dgp As Double, DgpToDisplay As String
    Dim Dgt As Double, DgtToDisplay As String
    Dim Edp As Double, EdpToDisplay As String
    Dim St As Double, StToDisplay As String
    Dim Bip As Double, BipToDisplay As String
    Dim PrintString As String
    Dim TabHeader As Integer
    
    On Error GoTo ErrorWhilePrinting

    Printer.ScaleLeft = -1440
    Printer.ScaleTop = -1440
    Printer.CurrentX = 0
    Printer.CurrentY = 0
    Printer.FontName = PRINTER_FONT
    Printer.FontSize = 14
    Printer.FontBold = True
    Printer.Print "Ion Exchange Simulation Software"
    Printer.Print
    Printer.Print
    Printer.FontUnderline = True
    Printer.FontSize = 12
    Printer.Print "Operating Conditons:"
    Printer.Print
    Printer.FontUnderline = False
    Printer.FontBold = False
    Printer.FontSize = PRINT_FONTSIZE_DATA
    Printer.Print "Pressure:"; Tab(TAB_OPERATING_CONDITIONS); Trim$(frmIonExchangeMain!txtOperatingConditions(0).Text); " "; frmIonExchangeMain!cboOperatingConditionsUnits(0).Text
    Printer.Print "Temperature:"; Tab(TAB_OPERATING_CONDITIONS); Trim$(frmIonExchangeMain!txtOperatingConditions(1).Text); " "; frmIonExchangeMain!cboOperatingConditionsUnits(1).Text
    Printer.Print
    Printer.Print
    Printer.FontBold = True
    Printer.FontUnderline = True
    Printer.FontSize = 12
    Printer.Print "Bed Data:"
    Printer.Print
    Printer.FontUnderline = False
    Printer.FontBold = False
    Printer.FontSize = PRINT_FONTSIZE_DATA
    Printer.Print "Bed Length:"; Tab(TAB_BED_DATA); Trim$(frmIonExchangeMain!txtBedData(0).Text); " "; frmIonExchangeMain!cboBedDataUnits(0).Text
    Printer.Print "Bed Diameter:"; Tab(TAB_BED_DATA); Trim$(frmIonExchangeMain!txtBedData(1).Text); " "; frmIonExchangeMain!cboBedDataUnits(1).Text
    Printer.Print "Mass of Resin:"; Tab(TAB_BED_DATA); Trim$(frmIonExchangeMain!txtBedData(2).Text); " "; frmIonExchangeMain!cboBedDataUnits(2).Text
    Printer.Print "Inlet Flowrate:"; Tab(TAB_BED_DATA); Trim$(frmIonExchangeMain!txtBedData(3).Text); " "; frmIonExchangeMain!cboBedDataUnits(3).Text
    Printer.Print "EBCT:"; Tab(TAB_BED_DATA); Trim$(frmIonExchangeMain!txtBedData(4).Text); " "; frmIonExchangeMain!cboBedDataUnits(4).Text
    Printer.Print "Bed Porosity:"; Tab(TAB_BED_DATA); Format$(NowProj.Bed.Porosity, GetTheFormat(NowProj.Bed.Porosity)); " "; "(-)"
    Printer.Print
    Printer.Print
    Printer.FontBold = True
    Printer.FontUnderline = True
    Printer.FontSize = 12
    Printer.Print "Resin Properties:"
    Printer.Print
    Printer.FontUnderline = False
    Printer.FontBold = False
    Printer.FontSize = PRINT_FONTSIZE_DATA
    Printer.Print "Name:"; Tab(TAB_RESIN_PROPERTIES); Trim$(frmIonExchangeMain!cboAdsorbents.Text)
    Printer.Print "Apparent Density:"; Tab(TAB_RESIN_PROPERTIES); Trim$(frmIonExchangeMain!txtAdsorbentProperties(1).Text); " "; frmIonExchangeMain!cboAdsorbentPropertyUnits(1).Text
    Printer.Print "Particle Radius:"; Tab(TAB_RESIN_PROPERTIES); Trim$(frmIonExchangeMain!txtAdsorbentProperties(2).Text); " "; frmIonExchangeMain!cboAdsorbentPropertyUnits(2).Text
    Printer.Print "Particle Porosity:"; Tab(TAB_RESIN_PROPERTIES); Trim$(frmIonExchangeMain!txtAdsorbentProperties(3).Text); " "; "(-)"
    Printer.Print "Tortuosity:"; Tab(TAB_RESIN_PROPERTIES); Trim$(frmIonExchangeMain!txtAdsorbentProperties(4).Text); " "; "(-)"
    Printer.Print "Total Capacity:"; Tab(TAB_RESIN_PROPERTIES); Trim$(frmIonExchangeMain!txtAdsorbentProperties(5).Text); " "; frmIonExchangeMain!cboAdsorbentPropertyUnits(5).Text

    If (NowProj.NumberOfCations <= 0) And (NowProj.NumberOfAnions <= 0) Then GoTo EndDocument

    Printer.Print
    Printer.Print
    Printer.FontBold = True
    Printer.FontUnderline = True
    Printer.FontSize = 12
    Printer.Print "Component Properties:"
    Printer.Print
    Printer.FontBold = False
    Printer.FontUnderline = False
    Printer.FontSize = PRINT_FONTSIZE_DATA
    
    If Cations.Available Then
       Printer.FontItalic = True
       Printer.Print "Cations"
       Printer.FontItalic = False
       Printer.Print
       Printer.Print Tab(TAB_COMPONENT_PROPERTIES_1); "Molecular Wt."; Tab(TAB_COMPONENT_PROPERTIES_2); "Initial Conc."; Tab(TAB_COMPONENT_PROPERTIES_3); "Valence:"
       MolWtUnits = frmAddComponent!cboAddIonUnits(0).List(frmAddComponent!cboAddIonUnits(0).ListIndex)
       InitialConcUnits = frmAddComponent!cboAddIonUnits(1).List(frmAddComponent!cboAddIonUnits(1).ListIndex)
       Printer.Print "Name:"; Tab(TAB_COMPONENT_PROPERTIES_1); "("; MolWtUnits; ")"; Tab(TAB_COMPONENT_PROPERTIES_2); "("; InitialConcUnits; ")"; Tab(TAB_COMPONENT_PROPERTIES_3); "(-)"
       Printer.Print
    
       For i = 1 To NowProj.NumberOfCations
           Ion_Name = NowProj.Cation(i).Name
           If Len(Ion_Name) > (TAB_COMPONENT_PROPERTIES_1 - 2) Then
              Ion_Name = left$(Ion_Name, TAB_COMPONENT_PROPERTIES_1 - 2)
           End If
           MolWt = NowProj.Cation(i).MolecularWeight
           MolWtToDisplay = Format$(MolWt, GetTheFormat(MolWt))
           Valence = NowProj.Cation(i).Valence
           ValenceToDisplay = Format$(Valence, "0")
           CurrentUnits = frmAddComponent!cboAddIonUnits(1).ListIndex
           InitialConc = NowProj.Cation(i).InitialConcentration
           If CurrentUnits <> 0 Then
              InitialConc = InitialConc * ConcentrationConversionFactor(CurrentUnits, Valence, MolWt)
           End If
           InitialConcToDisplay = Format$(InitialConc, GetTheFormat(InitialConc))
           Printer.Print Ion_Name; Tab(TAB_COMPONENT_PROPERTIES_1); MolWtToDisplay; Tab(TAB_COMPONENT_PROPERTIES_2); InitialConcToDisplay; Tab(TAB_COMPONENT_PROPERTIES_3); ValenceToDisplay
       Next i
       Printer.Print
    End If

    If Anions.Available Then
       Printer.FontItalic = True
       Printer.Print "Anions"
       Printer.FontItalic = False
       Printer.Print
       Printer.Print Tab(TAB_COMPONENT_PROPERTIES_1); "Molecular Wt."; Tab(TAB_COMPONENT_PROPERTIES_2); "Initial Conc."; Tab(TAB_COMPONENT_PROPERTIES_3); "Valence:"
       MolWtUnits = frmAddComponent!cboAddIonUnits(0).List(frmAddComponent!cboAddIonUnits(0).ListIndex)
       InitialConcUnits = frmAddComponent!cboAddIonUnits(1).List(frmAddComponent!cboAddIonUnits(1).ListIndex)
       Printer.Print "Name:"; Tab(TAB_COMPONENT_PROPERTIES_1); "("; MolWtUnits; ")"; Tab(TAB_COMPONENT_PROPERTIES_2); "("; InitialConcUnits; ")"; Tab(TAB_COMPONENT_PROPERTIES_3); "(-)"
       Printer.Print

       For i = 1 To NowProj.NumberOfAnions
           Ion_Name = NowProj.Anion(i).Name
           If Len(Ion_Name) > (TAB_COMPONENT_PROPERTIES_1 - 2) Then
              Ion_Name = left$(Ion_Name, TAB_COMPONENT_PROPERTIES_1 - 2)
           End If
           MolWt = NowProj.Anion(i).MolecularWeight
           MolWtToDisplay = Format$(MolWt, GetTheFormat(MolWt))
           Valence = NowProj.Anion(i).Valence
           ValenceToDisplay = Format$(Valence, "0")
           CurrentUnits = frmAddComponent!cboAddIonUnits(1).ListIndex
           InitialConc = NowProj.Anion(i).InitialConcentration
           If CurrentUnits <> 0 Then
              InitialConc = InitialConc * ConcentrationConversionFactor(CurrentUnits, Valence, MolWt)
           End If
           InitialConcToDisplay = Format$(InitialConc, GetTheFormat(InitialConc))
           Printer.Print Ion_Name; Tab(TAB_COMPONENT_PROPERTIES_1); MolWtToDisplay; Tab(TAB_COMPONENT_PROPERTIES_2); InitialConcToDisplay; Tab(TAB_COMPONENT_PROPERTIES_3); ValenceToDisplay
       Next i
       Printer.Print
    End If

    Printer.Print
    Printer.FontBold = True
    Printer.FontUnderline = True
    Printer.FontSize = 12
    Printer.Print "Separation Factors:"
    Printer.Print
    Printer.FontBold = False
    Printer.FontUnderline = False
    Printer.FontSize = PRINT_FONTSIZE_DATA

    If Cations.Available Then
       'Calculate Cation Separation Factors
       For i = 1 To NowProj.NumberOfCations
           OneDimSeparationFactors(i) = NowProj.Cation(i).SeparationFactor
       Next i
       NumberOfIons = NowProj.NumberOfCations
       SeparationFactorInput.Row = NowProj.CationSeparationFactorInput.Row
       SeparationFactorInput.Value = NowProj.CationSeparationFactorInput.Value
       Call CalculateSeparationFactors

       PrintString = "Alpha (i,j)"
       TabHeader = TAB_SEPARATION_FACTORS_1 + (((TAB_SEPARATION_FACTORS_1 + NowProj.NumberOfCations * TAB_SEPARATION_FACTORS_INTERVAL) - TAB_SEPARATION_FACTORS_1) / 2 - Len(PrintString) / 2)
       Printer.Print Tab(TabHeader); PrintString
       
       TabHeader = TabHeader + Len(PrintString) / 2
       Printer.Print Tab(TabHeader); "i"
       Printer.Print "j";

       For i = 1 To NowProj.NumberOfCations
           Printer.Print Tab(TAB_SEPARATION_FACTORS_1 + (i - 1) * TAB_SEPARATION_FACTORS_INTERVAL); Trim$(Str$(i));
       Next i
       Printer.Print
       For j = 1 To NowProj.NumberOfCations
           Printer.Print Trim$(Str$(j));
           For i = 1 To NowProj.NumberOfCations
               Printer.Print Tab(TAB_SEPARATION_FACTORS_1 + (i - 1) * TAB_SEPARATION_FACTORS_INTERVAL); Format$(TwoDimSeparationFactors(i, j), GetTheFormat(TwoDimSeparationFactors(i, j)));
           Next i
           Printer.Print
       Next j
       Printer.Print
       For i = 1 To NowProj.NumberOfCations
           Printer.Print Trim$(Str$(i)); " ="; Tab(7); Trim$(NowProj.Cation(i).Name)
       Next i
       Printer.Print
    End If
    Printer.Print
    
    If Anions.Available Then
       'Calculate Anion Separation Factors
       For i = 1 To NowProj.NumberOfAnions
           OneDimSeparationFactors(i) = NowProj.Anion(i).SeparationFactor
       Next i
       NumberOfIons = NowProj.NumberOfAnions
       SeparationFactorInput.Row = NowProj.AnionSeparationFactorInput.Row
       SeparationFactorInput.Value = NowProj.AnionSeparationFactorInput.Value
       Call CalculateSeparationFactors

       PrintString = "Alpha (i,j)"
       TabHeader = TAB_SEPARATION_FACTORS_1 + (((TAB_SEPARATION_FACTORS_1 + NowProj.NumberOfAnions * TAB_SEPARATION_FACTORS_INTERVAL) - TAB_SEPARATION_FACTORS_1) / 2 - Len(PrintString) / 2)
       Printer.Print Tab(TabHeader); PrintString
       
       TabHeader = TabHeader + Len(PrintString) / 2
       Printer.Print Tab(TabHeader); "i"
       Printer.Print "j";

       For i = 1 To NowProj.NumberOfAnions
           Printer.Print Tab(TAB_SEPARATION_FACTORS_1 + (i - 1) * TAB_SEPARATION_FACTORS_INTERVAL); Trim$(Str$(i));
       Next i
       Printer.Print
       For j = 1 To NowProj.NumberOfAnions
           Printer.Print Trim$(Str$(j));
           For i = 1 To NowProj.NumberOfAnions
               Printer.Print Tab(TAB_SEPARATION_FACTORS_1 + (i - 1) * TAB_SEPARATION_FACTORS_INTERVAL); Format$(TwoDimSeparationFactors(i, j), GetTheFormat(TwoDimSeparationFactors(i, j)));
           Next i
           Printer.Print
       Next j
       Printer.Print
       For i = 1 To NowProj.NumberOfAnions
           Printer.Print Trim$(Str$(i)); " ="; Tab(7); Trim$(NowProj.Anion(i).Name)
       Next i
       Printer.Print
    End If

    Printer.Print
    Printer.FontBold = True
    Printer.FontUnderline = True
    Printer.FontSize = 12
    Printer.Print "Kinetic Parameters:"
    Printer.Print
    Printer.FontBold = False
    Printer.FontUnderline = False
    Printer.FontSize = PRINT_FONTSIZE_DATA

    If Cations.Available Then
       Printer.Print Tab(TAB_KINETIC_PARAMETERS_1); "Ion Trans. - kf"; Tab(TAB_KINETIC_PARAMETERS_2); "Liq. Dif. - Dl"; Tab(TAB_KINETIC_PARAMETERS_3); "Pore Dif. - Dp"

       LiqDiffUnitsOfDisplay = "cm" & Chr$(178) & "/s"
       IonicTransUnitsOfDisplay = "cm/s"
       PoreDiffUnitsOfDisplay = "cm" & Chr$(178) & "/s"
       Printer.Print "Name:"; Tab(TAB_KINETIC_PARAMETERS_1); "("; IonicTransUnitsOfDisplay; ")"; Tab(TAB_KINETIC_PARAMETERS_2); "("; LiqDiffUnitsOfDisplay; ")"; Tab(TAB_KINETIC_PARAMETERS_3); PoreDiffUnitsOfDisplay
       Printer.Print
    
       For i = 1 To NowProj.NumberOfCations
           Ion_Name = NowProj.Cation(i).Name
           If Len(Ion_Name) > (TAB_KINETIC_PARAMETERS_1 - 2) Then
              Ion_Name = left$(Ion_Name, TAB_KINETIC_PARAMETERS_1 - 2)
           End If
           LiqDiff = NowProj.Cation(i).Kinetic.LiquidDiffusivity.Value
           LiqDiffToDisplay = Format$(LiqDiff, GetTheFormat(LiqDiff))
           IonicTrans = NowProj.Cation(i).Kinetic.IonicTransportCoefficient.Value
           IonicTransToDisplay = Format$(IonicTrans, GetTheFormat(IonicTrans))
           PoreDiff = NowProj.Cation(i).Kinetic.PoreDiffusivity.Value
           PoreDiffToDisplay = Format$(PoreDiff, GetTheFormat(PoreDiff))
           Printer.Print Ion_Name; Tab(TAB_KINETIC_PARAMETERS_1); IonicTransToDisplay; Tab(TAB_KINETIC_PARAMETERS_2); LiqDiffToDisplay; Tab(TAB_KINETIC_PARAMETERS_3); PoreDiffToDisplay
       Next i
       Printer.Print
    End If

    If Anions.Available Then
       Printer.Print Tab(TAB_KINETIC_PARAMETERS_1); "Ion Trans. - kf"; Tab(TAB_KINETIC_PARAMETERS_2); "Liq. Dif. - Dl"; Tab(TAB_KINETIC_PARAMETERS_3); "Pore Dif. - Dp"

       LiqDiffUnitsOfDisplay = "cm" & Chr$(178) & "/s"
       IonicTransUnitsOfDisplay = "cm/s"
       PoreDiffUnitsOfDisplay = "cm" & Chr$(178) & "/s"
       Printer.Print "Name:"; Tab(TAB_KINETIC_PARAMETERS_1); "("; IonicTransUnitsOfDisplay; ")"; Tab(TAB_KINETIC_PARAMETERS_2); "("; LiqDiffUnitsOfDisplay; ")"; Tab(TAB_KINETIC_PARAMETERS_3); PoreDiffUnitsOfDisplay
       Printer.Print
    
       For i = 1 To NowProj.NumberOfAnions
           Ion_Name = NowProj.Anion(i).Name
           If Len(Ion_Name) > (TAB_KINETIC_PARAMETERS_1 - 2) Then
              Ion_Name = left$(Ion_Name, TAB_KINETIC_PARAMETERS_1 - 2)
           End If
           LiqDiff = NowProj.Anion(i).Kinetic.LiquidDiffusivity.Value
           LiqDiffToDisplay = Format$(LiqDiff, GetTheFormat(LiqDiff))
           IonicTrans = NowProj.Anion(i).Kinetic.IonicTransportCoefficient.Value
           IonicTransToDisplay = Format$(IonicTrans, GetTheFormat(IonicTrans))
           PoreDiff = NowProj.Anion(i).Kinetic.PoreDiffusivity.Value
           PoreDiffToDisplay = Format$(PoreDiff, GetTheFormat(PoreDiff))
           Printer.Print Ion_Name; Tab(TAB_KINETIC_PARAMETERS_1); IonicTransToDisplay; Tab(TAB_KINETIC_PARAMETERS_2); LiqDiffToDisplay; Tab(TAB_KINETIC_PARAMETERS_3); PoreDiffToDisplay
       Next i
       Printer.Print
    End If

    If (NumSelectedCations <= 1) And (NumSelectedAnions <= 1) Then GoTo EndDocument

    Printer.Print
    Printer.FontBold = True
    Printer.FontUnderline = True
    Printer.FontSize = 12
    Printer.Print "Dimensionless Groups:"
    Printer.Print
    Printer.FontBold = False
    Printer.FontUnderline = False
    Printer.FontSize = PRINT_FONTSIZE_DATA

    If Cations.Available Then
       Printer.Print Tab(TAB_DIMENSIONLESS_GROUPS_1); "Dgs"; Tab(TAB_DIMENSIONLESS_GROUPS_2); "Dgp"; Tab(TAB_DIMENSIONLESS_GROUPS_3); "Dgt"; Tab(TAB_DIMENSIONLESS_GROUPS_4); "Edp"; Tab(TAB_DIMENSIONLESS_GROUPS_5); "St"; Tab(TAB_DIMENSIONLESS_GROUPS_6); "Bip"
       Printer.Print "Name:"; Tab(TAB_DIMENSIONLESS_GROUPS_1); "(-)"; Tab(TAB_DIMENSIONLESS_GROUPS_2); "(-)"; Tab(TAB_DIMENSIONLESS_GROUPS_3); "(-)"; Tab(TAB_DIMENSIONLESS_GROUPS_4); "(-)"; Tab(TAB_DIMENSIONLESS_GROUPS_5); "(-)"; Tab(TAB_DIMENSIONLESS_GROUPS_6); "(-)"
       Printer.Print
    
       For i = 1 To NumSelectedCations
           Ion_Name = NowProj.Cation(Cations_Selected(i)).Name
           If Len(Ion_Name) > (TAB_DIMENSIONLESS_GROUPS_1 - 2) Then
              Ion_Name = left$(Ion_Name, TAB_DIMENSIONLESS_GROUPS_1 - 2)
           End If

           Dgs = NowProj.Cation(Cations_Selected(i)).Dimensionless.SurfaceDistributionParameter
           DgsToDisplay = Format$(Dgs, GetTheFormat(Dgs))
           Dgp = NowProj.Cation(Cations_Selected(i)).Dimensionless.PoreDistributionParameter
           DgpToDisplay = Format$(Dgp, GetTheFormat(Dgp))
           Dgt = NowProj.Cation(Cations_Selected(i)).Dimensionless.TotalDistributionParameter
           DgtToDisplay = Format$(Dgt, GetTheFormat(Dgt))
           Edp = NowProj.Cation(Cations_Selected(i)).Dimensionless.PoreDiffusionModulus
           EdpToDisplay = Format$(Edp, GetTheFormat(Edp))
           St = NowProj.Cation(Cations_Selected(i)).Dimensionless.StantonNumber
           StToDisplay = Format$(St, GetTheFormat(St))
           Bip = NowProj.Cation(Cations_Selected(i)).Dimensionless.PoreBiotNumber
           BipToDisplay = Format$(Bip, GetTheFormat(Bip))
                                                         
           Printer.Print Ion_Name; Tab(TAB_DIMENSIONLESS_GROUPS_1); DgsToDisplay; Tab(TAB_DIMENSIONLESS_GROUPS_2); DgpToDisplay; Tab(TAB_DIMENSIONLESS_GROUPS_3); DgtToDisplay; Tab(TAB_DIMENSIONLESS_GROUPS_4); EdpToDisplay; Tab(TAB_DIMENSIONLESS_GROUPS_5); StToDisplay; Tab(TAB_DIMENSIONLESS_GROUPS_6); BipToDisplay
       Next i
       Printer.Print
    End If

    If Anions.Available Then
       Printer.Print Tab(TAB_DIMENSIONLESS_GROUPS_1); "Dgs"; Tab(TAB_DIMENSIONLESS_GROUPS_2); "Dgp"; Tab(TAB_DIMENSIONLESS_GROUPS_3); "Dgt"; Tab(TAB_DIMENSIONLESS_GROUPS_4); "Edp"; Tab(TAB_DIMENSIONLESS_GROUPS_5); "St"; Tab(TAB_DIMENSIONLESS_GROUPS_6); "Bip"
       Printer.Print "Name:"; Tab(TAB_DIMENSIONLESS_GROUPS_1); "(-)"; Tab(TAB_DIMENSIONLESS_GROUPS_2); "(-)"; Tab(TAB_DIMENSIONLESS_GROUPS_3); "(-)"; Tab(TAB_DIMENSIONLESS_GROUPS_4); "(-)"; Tab(TAB_DIMENSIONLESS_GROUPS_5); "(-)"; Tab(TAB_DIMENSIONLESS_GROUPS_6); "(-)"
       Printer.Print
    
       For i = 1 To NumSelectedAnions
           Ion_Name = NowProj.Anion(Anions_Selected(i)).Name
           If Len(Ion_Name) > (TAB_DIMENSIONLESS_GROUPS_1 - 2) Then
              Ion_Name = left$(Ion_Name, TAB_DIMENSIONLESS_GROUPS_1 - 2)
           End If

           Dgs = NowProj.Anion(Anions_Selected(i)).Dimensionless.SurfaceDistributionParameter
           DgsToDisplay = Format$(Dgs, GetTheFormat(Dgs))
           Dgp = NowProj.Anion(Anions_Selected(i)).Dimensionless.PoreDistributionParameter
           DgpToDisplay = Format$(Dgp, GetTheFormat(Dgp))
           Dgt = NowProj.Anion(Anions_Selected(i)).Dimensionless.TotalDistributionParameter
           DgtToDisplay = Format$(Dgt, GetTheFormat(Dgt))
           Edp = NowProj.Anion(Anions_Selected(i)).Dimensionless.PoreDiffusionModulus
           EdpToDisplay = Format$(Edp, GetTheFormat(Edp))
           St = NowProj.Anion(Anions_Selected(i)).Dimensionless.StantonNumber
           StToDisplay = Format$(St, GetTheFormat(St))
           Bip = NowProj.Anion(Anions_Selected(i)).Dimensionless.PoreBiotNumber
           BipToDisplay = Format$(Bip, GetTheFormat(Bip))
                                                         
           Printer.Print Ion_Name; Tab(TAB_DIMENSIONLESS_GROUPS_1); DgsToDisplay; Tab(TAB_DIMENSIONLESS_GROUPS_2); DgpToDisplay; Tab(TAB_DIMENSIONLESS_GROUPS_3); DgtToDisplay; Tab(TAB_DIMENSIONLESS_GROUPS_4); EdpToDisplay; Tab(TAB_DIMENSIONLESS_GROUPS_5); StToDisplay; Tab(TAB_DIMENSIONLESS_GROUPS_6); BipToDisplay
       Next i
       Printer.Print
    End If

    Printer.Print "Dgs = Surface Distribution Parameter"
    Printer.Print "Dgp = Pore Distribution Parameter"
    Printer.Print "Dgt = Total Equivalent Distribution Parameter"
    Printer.Print "Edp = Pore Diffusion Modulus"
    Printer.Print "St = Modified Stanton Number"
    Printer.Print "Bip = Pore Biot Number"

EndDocument:
    Printer.EndDoc
    Exit Sub

ErrorWhilePrinting:
   msg = Error(Err)
   MsgBox msg, MB_ICONSTOP, "Printing Error"
   Resume ExitPrintRoutine

ExitPrintRoutine:

End Sub

Sub PrintIonExchangeToFile()

End Sub

Sub ReadVarInfluentConcs()
    Dim i As Integer, j As Integer
    Dim NumberOfPoints As Integer, NumberOfComponents As Integer
    Dim FileID As String

  On Error GoTo VarInfluentError

    If Cations.Available And Anions.Available Then

    ElseIf Cations.Available Then
       If NowProj.VarInfluentFileCation = "NONE" Then
          Number_Influent_Points = 0
       Else 'Read in variable influent data
          Open NowProj.VarInfluentFileCation For Input As #1
             Input #1, FileID
             If FileID <> VAR_INFLUENT_CATION_FILEID Then
                Close #1
                GoTo ExitVarInfluent
             End If
             Input #1, NumberOfPoints, NumberOfComponents
             If (NumberOfComponents <> NowProj.NumberOfCations) Or (NumberOfPoints > Number_Max_Influent_Points) Then
                Close #1
                GoTo ExitVarInfluent
             End If
           For i = 1 To NumberOfPoints
               Select Case NumberOfComponents
                  Case 1
                     Input #1, T_Influent(i), C_Influent(1, i)
                  Case 2
                     Input #1, T_Influent(i), C_Influent(1, i), C_Influent(2, i)
                  Case 3
                     Input #1, T_Influent(i), C_Influent(1, i), C_Influent(2, i), C_Influent(3, i)
                  Case 4
                     Input #1, T_Influent(i), C_Influent(1, i), C_Influent(2, i), C_Influent(3, i), C_Influent(4, i)
                  Case 5
                     Input #1, T_Influent(i), C_Influent(1, i), C_Influent(2, i), C_Influent(3, i), C_Influent(4, i), C_Influent(5, i)
                  Case 6
                     Input #1, T_Influent(i), C_Influent(1, i), C_Influent(2, i), C_Influent(3, i), C_Influent(4, i), C_Influent(5, i), C_Influent(6, i)
                  Case 7
                     Input #1, T_Influent(i), C_Influent(1, i), C_Influent(2, i), C_Influent(3, i), C_Influent(4, i), C_Influent(5, i), C_Influent(6, i), C_Influent(7, i)
                  Case 8
                     Input #1, T_Influent(i), C_Influent(1, i), C_Influent(2, i), C_Influent(3, i), C_Influent(4, i), C_Influent(5, i), C_Influent(6, i), C_Influent(7, i), C_Influent(8, i)
                  Case 9
                     Input #1, T_Influent(i), C_Influent(1, i), C_Influent(2, i), C_Influent(3, i), C_Influent(4, i), C_Influent(5, i), C_Influent(6, i), C_Influent(7, i), C_Influent(8, i), C_Influent(9, i)
                  Case 10
                     Input #1, T_Influent(i), C_Influent(1, i), C_Influent(2, i), C_Influent(3, i), C_Influent(4, i), C_Influent(5, i), C_Influent(6, i), C_Influent(7, i), C_Influent(8, i), C_Influent(9, i), C_Influent(10, i)
               End Select
           Next i
             Number_Influent_Points = NumberOfPoints
          Close #1
       End If
    ElseIf Anions.Available Then

    End If

    Exit Sub

VarInfluentError:
   MsgBox "Error in reading data file for variable influent points.  The number of influent points will be set to zero.", MB_ICONINFORMATION
   Resume ExitVarInfluent

ExitVarInfluent:
   Number_Influent_Points = 0

End Sub

Sub SaveFileIonExchange(filename As String)

    On Error Resume Next
    frmIonExchangeMain!CMDialog1.DefaultExt = "iex"
    frmIonExchangeMain!CMDialog1.Filter = "Ion Exchange Files (*.iex)|*.iex"
    frmIonExchangeMain!CMDialog1.DialogTitle = "Save Ion Exchange File"
    frmIonExchangeMain!CMDialog1.flags = OFN_OVERWRITEPROMPT Or OFN_PATHMUSTEXIST
'------Begin Modification Hokanson: 12-Aug2000
    frmIonExchangeMain!CMDialog1.CancelError = True
'------End Modification Hokanson: 11-Aug2000
    frmIonExchangeMain!CMDialog1.Action = 2
    filename$ = frmIonExchangeMain!CMDialog1.filename
    If Err = 32755 Then   'Cancel selected by user
       filename$ = ""
    End If

End Sub

Sub SaveFileVariableInfluent(VarInfluentFileName As String)

    On Error Resume Next
    frmConcentrations!CMDialog1.DefaultExt = "var"
    frmConcentrations!CMDialog1.Filter = "Variable Influent (*.var)|*.var"
    frmConcentrations!CMDialog1.DialogTitle = "Ion Exchange Variable Influent Files"
    frmConcentrations!CMDialog1.flags = OFN_OVERWRITEPROMPT Or OFN_PATHMUSTEXIST
'------Begin Modification Hokanson: 12-Aug2000
    frmIonExchangeMain!CMDialog1.CancelError = True
'------End Modification Hokanson: 11-Aug2000
    frmConcentrations!CMDialog1.Action = 2
    VarInfluentFileName$ = frmConcentrations!CMDialog1.filename
    If Err = 32755 Then   'Cancel selected by user
       VarInfluentFileName$ = ""
    End If

End Sub

Sub SaveIonExchange()
    Dim FileID As String
    Dim i As Integer
    Dim VariableToDetermineRecordLength As ComponentPropertyType
    Dim StringLength As Integer
    Dim MainString As String

    If filename$ = "" Then
       Call SaveFileIonExchange(filename)
    End If

    If filename$ <> "" Then
       FileID = ION_EXCHANGE_FILEID
       Open filename$ For Output As #1
       
       Write #1, FileID
      
       'Write Pressure and Temperature
       Write #1, "Pressure", NowProj.Operating.Pressure, frmIonExchangeMain!cboOperatingConditionsUnits(0).List(0), frmIonExchangeMain!cboOperatingConditionsUnits(0).List(frmIonExchangeMain!cboOperatingConditionsUnits(0).ListIndex)
       Write #1, "Temperature", NowProj.Operating.Temperature, frmIonExchangeMain!cboOperatingConditionsUnits(1).List(0), frmIonExchangeMain!cboOperatingConditionsUnits(1).List(frmIonExchangeMain!cboOperatingConditionsUnits(1).ListIndex)

       'Write Bed Data
       Write #1, "Bed Length", NowProj.Bed.length, frmIonExchangeMain!cboBedDataUnits(0).List(0), frmIonExchangeMain!cboBedDataUnits(0).List(frmIonExchangeMain!cboBedDataUnits(0).ListIndex)
       Write #1, "Bed Diameter", NowProj.Bed.Diameter, frmIonExchangeMain!cboBedDataUnits(1).List(0), frmIonExchangeMain!cboBedDataUnits(1).List(frmIonExchangeMain!cboBedDataUnits(1).ListIndex)
       Write #1, "Bed Weight", NowProj.Bed.Weight, frmIonExchangeMain!cboBedDataUnits(2).List(0), frmIonExchangeMain!cboBedDataUnits(2).List(frmIonExchangeMain!cboBedDataUnits(2).ListIndex)
       Write #1, "Bed Flowrate", NowProj.Bed.Flowrate.Value, frmIonExchangeMain!cboBedDataUnits(3).List(0), frmIonExchangeMain!cboBedDataUnits(3).List(frmIonExchangeMain!cboBedDataUnits(3).ListIndex), NowProj.Bed.Flowrate.UserInput
       Write #1, "Bed EBCT", NowProj.Bed.EBCT.Value, frmIonExchangeMain!cboBedDataUnits(4).List(0), frmIonExchangeMain!cboBedDataUnits(4).List(frmIonExchangeMain!cboBedDataUnits(4).ListIndex), NowProj.Bed.EBCT.UserInput
       Write #1, "Number of Beds (in series)", NowProj.Bed.NumberOfBeds

       'Write Adsorbent Properties
       Write #1, "Resin Name", NowProj.Resin.Name
       Write #1, "Apparent Density", NowProj.Resin.ApparentDensity, frmIonExchangeMain!cboAdsorbentPropertyUnits(1).List(0), frmIonExchangeMain!cboAdsorbentPropertyUnits(1).List(frmIonExchangeMain!cboAdsorbentPropertyUnits(1).ListIndex)
       Write #1, "Particle Radius", NowProj.Resin.ParticleRadius, frmIonExchangeMain!cboAdsorbentPropertyUnits(2).List(0), frmIonExchangeMain!cboAdsorbentPropertyUnits(2).List(frmIonExchangeMain!cboAdsorbentPropertyUnits(2).ListIndex)
       Write #1, "Particle Porosity (-)", NowProj.Resin.ParticlePorosity
       Write #1, "Tortuosity (-)", NowProj.Resin.Tortuosity
       Write #1, "Total Resin Capacity", NowProj.Resin.TotalCapacity, frmIonExchangeMain!cboAdsorbentPropertyUnits(5).List(0), frmIonExchangeMain!cboAdsorbentPropertyUnits(5).List(frmIonExchangeMain!cboAdsorbentPropertyUnits(5).ListIndex)

       'Write Time Parameters
       Write #1, "Time Parameters - Total Run Time", NowProj.TimeParameters.FinalTime, frmOptionsInputParameters!cboTimeParametersUnits(0).List(0), frmOptionsInputParameters!cboTimeParametersUnits(0).List(frmOptionsInputParameters!cboTimeParametersUnits(0).ListIndex)
       Write #1, "Time Parameters - InitialTime", NowProj.TimeParameters.InitialTime, frmOptionsInputParameters!cboTimeParametersUnits(1).List(0), frmOptionsInputParameters!cboTimeParametersUnits(1).List(frmOptionsInputParameters!cboTimeParametersUnits(1).ListIndex)
       Write #1, "Time Parameters - Time Step", NowProj.TimeParameters.TimeStep, frmOptionsInputParameters!cboTimeParametersUnits(2).List(0), frmOptionsInputParameters!cboTimeParametersUnits(2).List(frmOptionsInputParameters!cboTimeParametersUnits(2).ListIndex)

'------Begin Modification Hokanson: 11-Aug2000
       'Input EPS_ErrorCriteria and DH0_Initial Time step for DGEAR integrator
       Write #1, "EPS_ErrorCriteriaForDGEARIntegrator", EPS_ErrorCriteriaForDGEARIntegrator
       Write #1, "DH0_InitialTimeStepForDGEARIntegrator", DH0_InitialTimeStepForDGEARIntegrator, frmOptionsInputParameters!cboTimeParametersUnits(4).List(0), frmOptionsInputParameters!cboTimeParametersUnits(4).List(frmOptionsInputParameters!cboTimeParametersUnits(4).ListIndex)
'------End Modification Hokanson: 11-Aug2000

       'Write Collocation Points
       Write #1, "Number of Axial Collocation Points", NowProj.NumAxialCollocationPoints
       Write #1, "Number of Radial Collocation Points", NowProj.NumRadialCollocationPoints

       'Correlation Used to Calculate Ionic Transport Coefficient
       Write #1, "Correlation for Ionic Transport Coeff., kf", NowProj.IonicTransportCoeffCorrName

       'Write Component Properties

       Write #1, "Component Properties - Properties of Cations"
       Write #1, "Number of Cations", NowProj.NumberOfCations
       Write #1, "Presaturant Cation", NowProj.PresaturantCation
       Write #1, "Sum of Cation Time-Averaged Initial Influent Concs.", NowProj.SumCationInitialEquivalents, NowProj.OKToGetCationDimensionless
       Write #1, "Separation Factor Info.", NowProj.CationSeparationFactorInput.Row, NowProj.CationSeparationFactorInput.Value
       For i = 1 To NowProj.NumberOfCations
           Write #1, NowProj.Cation(i).Name
           Write #1, NowProj.Cation(i).MolecularWeight, NowProj.Cation(i).InitialConcentration, NowProj.Cation(i).EquivalentInitialConcentration, NowProj.Cation(i).Valence, NowProj.Cation(i).SeparationFactor
           Write #1, NowProj.Cation(i).Kinetic.LiquidDiffusivity.Value, NowProj.Cation(i).Kinetic.LiquidDiffusivity.UserInput
           Write #1, NowProj.Cation(i).Kinetic.LiquidDiffusivityCorrelation, frmInputKineticParameters!cboLiquidDiffusivityUnits(0).List(0)
           Write #1, NowProj.Cation(i).Kinetic.LiquidDiffusivityUserInput, frmInputKineticParameters!cboLiquidDiffusivityUnits(1).List(0)
           Write #1, NowProj.Cation(i).Kinetic.IonicTransportCoefficient.Value, NowProj.Cation(i).Kinetic.IonicTransportCoefficient.UserInput
           Write #1, NowProj.Cation(i).Kinetic.IonicTransportCoeffCorrelation, frmInputKineticParameters!cboIonicTransportUnits(0).List(0)
           Write #1, NowProj.Cation(i).Kinetic.IonicTransportCoeffUserInput, frmInputKineticParameters!cboIonicTransportUnits(1).List(0)
           Write #1, NowProj.Cation(i).Kinetic.PoreDiffusivity.Value, NowProj.Cation(i).Kinetic.PoreDiffusivity.UserInput
           Write #1, NowProj.Cation(i).Kinetic.PoreDiffusivityCorrelation, frmInputKineticParameters!cboPoreDiffusivityUnits(0).List(0)
           Write #1, NowProj.Cation(i).Kinetic.PoreDiffusivityUserInput, frmInputKineticParameters!cboPoreDiffusivityUnits(1).List(0)
           Write #1, NowProj.Cation(i).Kinetic.NernstHaskellCation.Ion_Name, NowProj.Cation(i).Kinetic.NernstHaskellCation.Valence, NowProj.Cation(i).Kinetic.NernstHaskellCation.LimitingIonicConductance
           Write #1, NowProj.Cation(i).Kinetic.NernstHaskellAnion.Ion_Name, NowProj.Cation(i).Kinetic.NernstHaskellAnion.Valence, NowProj.Cation(i).Kinetic.NernstHaskellAnion.LimitingIonicConductance
       Next i

       Write #1, "Component Properties - Properties of Anions"
       Write #1, "Number of Anions", NowProj.NumberOfAnions
       Write #1, "Presaturant Anion", NowProj.PresaturantAnion
       Write #1, "Sum of Anion Time-Averaged Initial Influent Concs.", NowProj.SumAnionInitialEquivalents, NowProj.OKToGetAnionDimensionless
       Write #1, "Separation Factor Info.", NowProj.AnionSeparationFactorInput.Row, NowProj.AnionSeparationFactorInput.Value
       For i = 1 To NowProj.NumberOfAnions
           Write #1, NowProj.Anion(i).Name
           Write #1, NowProj.Anion(i).MolecularWeight, NowProj.Anion(i).InitialConcentration, NowProj.Anion(i).EquivalentInitialConcentration, NowProj.Anion(i).Valence, NowProj.Anion(i).SeparationFactor
           Write #1, NowProj.Anion(i).Kinetic.LiquidDiffusivity.Value, NowProj.Anion(i).Kinetic.LiquidDiffusivity.UserInput
           Write #1, NowProj.Anion(i).Kinetic.LiquidDiffusivityCorrelation, frmInputKineticParameters!cboLiquidDiffusivityUnits(0).List(0)
           Write #1, NowProj.Anion(i).Kinetic.LiquidDiffusivityUserInput, frmInputKineticParameters!cboLiquidDiffusivityUnits(1).List(0)
           Write #1, NowProj.Anion(i).Kinetic.IonicTransportCoefficient.Value, NowProj.Anion(i).Kinetic.IonicTransportCoefficient.UserInput
           Write #1, NowProj.Anion(i).Kinetic.IonicTransportCoeffCorrelation, frmInputKineticParameters!cboIonicTransportUnits(0).List(0)
           Write #1, NowProj.Anion(i).Kinetic.IonicTransportCoeffUserInput, frmInputKineticParameters!cboIonicTransportUnits(1).List(0)
           Write #1, NowProj.Anion(i).Kinetic.PoreDiffusivity.Value, NowProj.Anion(i).Kinetic.PoreDiffusivity.UserInput
           Write #1, NowProj.Anion(i).Kinetic.PoreDiffusivityCorrelation, frmInputKineticParameters!cboPoreDiffusivityUnits(0).List(0)
           Write #1, NowProj.Anion(i).Kinetic.PoreDiffusivityUserInput, frmInputKineticParameters!cboPoreDiffusivityUnits(1).List(0)
           Write #1, NowProj.Anion(i).Kinetic.NernstHaskellCation.Ion_Name, NowProj.Anion(i).Kinetic.NernstHaskellCation.Valence, NowProj.Anion(i).Kinetic.NernstHaskellCation.LimitingIonicConductance
           Write #1, NowProj.Anion(i).Kinetic.NernstHaskellAnion.Ion_Name, NowProj.Anion(i).Kinetic.NernstHaskellAnion.Valence, NowProj.Anion(i).Kinetic.NernstHaskellAnion.LimitingIonicConductance
       Next i

       'Units of Display of Properties Related to Components
       Write #1, "Molecular Weight Units", frmAddComponent!cboAddIonUnits(0).List(frmAddComponent!cboAddIonUnits(0).ListIndex)
       Write #1, "Initial Conc. Units", frmAddComponent!cboAddIonUnits(1).List(frmAddComponent!cboAddIonUnits(1).ListIndex)
       Write #1, "Liquid Diffusivity Correlation Units", frmInputKineticParameters!cboLiquidDiffusivityUnits(0).List(frmInputKineticParameters!cboLiquidDiffusivityUnits(0).ListIndex)
       Write #1, "Liquid Diffusivity User Input Units", frmInputKineticParameters!cboLiquidDiffusivityUnits(1).List(frmInputKineticParameters!cboLiquidDiffusivityUnits(1).ListIndex)
       Write #1, "Ionic Transport Coeff Correlation Units", frmInputKineticParameters!cboIonicTransportUnits(0).List(frmInputKineticParameters!cboIonicTransportUnits(0).ListIndex)
       Write #1, "Ionic Transport Coeff User Input Units", frmInputKineticParameters!cboIonicTransportUnits(1).List(frmInputKineticParameters!cboIonicTransportUnits(1).ListIndex)
       Write #1, "Pore Diffusivity Correlation Units", frmInputKineticParameters!cboPoreDiffusivityUnits(0).List(frmInputKineticParameters!cboPoreDiffusivityUnits(0).ListIndex)
       Write #1, "Pore Diffusivity User Input Units", frmInputKineticParameters!cboPoreDiffusivityUnits(1).List(frmInputKineticParameters!cboPoreDiffusivityUnits(1).ListIndex)

       Write #1, "Name of File for Cation Variable Influent", NowProj.VarInfluentFileCation
       Write #1, "Name of File for Anion Variable Influent", NowProj.VarInfluentFileAnion

       Close #1

       If filename$ <> OldFileName$ Then
          frmIonExchangeMain.Caption = "Ion Exchange Simulation Software - " & filename
       End If
       
      Current_Filename = filename
      Call OldFileList_Promote( _
      Current_Filename, _
      1, _
      frmIonExchangeMain.mnuFile(13), _
      frmIonExchangeMain.mnuFile(8), _
      frmIonExchangeMain.mnuFile(9), _
      frmIonExchangeMain.mnuFile(10), _
      frmIonExchangeMain.mnuFile(11))
      
    Else
       filename$ = OldFileName$

    End If

End Sub

Sub SaveVariableInfluent(VarInfluentFileName As String)
    Dim i As Integer, j As Integer

        Open VarInfluentFileName For Output As #1
           If Cations.Available And Anions.Available Then
           ElseIf Cations.Available Then
              Write #1, VAR_INFLUENT_CATION_FILEID
           ElseIf Anions.Available Then
              Write #1, VAR_INFLUENT_ANION_FILEID
           End If
           Write #1, Number_Influent_Points, Total_NumberOfComponents
           
           For i = 1 To Number_Influent_Points
               Select Case Total_NumberOfComponents
                  Case 1
                     Write #1, T_Influent(i), C_Influent(1, i)
                  Case 2
                     Write #1, T_Influent(i), C_Influent(1, i), C_Influent(2, i)
                  Case 3
                     Write #1, T_Influent(i), C_Influent(1, i), C_Influent(2, i), C_Influent(3, i)
                  Case 4
                     Write #1, T_Influent(i), C_Influent(1, i), C_Influent(2, i), C_Influent(3, i), C_Influent(4, i)
                  Case 5
                     Write #1, T_Influent(i), C_Influent(1, i), C_Influent(2, i), C_Influent(3, i), C_Influent(4, i), C_Influent(5, i)
                  Case 6
                     Write #1, T_Influent(i), C_Influent(1, i), C_Influent(2, i), C_Influent(3, i), C_Influent(4, i), C_Influent(5, i), C_Influent(6, i)
                  Case 7
                     Write #1, T_Influent(i), C_Influent(1, i), C_Influent(2, i), C_Influent(3, i), C_Influent(4, i), C_Influent(5, i), C_Influent(6, i), C_Influent(7, i)
                  Case 8
                     Write #1, T_Influent(i), C_Influent(1, i), C_Influent(2, i), C_Influent(3, i), C_Influent(4, i), C_Influent(5, i), C_Influent(6, i), C_Influent(7, i), C_Influent(8, i)
                  Case 9
                     Write #1, T_Influent(i), C_Influent(1, i), C_Influent(2, i), C_Influent(3, i), C_Influent(4, i), C_Influent(5, i), C_Influent(6, i), C_Influent(7, i), C_Influent(8, i), C_Influent(9, i)
                  Case 10
                     Write #1, T_Influent(i), C_Influent(1, i), C_Influent(2, i), C_Influent(3, i), C_Influent(4, i), C_Influent(5, i), C_Influent(6, i), C_Influent(7, i), C_Influent(8, i), C_Influent(9, i), C_Influent(10, i)
               End Select
           Next i

        Close #1

End Sub

Sub TextGetFocus(txt As TextBox, Temp_Text As String)
    Temp_Text = txt.Text
    txt.SelStart = 0
    txt.SelLength = Len(txt.Text)

End Sub

Sub TextHandleError(IsError As Integer, txt As TextBox, Temp_Text As String)
    Dim Dummy As Double
    Dim i As Integer

    IsError = False
    On Error GoTo ErrorHandler
       Dummy = CDbl(txt.Text)
'       If Dummy < 0# Then GoTo NegativeNumberError
       If IsError Then txt.SetFocus
       GoTo ContinueSub

ErrorHandler:
    IsError = True
    'frmAirWaterProperties.Print "Error Occurred"
    MsgBox "Incorrect Value Will Be Replaced By Previous Value", , "Invalid Data Error"
    txt.Text = Temp_Text
    Resume

NegativeNumberError:
    IsError = True
    txt.Text = Temp_Text
    txt.SetFocus

ContinueSub:

End Sub

