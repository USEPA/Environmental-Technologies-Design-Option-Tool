Attribute VB_Name = "modmain"
Option Explicit
Sub TabFolderEnable(value As Boolean)
'// REMOVED BY EJOMAN (TABViewProp).
    'frmmain!TABViewProp.TabEnabled(0) = value
    'frmmain!TABViewProp.TabEnabled(1) = value
    'frmmain!TABViewProp.TabEnabled(2) = value
    'frmmain!TABViewProp.TabEnabled(3) = value
    'frmmain!TABViewProp.TabEnabled(4) = value
    'frmmain!TABViewProp.TabEnabled(5) = value
    'frmmain!TABViewProp.TabEnabled(7) = value
    'frmmain!TABViewProp.TabEnabled(8) = value

End Sub

Sub LoadDemo()

    Dim DBTbl As Recordset
    Dim Response As Integer
    Dim CLCAS As Long
    Dim ULCAS As Long
    Dim LastList As Integer
    
    On Error GoTo LoadDemoError
    If PathDemo = NULLPATH Then
        MsgBox ("Demo file not found")
        Exit Sub
    End If
    'Set mousepointer to hourglass (wait mode)
    Screen.MousePointer = 11
    
    'Refresh main form
    frmmain.Refresh
        
    DBJetUser.Close
    frmmain!Data2.databasename = App.path & "\temp.mdb"
    frmmain!Data2.RecordSource = "User List"
    frmmain!Data2.Refresh
    
        ' we're removing the old dbuser and copying the demo file to make a new one
    Kill PathUser
    SaveFileName = PathDemo & "\demo.prl"
    FileCopy SaveFileName, PathUser & "\dbuser.mdb"
    Set DBJetUser = OpenDatabase(PathUser, False, False)
        
    UserDBName = PathUser
    frmmain!Data2.databasename = PathUser & "\dbuser.mdb"
    frmmain!Data2.RecordSource = "User List"
    frmmain!Data2.RecordsetType = 2
    frmmain!Data2.Refresh
    
    Set DBTbl = DBJetUser.OpenRecordset("Last CAS", dbOpenTable)
        
    Cur_Info.CAS = DBTbl("LastCAS")
    CLCAS = DBTbl("CLCAS")
    ULCAS = DBTbl("ULCAS")
    LastList = DBTbl("LastList")
    
    DBTbl.Close
    
    frmmain!Data1.Recordset.FindFirst "CAS =" & Cur_Info.CAS
    
    If GetUserData = True Then
        If LastList = 1 And CLCAS = Cur_Info.CAS Then
            Call LoadUserPreferences
            Call Recalculate
            Call DisplayProps
            Call TabFolderEnable(True)
        End If
        If LastList = 2 And ULCAS = Cur_Info.CAS Then
            Call LoadUserPreferences
            Call Recalculate
            Call DisplayProps
            Call TabFolderEnable(True)
        End If
    End If
        
    frmmain!Data1.Recordset.FindFirst "CAS =" & CLCAS
    frmmain!LSTSelList.Text = frmmain!Data1.Recordset("Name")
    
    frmmain!Data2.Recordset.FindFirst "CAS =" & ULCAS
    frmmain!LSTUserList.Text = frmmain!Data2.Recordset("Name")
        
    If LastList = 1 Then
        frmmain!LSTSelList.SetFocus
    Else
        frmmain!LSTUserList.SetFocus
    End If
        
    'Set mousepointer to arrow (normal mode)
    Screen.MousePointer = 1
        
    Exit Sub
       
LoadDemoError:
    Screen.MousePointer = 1
    If Err <> 32755 Then
        MsgBox "Error loading PEARLS file", 48, "Error"
    End If

End Sub

Sub LoadUNIFACCalcData()
            
    Dim i As Integer
    Dim J As Integer
    Dim MainGroup As Integer
    Dim DBTbl As Recordset
    
    Set DBJetMaster = OpenDatabase(PathMaster, False, True)
        
    'Load AGLB BIP database
    Set DBTbl = DBJetMaster.OpenRecordset("AGLB", dbOpenTable)
    On Error GoTo error_loading_UNIFAC
    For i = 1 To 58
        For J = 1 To 58
            MainGroup = DBTbl("Main Group")
            BIP(1, MainGroup, J) = DBTbl(J)
        Next J
        DBTbl.MoveNext
    Next i
    DBTbl.Close
        
    'Load AVLE BIP database
    Set DBTbl = DBJetMaster.OpenRecordset("AVLE", dbOpenTable)
    
    For i = 1 To 58
        For J = 1 To 58
            MainGroup = DBTbl("Main Group")
            BIP(2, MainGroup, J) = DBTbl(J)
        Next J
        DBTbl.MoveNext
    Next i
    DBTbl.Close
    
    'Load AENV BIP database
    Set DBTbl = DBJetMaster.OpenRecordset("AENV", dbOpenTable)
    
    For i = 1 To 58
        For J = 1 To 58
            MainGroup = DBTbl("Main Group")
            BIP(3, MainGroup, J) = DBTbl(J)
        Next J
        DBTbl.MoveNext
    Next i
    DBTbl.Close
    
    'Load ALLE BIP database
    Set DBTbl = DBJetMaster.OpenRecordset("ALLE", dbOpenTable)
    
    For i = 1 To 32
        For J = 1 To 32
            MainGroup = DBTbl("Main Group")
            BIP(4, MainGroup, J) = DBTbl(J)
        Next J
        DBTbl.MoveNext
    Next i
    DBTbl.Close
    
    'Load area and group parameters and MW and MV group contributions
    Set DBTbl = DBJetMaster.OpenRecordset("UNIFAC", dbOpenTable)
    
    For i = 1 To 116
        MGSG(i) = DBTbl("Main Group")
        RI(i) = DBTbl("Rk")
        QI(i) = DBTbl("Qk")
        MWS(i) = DBTbl("MW Group")
        MVS(i) = DBTbl("MV Group")
        DBTbl.MoveNext
    Next i
    DBTbl.Close
    Exit Sub
error_loading_UNIFAC:
   MsgBox ("ERror loading unifac data")
   Resume Next
End Sub

Sub LoadUnitConversions()
    
    Dim i As Integer
    Dim DBTbl As Recordset
            
    On Error Resume Next
    
    Set DBTbl = DBJetMaster.OpenRecordset("Unit Conversions", dbOpenTable)

    i = 1
    Do While Not DBTbl.EOF
        Unit1(i) = DBTbl("DefaultUnit")
        Unit2(i) = DBTbl("UnitTo")
        AddProp(i) = DBTbl("PropMult")
        Mult(i) = DBTbl("Mult")
        AddConst(i) = DBTbl("AddEnd")
        AddPropOpFlag(i) = DBTbl("OperFlag")
        DBTbl.MoveNext
        i = i + 1
    Loop
    Unit1(i) = "End"
    DBTbl.Close
    
End Sub

Sub SetUpSecurity()

End Sub

Sub InitializeVariables()
    
    ' InitializeVariables:  Initializes:
    '                           - preference variables
    '                           - CurInfo stuff
    Dim i As Integer
    Dim J As Integer
    'Set number of method screens to 0
    ScreenNum = 0

    UserDBName = PathUser
    SaveFileName = ""
    
    'Set sorting flags
    SortChemListAsc = True
    SortUserListAsc = True
    
    'Set CAS number
    Cur_Info.CAS = 0
        
    'Set the operating conditions to STP
    Cur_Info.OpT = 298.15
    Cur_Info.OpTUnit = Get_DefaultUnit(OptTemp)
    Cur_Info.OpP = 101000
    Cur_Info.OpPUnit = Get_DefaultUnit(OptPress)
    
    'Set user input and F(T) temperature
    For i = 0 To NumProperties
        InfoMethod(i).value(10) = 0
        InfoMethod(i).MethodName(10) = "User Input"
        InfoMethod(i).Unit = ""
        InfoMethod(i).TFT = 298.15
    Next i
        
    'Set user default equation numbers
    InfoMethod(VP).EqNum(10) = 101
    InfoMethod(ST).EqNum(10) = 106
    InfoMethod(LD).EqNum(10) = 105
    InfoMethod(VV).EqNum(10) = 102
    InfoMethod(LV).EqNum(10) = 101
    InfoMethod(LHC).EqNum(10) = 100
    InfoMethod(VHC).EqNum(10) = 107
    InfoMethod(LTC).EqNum(10) = 100
    InfoMethod(VTC).EqNum(10) = 102
    InfoMethod(Hvap).EqNum(10) = 106
    
    'mrt- setup antoine info. This is done by saying there is no info.
    Antoine_Info.AntCalc = False
    print_antoine = False
    
    'Set calculation hierarchy for predictive methods
    CalcHierarchy(0) = 0
    CalcHierarchy(1) = 1
    CalcHierarchy(2) = 3
    CalcHierarchy(3) = 34
    CalcHierarchy(4) = 39
    CalcHierarchy(5) = 35
    CalcHierarchy(6) = 2
    CalcHierarchy(7) = 4
    CalcHierarchy(8) = 5
    CalcHierarchy(9) = 6
    CalcHierarchy(10) = 7
    CalcHierarchy(11) = 8
    CalcHierarchy(12) = 9
    CalcHierarchy(13) = 11
    CalcHierarchy(14) = 10
    CalcHierarchy(15) = 12
    CalcHierarchy(16) = 13
    CalcHierarchy(17) = 14
    CalcHierarchy(18) = 15
    CalcHierarchy(19) = 16
    CalcHierarchy(20) = 17
    CalcHierarchy(21) = 18
    CalcHierarchy(22) = 19
    CalcHierarchy(23) = 20
    CalcHierarchy(24) = 21
    CalcHierarchy(25) = 22
    CalcHierarchy(26) = 23
    CalcHierarchy(27) = 24
    CalcHierarchy(28) = 25
    CalcHierarchy(29) = 26
    CalcHierarchy(30) = 27
    CalcHierarchy(31) = 28
    CalcHierarchy(32) = 29
    CalcHierarchy(33) = 30
    CalcHierarchy(34) = 31
    CalcHierarchy(35) = 32
    CalcHierarchy(36) = 33
    CalcHierarchy(37) = 36
    CalcHierarchy(38) = 37
    CalcHierarchy(39) = 38
    CalcHierarchy(40) = 40
    CalcHierarchy(41) = 41
    CalcHierarchy(42) = 42
    CalcHierarchy(43) = 43
    CalcHierarchy(44) = 44
    CalcHierarchy(45) = 45
    CalcHierarchy(46) = 46
    CalcHierarchy(47) = 47
    CalcHierarchy(48) = 48
    CalcHierarchy(49) = 49
    CalcHierarchy(50) = 50
    CalcHierarchy(51) = 51
    CalcHierarchy(52) = 52
    CalcHierarchy(53) = 53
    CalcHierarchy(54) = 54

    ' initialize the Block 5 preferences
    ' remember the array is indexed from 0 but
    ' the methods are represented from 1 to NumMethods
    For i = 0 To 3
        For J = 0 To NumMethods - 1
            B5Preference(i, J) = J + 1
        Next J
    Next i
    
    'set display version of cur_info
    Call update_DisplayData
    
End Sub

Sub FillGauge(NumFill As Integer, LastStep As Integer)
    
    'Fill in initialization bar for each step
     
'msh/mrt 10/29/98
'   Modified Pauls code...added progressbar oppose to
'   picturebox.
'
    Dim Percent As Double
    
    Percent = frmtitle.ProgressBar1.value
    
    'if done fill guage
    If LastStep = True Then
        Percent = 100
        frmtitle!ProgressBar1.value = Percent
        Exit Sub
    End If
    
    'else fill % done
    frmtitle!ProgressBar1.value = Percent + (100 / NumFill)
   
End Sub


Sub Main()
    
    Dim NumSteps As Integer
    Dim Response As Integer
    Dim Msg As String
    Dim Title As String
    Dim Style As String
    Dim i As Integer
    
    'Number of steps in startup
    NumSteps = 8
    ' initialize the variables needed for file stuff
    Call Initialize_Nulls
    Call Initialize_File_Stuff
    'Initialize flood gauge and show title screen
    frmtitle.Show
    frmtitle.Refresh
    
    
    
    'Let the user enter a name if they want (this is used to identify the .def file)
    'this also calls the frmmaster to confirm file settings and allow the user to browse, if
    'appropriate
    frmuser.Show 1
    ' now make sure 801 and 911 are set to look for master loaded (should be, but just in case)
    Path801 = PathMaster
    Path911 = PathMaster
    
    'Check Whether this is the first use for this user (we base this
    '   on whether they've got a CUE in their .def file)
    Call Check_For_Cues
    
    'A new session, not yet modified
    WorkModified = False
    frmmain.caption = "PEARLS:  " & SaveFileName & " unmodified"
    
    'Set mousepointer to hourglass
    
   
    
    ' initialize the variables needed for file stuff
    'Call Initialize_File_Stuff
    
    ' Call the function that will read the def file, check the paths, and
    ' call the browser if necessary
    ' DENISE fix this, hadto comment out 11/12/97
    ' DENISE 3/12/98 - we're only using the frmmaster to select files - doesn't check whether
    ' the files are of a valid format!
    
    While simple_check_files = False
        Msg = "Some pearls files were not found, open browser to set paths?"
        Style = vbYesNo + vbCritical + vbDefaultButton1
        Title = "Set paths or exit pearls?"
        Response = MsgBox(Msg, Style, Title)
        
        If Response = vbNo Then
            MsgBox ("Exiting Pearls: Files Not Found")
            End
            Exit Sub
        Else
            Call load_frm_master_info
            Call CenterForm(frmmaster)
            frmmaster.Show 1
        End If
    Wend
    Screen.MousePointer = 11
    ' set the save path to app.path for now
    
    PathUser = App.path & "\dbuser.mdb"
    'Create new user database
    FileCopy PathSave, PathUser
    Set DBJetUser = OpenDatabase(PathUser, False, False)
                
    'Set up database security
    Call SetUpSecurity
    Call FillGauge(NumSteps, False)
    
    'Initialize global variables
    Call InitializeVariables
    Call FillGauge(NumSteps, False)
               
    'Load supplementary UNIFAC calculation data
    Call LoadUNIFACCalcData
    Call FillGauge(NumSteps, False)
    
    'Load unit conversions
    Call LoadUnitConversions
    Call FillGauge(NumSteps, False)
    
    'Set up user database for Save/Load option
    Call LoadUserPreferences
    Call FillGauge(NumSteps, False)
       
    'Load main form
    Load frmmain
    Call FillGauge(NumSteps, False)
                                       
    'Open chemical list
    frmmain!Data1.databasename = PathMaster
    frmmain!Data1.RecordSource = "PEARLS List"
    frmmain!Data1.RecordsetType = 2
    frmmain!Data1.Refresh
    frmmain!Data2.databasename = PathUser
    frmmain!Data2.RecordSource = "User List"
    frmmain!Data2.RecordsetType = 2
    frmmain!Data2.Refresh
    Call FillGauge(NumSteps, False)
    
    'Load data for first chemical
    frmmain!Data1.Recordset.FindFirst "Name = 'TOLUENE'"
    frmmain!LSTSelList.Text = "TOLUENE"
    frmmain!TXTFamily.Text = GetFamilyGroup(frmmain!Data1.Recordset("Chemical Family"))
'// REMOVED BY EJOMAN (TABViewProp).
    'frmmain!TABViewProp.CurrTab = 6
    Call TabFolderEnable(False)
    Call FillGauge(NumSteps, True)
    
    'Display operating T and P
    frmmain!TXTOpT.Text = FormatVal(Cur_Disp.OpT)
    frmmain!LBLOpTUnits.caption = Cur_Disp.OpTUnit
    frmmain!TXTOpP.Text = FormatVal(Cur_Disp.OpP)
    frmmain!LBLOpPUnits.caption = Cur_Disp.OpPUnit
    
    'initialize these so unit prefs doesnt choke
    For i = 0 To NumProperties
        InfoMethod(i).Unit = ""
        InfoMethod(i).TFTUnit = ""
    Next
    
    frmtitle.Hide
    frmmain.Show
    frmmain.Refresh
    Unload frmtitle
        
    'Set mousepointer to arrow
    Screen.MousePointer = 1
        
End Sub




Sub LoadUserPreferences()
    
    Dim i As Integer
    Dim J As Integer
    Dim whichB5 As Integer
    Dim DBTbl As Recordset
    
    'Load BIP database hierarchies
    Set DBTbl = DBJetUser.OpenRecordset("PrefDatabases", dbOpenTable)

    DBTbl.MoveFirst
    DIPPR801 = DBTbl("Available")
    
    DBTbl.MoveNext
    DIPPR911 = DBTbl("Available")
    
    
    DBTbl.Close
    
    'Load BIP database hierarchies
    Set DBTbl = DBJetUser.OpenRecordset("PrefBIPHierarchy", dbOpenTable)
    
    DBTbl.MoveFirst
    For i = 1 To 3
        BIPHierarchy(i, 1) = DBTbl("BIP 1")
        BIPHierarchy(i, 2) = DBTbl("BIP 2")
        BIPHierarchy(i, 3) = DBTbl("BIP 3")
        BIPHierarchy(i, 4) = DBTbl("BIP 4")
        DBTbl.MoveNext
    Next i
    
    DBTbl.Close
    
    ' Load Block 5 preferences, if they are there (first restore defaults)
    On Error GoTo after_B5_pref
    For i = 0 To 4
        For J = 0 To 6
            B5Preference(i, J) = J + 1
        Next J
    Next i
    
    ' DENISE FIX THIS
    Set DBTbl = DBJetUser.OpenRecordset("Block5pref", dbOpenTable)
    DBTbl.MoveFirst
    For i = 0 To 3
        If DBTbl.EOF = True Then
            Exit For
        End If
        Select Case Trim(DBTbl("Property"))
            Case "UFL"
                whichB5 = 0
            Case "LFL"
                whichB5 = 1
            Case "FP"
                whichB5 = 2
            Case "AIT"
                whichB5 = 3
        End Select
        B5Preference(whichB5, 0) = DBTbl("Method1")
        B5Preference(whichB5, 1) = DBTbl("Method2")
        B5Preference(whichB5, 2) = DBTbl("Method3")
        B5Preference(whichB5, 3) = DBTbl("Method4")
        B5Preference(whichB5, 4) = DBTbl("Method5")
        B5Preference(whichB5, 5) = DBTbl("Method6")
        B5Preference(whichB5, 6) = DBTbl("Method7")
        DBTbl.MoveNext
    Next i
    DBTbl.Close
after_B5_pref:
    
    'Load number formatting preferences
    Set DBTbl = DBJetUser.OpenRecordset("PrefFormatting", dbOpenTable)
    
    FormatGT1000 = DBTbl("Setting")
    DBTbl.MoveNext
    FormatLT001 = DBTbl("Setting")
    DBTbl.MoveNext
    FormatGeneral = DBTbl("Setting")
    
    DBTbl.Close
    
    'Load default units
    Set DBTbl = DBJetUser.OpenRecordset("PrefDefaultUnits", dbOpenTable)
    
'pth temp fix regaurding a problem that is occurring with new users
'    ... look into problems inside dbsave.mdb and dbuser.mdb
'    added the 'not eof' stuff in for and if statement.
    DBTbl.MoveFirst
    For i = 0 To NumProperties And Not DBTbl.EOF
        DefaultUnit(i) = DBTbl("Default Unit")
        
        DefaultTFTUnit = DBTbl("Property Name")
        
        DBTbl.MoveNext
    Next i
    
    If Not DBTbl.EOF Then
        DefaultTFTUnit = DBTbl("Default Unit")
    End If
    
    DBTbl.Close
    
End Sub


Private Sub Initialize_File_Stuff()

' REVISIONS:  DMW 6/7/97  - changed name of onechem.rpt to chemone.rpt (to make file searching easier)
    ' initializes the variables holding the files
    ' needed for the different functions
   
    default_master_name = "master.mdb"
    savefile(0) = "dbsave.mdb"
    masterfile(0) = "master.mdb"
    dbblock5file(0) = "block5.mdb"
    MasterDBName = "master.mdb"
    PathUser = App.path & "\dbuser.mdb"
    PathReport = App.path
    PathDemo = App.path & "\demo.prl"
End Sub

Public Sub Initialize_Nulls()

    SaveFileName = ""
    PathMaster = NULLPATH
    Path911 = NULLPATH
    Path801 = NULLPATH
    PathBlock5 = NULLPATH
    PathSave = NULLPATH
    
End Sub

Public Sub Check_For_Cues()

    Dim def_string As String
    Dim filename As String
    Dim FNum As Integer
    Dim i As Integer
    Dim found_cue As Boolean
    
    FNum = FreeFile
    filename = App.path & "\" & deffile
    found_cue = False
    On Error GoTo failed_cue_closed
    Open filename For Input As #FNum
    On Error GoTo failed_cue_open
    While Not EOF(FNum)
        Input #FNum, def_string
        If def_string = "Pearls.Cue" Then
            found_cue = True
            Close #FNum
            Exit Sub
        End If
    Wend
    Close #FNum
    If found_cue = False Then
        MsgBox "Run DCUT if necessary to update your database", vbOK, "Welcome"
        Open filename For Append As #FNum
        Write #FNum, "Pearls.Cue"
        Close #FNum
    End If
    
failed_cue_closed:
    Exit Sub
failed_cue_open:
    Close #FNum
    Exit Sub
End Sub

'-----------------------------------------------------------
' FUNCTION: FileExists
' Determines whether the specified file exists
'
' IN: [strPathName] - file to check for
'
' Returns: True if file exists, False otherwise
'-----------------------------------------------------------
'
Function FileExists(ByVal strPathName As String) As Integer
    Dim intFileNum As Integer

    On Error Resume Next

    '
    'Remove any trailing directory separator character
    '
    If Right$(strPathName, 1) = "\" Then
        strPathName = Left$(strPathName, Len(strPathName) - 1)
    End If

    '
    'Attempt to open the file, return value of this function is False
    'if an error occurs on open, True otherwise
    '
    intFileNum = FreeFile
    Open strPathName For Input As intFileNum

    FileExists = IIf(Err = 0, True, False)

    Close intFileNum

    Err = 0
End Function


