VERSION 5.00
Begin VB.Form frmAlias 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List of Synonyms"
   ClientHeight    =   4545
   ClientLeft      =   1245
   ClientTop       =   3885
   ClientWidth     =   7080
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtinput 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   " "
      Top             =   450
      Width           =   6850
   End
   Begin VB.CommandButton cmdSearch 
      Appearance      =   0  'Flat
      Caption         =   "&Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   810
      Width           =   6850
   End
   Begin VB.Data Data1 
      Appearance      =   0  'Flat
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4050
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton exit 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4050
      Width           =   1695
   End
   Begin VB.CommandButton cmdok 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4050
      Width           =   1695
   End
   Begin VB.ListBox vlstoutput 
      Height          =   1620
      Left            =   120
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2310
      Width           =   6825
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Please enter the string you want to search for in the synonyms database:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   180
      Width           =   6855
   End
   Begin VB.Label lblhits 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "No. of Hits:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1290
      Width           =   6855
   End
   Begin VB.Label lblinfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   $"Alias.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   1650
      Width           =   6855
   End
End
Attribute VB_Name = "frmAlias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DB_Alias As database
Dim RS_Alias As Recordset




Const frmAlias_declarations_end = True


Private Sub cmdOK_Click()
Dim SearchString As String
Dim NewString As String
Dim Q As String
Dim Name_To_Find As String
Dim Response As Integer
Dim X As Integer
Dim i As Integer
Dim CasNumber As String
  On Error GoTo err_cmdOK_Click
  Me.MousePointer = 11
  SearchString = vlstoutput.List(vlstoutput.ListIndex)
  '
  ' GET NAME OF ALIAS.
  '
  NewString = UCase$(Trim$(Right$(SearchString, Len(SearchString) - 8)))
  '
  ' GET CAS NUMBER.
  '
  CasNumber = Trim$(Left$(SearchString, 8))
  SearchString = CasNumber
  '
  ' GET ALL RECORDS THAT MATCH THIS CAS NUMBER, SORTING BY IupacName.
  '
  Set RS_Alias = DB_Alias.OpenRecordset( _
      "select * from synonyms where [cas] = " & _
      Trim$(SearchString) & _
      " and [IupacName]=true")
  RS_Alias.MoveFirst
  RS_Alias.MoveLast
  RS_Alias.MoveFirst
  '
  ' FIRST ITEM _SHOULD_ BE THE IUPAC NAME.
  '
  RS_Alias.MoveFirst
  SearchString = Database_Get_String(RS_Alias, "name")
  For i = 0 To 461
    If (Trim$(Left$(contam_prop_form.contam_combo.List(i), 8)) = CasNumber) Then
      contam_prop_form.contam_combo.ListIndex = i
      contam_prop_form.lblSelectedContaminant = contam_prop_form.contam_combo.List(i)
      contam_prop_form.cboSelectContaminant = contam_prop_form.contam_combo.List(i)
      'KAM - Changed from addtolist(i) because code was calling screens that aren't
      ' appropriate here
      Call ALias_addtolist(i)
      Exit For
    End If
  Next i
  Me.MousePointer = 0
  Unload Me
  Exit Sub
err_cmdOK_Click:
  Me.MousePointer = 0
  Call Show_Trapped_Error("cmdOK_Click")
  Resume bail_out
bail_out:
'
' OLD CODE HERE:
' ==============
'
'Dim SearchString As String, NewString As String, Q  As String
'Dim Name_To_Find As String
'Dim Response As Integer
'Dim X As Integer, i  As Integer
'
'On Error GoTo a_bad_record
'
'Me.MousePointer = 11
'
'SearchString = vlstoutput.List(vlstoutput.ListIndex)
'
''get name of alias
'NewString = UCase(Trim(Right(SearchString, Len(SearchString) - 8)))
'
''get cas number
'SearchString = Trim(Left(SearchString, 8))
'
''get all of that cas number and sort by If iupac name
'Data1.RecordSource = "select * from synonyms where [cas] = " & Trim$(SearchString) & " and [IupacName]=true"
'
'Data1.Refresh
'
' 'first item is iupac name
' Data1.Recordset.MoveFirst
' SearchString = Data1.Recordset("name")
'
'For i = 0 To 461
' If (Trim(Left(contam_prop_form.contam_combo.List(i), 8)) = CStr(Data1.Recordset("cas"))) Then
'          Data1.DatabaseName = Database_Path + "\stepp_db.mdb"
'    'ChDrive App.Path
'    'ChDir App.Path
'    Call ChangeDir_Main
'    contam_prop_form.contam_combo.ListIndex = i
'    contam_prop_form.lblSelectedContaminant = contam_prop_form.contam_combo.List(i)
'    contam_prop_form.cboSelectContaminant = contam_prop_form.contam_combo.List(i)
'
'    Call addtolist(i)
'    i = 462
' End If
'
'Next i
'
'Me.MousePointer = 0
'Unload Me
'Exit Sub
'
'a_bad_record:
'Me.MousePointer = 0
'MsgBox "Error in the database. The chemical selected does not have a iupac name"
'Resume bail_out
'
'bail_out:
End Sub

Sub ALias_addtolist(casnum As Integer)
    Dim i As Integer, J As Integer
    Dim msg$, Response As Integer

' RETURNS FALSE IF THE CHEMICAL CAN CONTINUE ON
' ALWAYS RETURNS FALSE IF NOT IN DEMOMODE CHECK DEMOMODE.BAS
    ''''If (demo_check_chemicals(contam_prop_form.contam_combo)) Then Exit Sub

    If NumSelectedChemicals = MAXSELECTEDCHEMICALS Then
       msg$ = "The maximum number of contaminants that can be selected at a time in the StEPP program is " & Str$(MAXSELECTEDCHEMICALS) & ".  Therefore, you may not select this chemical unless you Unselect a contaminant you selected previously or begin the program again."
       MsgBox msg$, MB_ICONSTOP, "Too Many Contaminants Selected"
       Exit Sub
    End If

    Screen.MousePointer = 11   'Hourglass

    Update_Fields (casnum)

    If NumSelectedChemicals = 0 Then
       contam_prop_form!mnuFile(4).Enabled = True
       contam_prop_form!mnuFile(5).Enabled = True
       contam_prop_form!mnuFile(7).Enabled = True
       contam_prop_form!cmdUnselectContaminant.Enabled = True
'''''''''       Call frmmain.frmMain_Reset_DemoVersionDisablings
    End If

    For i = 0 To contam_prop_form.cboSelectContaminant.ListCount - 1
        If Trim$(contam_prop_form.cboSelectContaminant.List(i)) = Trim$(contam_prop_form.contam_combo.List(contam_prop_form.contam_combo.ListIndex)) Then
           msg$ = "There is already a contaminant named "
           msg$ = msg$ + contam_prop_form.contam_combo.List(contam_prop_form.contam_combo.ListIndex) + " selected. "
           msg$ = msg$ + Chr$(13) + Chr$(13)
           msg$ = msg$ + "Do you wish to reinitialize it to default properties by selecting it now?"
           Response = MsgBox(msg$, MB_ICONQUESTION + MB_YESNO, "Contaminant Already Selected")
           If Response = IDYES Then
              If Trim$(contam_prop_form.contam_combo.List(contam_prop_form.contam_combo.ListIndex)) <> Trim$(contam_prop_form.cboSelectContaminant.Text) Then  'If contaminant currently selected is not the one being replaced then update its values before performing calculations
                 For J = 1 To NUMBER_OF_PROPERTIES
                     phprop.HaveProperty(J) = HaveProperty(J)
                 Next J
                 For J = 1 To NUMBER_OF_PROPERTIES_AVAILABLE
                     phprop.PROPAVAILABLE(J) = PROPAVAILABLE(J)
                 Next J
                 PropContaminant(PreviouslySelectedIndex) = phprop
              End If

              contam_prop_form.cboSelectContaminant.RemoveItem i
              For J = i + 2 To NumSelectedChemicals
                  PropContaminant(J - 1) = PropContaminant(J)
              Next J
              NumSelectedChemicals = NumSelectedChemicals - 1
              Exit For
           Else
              Screen.MousePointer = 0   'Arrow
              Exit Sub
           End If
        End If
    Next i

    contam_prop_form.cboSelectContaminant.AddItem contam_prop_form.contam_combo.List(contam_prop_form.contam_combo.ListIndex)
    contam_prop_form.lblSelectedContaminant.Caption = Trim$(dbinput.Name)

    'Update the contaminant selected prior to the new one if necessary
    If PreviouslySelectedIndex >= 0 Then
       For J = 1 To NUMBER_OF_PROPERTIES
           phprop.HaveProperty(J) = HaveProperty(J)
       Next J
       For J = 1 To NUMBER_OF_PROPERTIES_AVAILABLE
           phprop.PROPAVAILABLE(J) = PROPAVAILABLE(J)
       Next J
       PropContaminant(PreviouslySelectedIndex) = phprop
    End If

'* initialize binary interaction parameter database choices
    phprop.ActivityCoefficient.BinaryInteractionParameterDatabase = BIP_dbHierarchy.ActivityCoefficient(1)
    phprop.AqueousSolubility.BinaryInteractionParameterDatabase = BIP_dbHierarchy.AqueousSolubility(1)
    phprop.OctWaterPartCoeff.BinaryInteractionParameterDatabase = BIP_dbHierarchy.OctWaterPartCoeff(1)
    For i = 1 To 3
        phprop.ActivityCoefficient.BinaryInteractionParameterDBAvailable(i) = True
        phprop.AqueousSolubility.BinaryInteractionParameterDBAvailable(i) = True
        If i <> 3 Then phprop.OctWaterPartCoeff.BinaryInteractionParameterDBAvailable(i) = True
    Next i
    UserSelectedTheUnifacBIPDBActCoeff = False
    UserSelectedTheUnifacBIPDBAqSol = False
    UserSelectedTheUnifacBIPDBKow = False

'* Set Current Selections to None
    Call InitializeCurrentSelections

    NumSelectedChemicals = NumSelectedChemicals + 1
    Call InitializeHilights
    Call InitializePROPandHAVEAVAILABLEArrays
    Call InitializeUserInputs

    Call BlankAllTextBoxes
''''''''''    frmWaitForCalculations.Show
''''''''''    frmWaitForCalculations.Refresh

' THIS IS HERE TO MAKE SURE THAT THE CURRENT DIRECTORY IS WITH THE FORTRAN DLL
' FILES.   THE DIFFERENT *.DAT FILES THAT ARE USED BY THE FORTRAN DLLS MUST BE THERE
'    msg$ = CurDir$
'    ChDrive app.path
'    ChDir app.path + "\dlls"

    Call DoCalculationForThisContaminant

' RETURNING TO WHERE WE WERE BEFORE
'    ChDrive msg$
'    ChDir msg$

    phprop.AqueousSolubility.PreviousBinaryInteractionParameterDB = phprop.AqueousSolubility.BinaryInteractionParameterDatabase
    
    If NumSelectedChemicals > 0 Then contam_prop_form.cboSelectContaminant.Enabled = True

''''''''''    frmWaitForCalculations.Hide
    contam_prop_form.cboSelectContaminant.ListIndex = contam_prop_form.cboSelectContaminant.ListCount - 1
''''''''''    contam_prop_form.cboSelectContaminant.SetFocus
  
    Screen.MousePointer = 0   'Arrow
End Sub

Private Sub cmdSearch_Click()
Dim Encrypted_User_Name As String
Dim Encrypted_User_Password As String
Dim Response As Integer
Dim Q As String
Dim s As String

  On Error GoTo err_cmdSearch_Click
  If (Trim(txtinput) = "") Then
    Call Show_Error("You must enter a non-blank search string.")
    Exit Sub
  End If
  Me.MousePointer = 11
  '
  ' GET ALL RECORDS THAT MATCH THIS CAS NUMBER, SORTING BY IupacName.
  '
  Q = Chr$(34)
  Set RS_Alias = DB_Alias.OpenRecordset( _
      "select * from synonyms where [name] like " & Q & _
      "*" & Trim$(txtinput.Text) & "*" & Q & " order by name")
  RS_Alias.MoveFirst
  RS_Alias.MoveLast
  RS_Alias.MoveFirst
  If RS_Alias.EOF Then
    Call Show_Message("No records matching that criteria were found.")
    lblhits.Caption = "No. of Hits: 0"
    vlstoutput.Enabled = False
    Me.MousePointer = 0
    cmdok.Enabled = False
    vlstoutput.Clear
    lblinfo.Caption = "CAS #: " & Chr$(13) & "Synonym: " & Chr$(13) & " IUPAC Name: "
    Exit Sub
  End If
  vlstoutput.Enabled = True
  cmdok.Enabled = True
  vlstoutput.Clear
  RS_Alias.MoveFirst
  Do While Not RS_Alias.EOF
    s = Trim$(Str$(Database_Get_Long(RS_Alias, "cas"))) & "    " & Database_Get_String(RS_Alias, "name")
    If (Database_Get_Integer(RS_Alias, "IupacName") = True) Then
      s = s & " (*IUPAC NAME*)"
    End If
    vlstoutput.AddItem s
    RS_Alias.MoveNext
  Loop
  Me.MousePointer = 0
  lblhits.Caption = "No. of Hits: " & vlstoutput.ListCount
  Call GetNewData
  Exit Sub
  
  'DB: alias.mdb
  'table: synonyms
  'goal: search for 'text1.text'
  
err_cmdSearch_Click:
  Call Show_Trapped_Error("cmdSearch_Click")
  Resume QuitOut
QuitOut:
  Unload Me

'
' OLD CODE HERE:
' ==============
'
'Dim Encrypted_User_Name As String
'Dim Encrypted_User_Password As String
'Dim response As Integer
'Dim Q As String
'Dim s As String
'
'If (Trim(txtinput) = "") Then
'  Call Show_Error("You must enter a non-blank search string.")
'  Exit Sub
'  'response = MsgBox("This will select all possible synonyms and may take some time, Continue?", MB_ICONQUESTION + MB_YESNO, "Stepp")
'  'If (response = 7) Then
'  '  lblhits.Caption = "No. of Hits: 0"
'  '  ME.MOUSEPOINTER = 0
'  '  vlstoutput.Clear
'  '  lblinfo.Caption = "CAS #: " & Chr$(13) & "Synonym: " & Chr$(13) & " IUPAC Name: "
'  '  Exit Sub
'  'End If
'End If
'
'ME.MOUSEPOINTER = 11
'
''On Error GoTo Database_error
'On Error GoTo err_cmdSearch_Click
'
'Database_Path = App.Path + "\dbase"
'
''ChDrive Database_Path
''ChDir Database_Path
'Data1.DatabaseName = Database_Path + "\alias.mdb"
'SetDefaultWorkspace "victor t. hart", "frieda4wisc836"
'
''---
''Data1.DatabaseName = Database_Path + "\alias.mdb"
''---
''Data1.DatabaseName = Database_Path + "\alias.mdb" & ";pwd=frieda836"
''---
''Data1.DatabaseName = Database_Path + "\alias.mdb"
''Data1.Connect = "pwd=frieda836"
''---
''Set Data1.database = _
''      OpenDatabase(Database_Path + "\alias.mdb", True, False, _
''      ";pwd=" & "frieda836")
''---
''Data1.Exclusive = True
''Data1.DatabaseName = Database_Path + "\alias.mdb" & ";pwd=frieda836"
''---
''Data1.DatabaseName = Database_Path + "\alias.mdb"
''Data1.Exclusive = True
''Data1.ReadOnly = False
''Data1.Connect = "pwd=frieda836"
'
'
'Data1.Refresh
'
' 'Data1.DatabaseName = "alias.mdb"
' Q = Chr$(34)
' Data1.RecordSource = "select * from synonyms where [name] like " & Q & "*" & Trim$(txtinput.Text) & "*" & Q & " order by name"
'
' Data1.Refresh
'
' If Data1.Recordset.EOF Then
'    MsgBox "no data fitting that criteria was found"
'    lblhits.Caption = "No. of Hits: 0"
'    vlstoutput.Enabled = False
'    ME.MOUSEPOINTER = 0
'    cmdok.Enabled = False
'    vlstoutput.Clear
'    lblinfo.Caption = "CAS #: " & Chr$(13) & "Synonym: " & Chr$(13) & " IUPAC Name: "
'
'    Exit Sub
' Else
'
'  vlstoutput.Enabled = True
'  cmdok.Enabled = True
'  vlstoutput.Clear
'
'  Data1.Recordset.MoveFirst
'  Do While Not Data1.Recordset.EOF
'    s = Data1.Recordset("cas") & "    " & Data1.Recordset("name")
'    If (Data1.Recordset("IupacName") = True) Then s = s & " (*IUPAC NAME*)"
'    vlstoutput.AddItem s
'    Data1.Recordset.MoveNext
'  Loop
' End If
'
'  ME.MOUSEPOINTER = 0
'  lblhits.Caption = "No. of Hits: " & vlstoutput.ListCount
'
'  getnewdata
'
'Exit Sub
'
''DB: alias.mdb
''table: synonyms
''goal: search for 'text1.text'
'
'err_cmdSearch_Click:
'  Call Show_Trapped_Error("cmdSearch_Click")
'  Resume QuitOut
''Database_error:
''Dim temp As String, Error_Code As Integer
''    Error_Code = Err
''    temp = "Error " & Format$(Error_Code, "0") & " : " & error$(Error_Code)
''    'err.description
''    MsgBox Err.Description
''    If Err = 3024 Then
''       MsgBox "The File SYSTEM.MDA is missing.  The database is not accessible.  " & _
''       "The program will be terminated."
''    Else
''       MsgBox "Error while checking the security system.  " & _
''       Chr$(13) & temp & Chr$(13) & _
''       "The database is not accessible.  The program will be terminated."
''    End If
''
''    ME.MOUSEPOINTER = 0
''    Resume QuitOut
'QuitOut:
'   Unload Me
End Sub


Private Sub exit_Click()
  Unload Me
End Sub


Private Sub Form_Load()
  Call centerform_relative(contam_prop_form, Me)
  'NOTE: THE FOLLOWING OpenDatabase() COMMAND MUST BE
  'SPECIFIED EXACTLY AS-IS, OR ELSE IT WILL FAIL.
  Set DB_Alias = _
      Ws1.OpenDatabase(Database_Path + "\alias.mdb", _
            True, _
            False, _
            ";pwd=" & decrypt_string(Encrypted_User_Password))
End Sub
Private Sub Form_Unload(Cancel As Integer)
  'Data1.DatabaseName = Database_Path + "\stepp_db.mdb"
  'ChDrive App.Path
  'ChDir App.Path
  Call ChangeDir_Main
End Sub


Private Sub GetNewData()
Dim SearchString As String
Dim NewString As String
Dim Q As String
Dim Name_To_Find As String
Dim Response As Integer
Dim X As Integer
Dim i As Integer

  On Error GoTo err_GetNewData
  Me.MousePointer = 11
  If (vlstoutput.ListIndex = -1) Then vlstoutput.ListIndex = 0
  SearchString = vlstoutput.List(vlstoutput.ListIndex)
  '
  ' GET NAME OF ALIAS.
  '
  NewString = UCase$(Trim$(Right$(SearchString, Len(SearchString) - 8)))
  '
  ' GET CAS NUMBER.
  '
  SearchString = Trim$(Left$(SearchString, 8))
  '
  ' GET ALL RECORDS THAT MATCH THIS CAS NUMBER, SORTING BY IupacName.
  '
  Set RS_Alias = DB_Alias.OpenRecordset( _
      "select * from synonyms where [cas] = " & _
      Trim$(SearchString) & _
      " and [IupacName]=true")
  RS_Alias.MoveFirst
  RS_Alias.MoveLast
  RS_Alias.MoveFirst
  '
  ' FIRST ITEM _SHOULD_ BE THE IUPAC NAME.
  '
  RS_Alias.MoveFirst
  SearchString = Database_Get_String(RS_Alias, "name")
  If (Len(SearchString) >= 40) Then SearchString = (Left$(SearchString, 40))
  If (Len(NewString) >= 40) Then NewString = (Left$(NewString, 40))
  '
  ' DISPLAY THE INFO.
  '
  lblinfo.Caption = _
      "CAS #: " & Trim$(Str$(Database_Get_Long(RS_Alias, "cas"))) & vbCrLf & _
      "Synonym: " & LCase$(NewString) & vbCrLf & _
      " IUPAC Name: " & LCase$(CStr(SearchString))
  Me.MousePointer = 0
  Exit Sub
BailOut:
  Exit Sub
err_GetNewData:
  Me.MousePointer = 0
  Call Show_Trapped_Error("GetNewData")
  Resume BailOut


'
' OLD CODE HERE:
' ==============
'
'Dim searchstring As String, newstring As String, Q  As String
'Dim name_to_find As String
'Dim response As Integer
'Dim X As Integer, i  As Integer
'
'On Error GoTo bad_record
'
'ME.MOUSEPOINTER = 11
'
'If (vlstoutput.ListIndex = -1) Then vlstoutput.ListIndex = 0
'
'searchstring = vlstoutput.List(vlstoutput.ListIndex)
'
''get name of alias
'newstring = UCase(Trim(Right(searchstring, Len(searchstring) - 8)))
'
''get cas number
'searchstring = Trim(Left(searchstring, 8))
'
''get all of that cas number and sort by If iupac name
'Data1.RecordSource = "select * from synonyms where [cas] = " & Trim$(searchstring) & " and [IupacName]=true"
'
'Data1.Refresh
'
' 'first item is iupac name
' Data1.Recordset.MoveFirst
' searchstring = Data1.Recordset("name")
'
' If (Len(searchstring) >= 40) Then searchstring = (Left(searchstring, 40))
' If (Len(newstring) >= 40) Then newstring = (Left(newstring, 40))
'
' 'remove and replace with code starting stepp run
'  lblinfo.Caption = "CAS #: " & Data1.Recordset("cas") & Chr$(13) & "Synonym: " & newstring & Chr$(13) & " IUPAC Name: " & UCase(CStr(searchstring))
'ME.MOUSEPOINTER = 0
'Exit Sub
'
'bad_record:
'ME.MOUSEPOINTER = 0
'MsgBox "Error in the database. The chemical selected does not have a iupac name"
'Resume bailout
'
'bailout:
End Sub


Private Sub txtinput_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call cmdSearch_Click
  End If
End Sub


Private Sub vlstoutput_Click()
  Call GetNewData
End Sub


Private Sub vlstoutput_DblClick()
  Call cmdOK_Click
End Sub


