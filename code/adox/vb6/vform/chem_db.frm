VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmChemDB 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Chemical DataBase"
   ClientHeight    =   7305
   ClientLeft      =   495
   ClientTop       =   1260
   ClientWidth     =   11295
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   11295
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Search for Chemical"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   4995
      Begin VB.ComboBox cboSearchType 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   360
         Width           =   3555
      End
      Begin VB.TextBox txtDataStr 
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Text            =   "txtDataStr(0)"
         Top             =   960
         Width           =   4725
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   5
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lbl_cboIonType 
         Caption         =   "Search Order"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label1 
         Caption         =   "Search String"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   7770
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   1725
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   5265
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   4995
      _Version        =   65536
      _ExtentX        =   8811
      _ExtentY        =   9287
      _StockProps     =   14
      Caption         =   "Select"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ListBox lstRecords 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4620
         Left            =   240
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   360
         Width           =   4545
      End
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Use These Values"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   5730
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   90
      Width           =   1785
   End
   Begin Threed.SSFrame fraRecord 
      Height          =   6255
      Left            =   5220
      TabIndex        =   10
      Top             =   810
      Width           =   5145
      _Version        =   65536
      _ExtentX        =   9075
      _ExtentY        =   11033
      _StockProps     =   14
      Caption         =   "Values from Database"
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtData 
         Enabled         =   0   'False
         Height          =   300
         Index           =   3
         Left            =   1530
         TabIndex        =   14
         Text            =   "txtData(3)"
         Top             =   2835
         Width           =   1095
      End
      Begin VB.TextBox txtData 
         Enabled         =   0   'False
         Height          =   300
         Index           =   2
         Left            =   1530
         TabIndex        =   13
         Text            =   "txtData(2)"
         Top             =   2235
         Width           =   1095
      End
      Begin VB.TextBox txtData 
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   1530
         TabIndex        =   12
         Text            =   "txtData(1)"
         Top             =   1605
         Width           =   1695
      End
      Begin VB.TextBox txtDataStr 
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   1530
         MaxLength       =   250
         TabIndex        =   11
         Text            =   "txtDataStr(1)"
         Top             =   1080
         Width           =   3255
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Caption         =   "Second Order Rate Constant"
         Height          =   585
         Index           =   5
         Left            =   270
         TabIndex        =   20
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label lbldesc2 
         Caption         =   "L/gmol-s"
         Height          =   375
         Index           =   5
         Left            =   2730
         TabIndex        =   19
         Top             =   2850
         Width           =   975
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Caption         =   "Molecular Weight"
         Height          =   375
         Index           =   2
         Left            =   60
         TabIndex        =   18
         Top             =   2280
         Width           =   1425
      End
      Begin VB.Label lbldesc2 
         Caption         =   "g/gmol"
         Height          =   375
         Index           =   2
         Left            =   2730
         TabIndex        =   17
         Top             =   2250
         Width           =   855
      End
      Begin VB.Label lblData 
         Alignment       =   1  'Right Justify
         Caption         =   "CAS Number"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   16
         Top             =   1650
         Width           =   1125
      End
      Begin VB.Label lblDataStr 
         Alignment       =   1  'Right Justify
         Caption         =   "Chemical name"
         Height          =   375
         Index           =   0
         Left            =   180
         TabIndex        =   15
         Top             =   1140
         Width           =   1305
      End
   End
End
Attribute VB_Name = "frmChemDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iFormMode As Integer
Const iFormMode_IMPORT = 1
Const iFormMode_EDIT = 2

Dim USER_HIT_CANCEL As Boolean
Dim Db1 As Database
Public HALT_ALL_CONTROLS As Boolean
Dim strSearchType_FORCE As String
Dim fn_database
Dim fn_synonym
Dim NowTargetCompound As TargetCompound_Type
Dim TempProj As Project_Type
Dim tcNum As Integer

Const frmChemDB_decl_end = True


Sub frmChemDB_IMPORT_MODE(IN_strSearchType_FORCE As String)
  
  'IMPORT THIS PROJECT FROM MEMORY TO THE FORM.
  TempProj = NowProj
    
  iFormMode = iFormMode_IMPORT
  strSearchType_FORCE = IN_strSearchType_FORCE
  frmChemDB.Show 1
'  out_HitCancel = USER_HIT_CANCEL
   'UPDATE MEMORY.
  If (Not USER_HIT_CANCEL) Then
    TempProj.TargetCompounds(TempProj.TargetCompounds_Count) = NowTargetCompound
    NowProj = TempProj

    'RETURN TO MAIN WINDOW.
'    Call DirtyFlag_Throw(TempProj)
    Call refresh_frmTarget(NowTargetCompound)
  End If

End Sub
Sub frmChemDB_EDIT_MODE(IN_strSearchType_FORCE As String)

  TempProj = NowProj
  iFormMode = iFormMode_EDIT
  strSearchType_FORCE = IN_strSearchType_FORCE
  frmChemDB.Show 1
  If (Not USER_HIT_CANCEL) Then
    TempProj.TargetCompounds(frmMain.cboTarget.ListIndex + 1) = NowTargetCompound
    NowProj = TempProj
'    Call DirtyFlag_Throw(TempProj)
    Call refresh_frmTarget(NowTargetCompound)
  End If

End Sub


Function IsRecordSelected() As Boolean
  If (lstRecords.ListCount <= 0) Or (lstRecords.ListIndex < 0) Then
    Call Show_Error("You must first select a record.")
    IsRecordSelected = False
  Else
    IsRecordSelected = True
  End If
End Function
Sub SetRecordDefaults(IsBlank As Boolean)
  With NowTargetCompound
    If (IsBlank = True) Then
'      .strSearchType = cboSearchType.List(cboSearchType.ListIndex)
      .comname = ""
      .cas = 0
      .mw = 0
      .xk = 0
'''    Else
''''      .strSearchType = cboSearchType.List(cboSearchType.ListIndex)
'''      .comname =
'''      .cas = 0
'''      .mw = 0
'''      .xk = 0
    End If
  End With
End Sub


Sub Populate_cboSearchType(tcNum)
Dim i As Integer
  With cboSearchType
    Me.HALT_ALL_CONTROLS = True
    .AddItem "Chemical name"
    .AddItem "Synonyms"
    .AddItem "CAS number"
    .Locked = False
    
    If (TempProj.TargetCompounds(tcNum).cas <> 0 And _
      TempProj.TargetCompounds(tcNum).comname = "") Then
      .ListIndex = 2
    Else
      .ListIndex = 0
    End If

    Me.HALT_ALL_CONTROLS = False
  End With
End Sub


Sub Populate_frmChemDB_Units()
Dim Frm As Form
Set Frm = Me
  '------------------------------------------------------------------------------------------------------------------------
  '
  ' PART ONE.
  '
'  Call unitsys_register(Frm, lblData(0), txtData(0), cboUnits(0), "dimensionless", _
'      "dim'less", "dim'less", "", "", 100#, True)
End Sub


Function Populate_lstRecords(booSearch As Boolean, strSearchString As String) _
    As Boolean
  On Error GoTo err_ThisFunc
Dim Rs1 As Recordset
Dim strSearchType As String
Dim strChemicalName As String
Dim strSynonymName As String
Dim lngCASNumber As Long
Dim strMW As String
Dim strXK As String
Dim strSearchCriteria As String
Dim strSearchType_USE_THIS As String
Dim Q As String


  lstRecords.Clear
  strSearchType_USE_THIS = cboSearchType.List(cboSearchType.ListIndex)
  Q = Chr$(34)    '' quote
  
   Select Case strSearchType_USE_THIS
    Case "Chemical name"
      fn_database = fn_DB_dir & "\adox_rate_database.mdb"
      Call CHEMDB_MDB_Open(Db1)
      If booSearch Then
         strSearchCriteria = _
         "select * from rate_const " & _
         "where name like " & Q & _
         "*" & Trim$(strSearchString) & "*" & Q & _
         " order by name"
      Else
        strSearchCriteria = _
         "select * from rate_const " & _
         "order by name"
      End If
    Case "Synonyms"
      fn_database = fn_DB_dir & "\synonyms.mdb"
      Call CHEMDB_MDB_Open(Db1)
      If booSearch Then
        strSearchCriteria = _
         "select * from synonyms " & _
         "where name like " & Q & "*" & Trim$(strSearchString) & _
         "*" & Q & " order by name"
      Else
'        strSearchCriteria = _
'         "select * from [synonyms] " & _
'         "order by name"
      End If
    Case "CAS number"
      fn_database = fn_DB_dir & "\adox_rate_database.mdb"
      Call CHEMDB_MDB_Open(Db1)
      If booSearch Then
        strSearchCriteria = _
          "select * from [rate_const] " & _
          "where [CAS] like " & val(Trim$(strSearchString)) & _
          " order by CAS"
      Else
        strSearchCriteria = _
          "select * from rate_const " & _
          "order by CAS"
      End If
   End Select

  
  If (False = Database_TestForCriteria( _
    Db1, _
    Rs1, _
    strSearchCriteria)) Then
    GoTo exit_err_ThisFunc
  End If
  Do While (Not Rs1.EOF)
    Select Case strSearchType_USE_THIS
    Case "Chemical name"
        strChemicalName = Database_Get_String(Rs1, "Name")
        lngCASNumber = Database_Get_Long(Rs1, "CAS")
        lstRecords.AddItem strChemicalName & " : " & Str(lngCASNumber)
    Case "Synonyms"
        strSynonymName = Database_Get_String(Rs1, "name")
        lngCASNumber = Database_Get_Long(Rs1, "cas")
        lstRecords.AddItem strSynonymName & " : " & Str(lngCASNumber)
    Case "CAS number"
        strChemicalName = Database_Get_String(Rs1, "Name")
        lngCASNumber = Database_Get_Long(Rs1, "CAS")
        lstRecords.AddItem Str(lngCASNumber) & " : " & strChemicalName
    End Select
    lstRecords.ItemData(lstRecords.NewIndex) = lngCASNumber
    Rs1.MoveNext
  Loop
  Rs1.Close
exit_normally_ThisFunc:
  Populate_lstRecords = True
  Exit Function
exit_err_ThisFunc:
  Populate_lstRecords = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("Populate_lstRecords")
  Resume exit_err_ThisFunc
End Function


Function Populate_RecordData() _
    As Boolean
  On Error GoTo err_ThisFunc
Dim Rs1 As Recordset
Dim lngCASNumber As Long
Dim strRecordCriteria As String
Dim strSearchType_USE_THIS As String

  fn_database = fn_DB_dir & "\adox_rate_database.mdb"
      Call CHEMDB_MDB_Open(Db1)
  strSearchType_USE_THIS = cboSearchType.List(cboSearchType.ListIndex)
  If (lstRecords.ListCount <= 0) Or _
      (lstRecords.ListIndex < 0) Then
    GoTo exit_err_ThisFunc
  End If
  lngCASNumber = lstRecords.ItemData(lstRecords.ListIndex)
  
  If (False = Database_TestForCriteria( _
    Db1, _
    Rs1, _
    "select * from [rate_const] " & _
    "where [CAS]=" & lngCASNumber)) Then
    GoTo exit_err_ThisFunc
  End If
  With NowTargetCompound
    .comname = Database_Get_String(Rs1, "Name")
    .cas = Database_Get_Long(Rs1, "CAS")
    .mw = Database_Get_Long(Rs1, "Molecular Weight")
    .xk = Database_Get_Double(Rs1, "RateConstant_OH")
    .dep_mw = NowTargetCompound.mw + NowTargetCompound.dep_val
  End With
  Call refresh_frmChemDB(NowTargetCompound)
  Rs1.Close
exit_normally_ThisFunc:
  Populate_RecordData = True
  Exit Function
exit_err_ThisFunc:
  '
  ' DISPLAY BLANK DATA.
  '
  Call SetRecordDefaults(True)
  Call refresh_frmChemDB(NowTargetCompound)
  Populate_RecordData = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("Populate_RecordData")
  Resume exit_err_ThisFunc
End Function


Private Sub cboSearchType_Click()
  '
  ' POPULATE RECORDS FROM DATABASE, AND INITIAL (BLANK) RECORD DATA.
  '
  Select Case cboSearchType.ListIndex
    Case 0
      Call Populate_lstRecords(True, TempProj.TargetCompounds(tcNum).comname)
    Case 1
      Call Populate_lstRecords(False, "")
    Case 2
      Call Populate_lstRecords(True, Str(TempProj.TargetCompounds(tcNum).cas))
   End Select
  Call Populate_RecordData
  Call refresh_frmChemDB(NowTargetCompound)
End Sub


'Private Sub cboUnits_Click(Index As Integer)
'Dim Ctl As Control
'Set Ctl = cboUnits(Index)
'  Call unitsys_control_cbox_click(Ctl)
'  Call refresh_frmChemDB(frmChemDB_Record)
'End Sub
Private Sub cboUnits_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = Global_TextKeyPress(KeyAscii)
End Sub


Private Sub cmdButton_Click(Index As Integer)
Dim BadName As Boolean
Dim sName_SAVE As String
Dim strThisCompName As String
Dim strMsg As String
  Select Case Index
    '
    '////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////
    Case 0:       'IMPORT.
      If (IsRecordSelected() = False) Then
        Exit Sub
      End If
      '
      ' EXIT OUT OF HERE.
      USER_HIT_CANCEL = False
      Unload Me
      Exit Sub
    '
    '////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////
    Case 1:       'CLOSE/EXIT.
      USER_HIT_CANCEL = True
      Unload Me
      Exit Sub
  End Select
End Sub



Private Sub cmdSearch_Click()
  Dim strSearchString As String

'  If (Trim(txtDataStr(0).Text) = "") Then
'    Call Show_Error("You must enter a non-blank search string.")
'    Exit Sub
'  End If
  strSearchString = txtDataStr(0).Text
  Screen.MousePointer = vbHourglass
  Select Case strSearchString = ""
    Case False
      Call Populate_lstRecords(True, strSearchString)
    Case True
      Call Populate_lstRecords(False, "")
  End Select
  Call Populate_RecordData
  Screen.MousePointer = vbDefault
  Call refresh_frmChemDB(NowTargetCompound)

End Sub


Private Sub Form_Load()
  '
  ' ------- OPEN THE DATABASE. --------------------------------------
  '
  Set Ws1 = Workspaces(0)
  fn_DB_dir = App.Path & "\dbase"
  fn_database = fn_DB_dir & "\adox_rate_database.mdb"
  
  
  Call CHEMDB_MDB_Open(Db1)
  '
  ' MISC INITS.
  '
  Call CenterOnScreen(Me)
'  Call refresh_frmChemDB(frmChemDB_Record)_FirstFormLoad
'  Call Global_GotFocus_ResetColors("ALL_WHITE")
  Select Case iFormMode
    Case iFormMode_IMPORT:
      tcNum = TempProj.TargetCompounds_Count
    Case iFormMode_EDIT:
      tcNum = frmMain.cboTarget.ListIndex + 1
  End Select
      'IMPORT THIS TARGET COMPOUND FROM MEMORY TO THE FORM.
  NowTargetCompound = TempProj.TargetCompounds(tcNum)

  cmdButton(0).visible = True
  cmdButton(1).Caption = "&Cancel"
  
  Call Populate_cboSearchType(tcNum)
  
  frmChemDB_Record.DB_Mode = DB_Mode_VIEW
  '
  ' POPULATE RECORDS FROM DATABASE
  '
  txtDataStr(0).Text = TempProj.TargetCompounds(tcNum).comname
  txtDataStr(0).Tag = TempProj.TargetCompounds(tcNum).comname
  Select Case iFormMode
    Case iFormMode_IMPORT:
      Call Populate_lstRecords(True, TempProj.TargetCompounds( _
        tcNum).comname)
    Case iFormMode_EDIT:
      If TempProj.TargetCompounds( _
        tcNum).comname = "" Then
          Call Populate_lstRecords(False, "")
      Else
        Call Populate_lstRecords(True, TempProj.TargetCompounds( _
        tcNum).comname)
      End If
  End Select
  Call Populate_RecordData
'  Call refresh_frmChemDB(frmChemDB_Record)
End Sub
Private Sub Form_Unload(Cancel As Integer)
  '
  ' DEREGISTER UNIT CONTROLS.
  '
  Call unitsys_unregister_all_on_form(Me)
  '
  ' ------- CLOSE THE DATABASE. --------------------------------------
  '
  Call CHEMDB_MDB_Close(Db1)
  '
  ' MISC UNLOAD STUFF.
  '
'  Call Global_GotFocus_ResetColors("NORMAL")
End Sub


Private Sub lstRecords_Click()
  Call Populate_RecordData
End Sub


'Private Sub mnuRecordItem_Click(Index As Integer)
'  On Error GoTo err_ThisFunc
'Dim sThisName As String
'Dim sMsg As String
'Dim RetVal As Integer
'Dim lngRecID As Long
'Dim Rs1 As Recordset
'  Select Case Index
'    '
'    '////////////////////////////////////////////////////////////////////////
'    '////////////////////////////////////////////////////////////////////////
'    Case 10:      'NEW.
'      frmChemDB_Record.DB_Mode = DB_Mode_ADDNEW
'      '
'      ' SET DEFAULTS.
'      Call SetRecordDefaults(False)
'      Call refresh_frmChemDB(frmChemDB_Record)
'    '
'    '////////////////////////////////////////////////////////////////////////
'    '////////////////////////////////////////////////////////////////////////
'    Case 20:      'EDIT.
'      If (IsRecordSelected() = False) Then
'        GoTo exit_err_ThisFunc
'      End If
'      frmChemDB_Record.DB_Mode = DB_Mode_EDIT
'      Call refresh_frmChemDB(frmChemDB_Record)
'    '
'    '////////////////////////////////////////////////////////////////////////
'    '////////////////////////////////////////////////////////////////////////
'    Case 30:      'DELETE.
'      If (IsRecordSelected() = False) Then
'        GoTo exit_err_ThisFunc
'      End If
'      sThisName = Trim$(lstRecords.List(lstRecords.ListIndex))
'      sMsg = "Do you really want to delete record '" & _
'          sThisName & "' from the database ?"
'      RetVal = MsgBox(sMsg, vbCritical + vbYesNo, AppName_For_Display_Long)
'      If (RetVal = vbNo) Then Exit Sub
'      '
'      ' DELETE THIS RECORD.
'      '
'      lngRecID = lstRecords.ItemData(lstRecords.ListIndex)
'      If (False = Database_TestForCriteria( _
'        db1, _
'        Rs1, _
'        "select * from [DB_EqParams] " & _
'        "where [lngRecId]=" & Trim$(Str$(lngRecID)))) Then
'        GoTo exit_err_ThisFunc
'      End If
'      Rs1.Delete
'      Rs1.Close
'      '
'      ' REFRESH THE WINDOW.
'      '
'      Call SetRecordDefaults(True)
'      Call Populate_lstRecords
'      Call Populate_RecordData
'      Call refresh_frmChemDB(frmChemDB_Record)
'    '
'    '////////////////////////////////////////////////////////////////////////
'    '////////////////////////////////////////////////////////////////////////
'    Case 50:      'SAVE CHANGES.
'      With frmChemDB_Record
'        If (.DB_Mode = DB_Mode_ADDNEW) Then
'          ' START TO ADD NEW RECORD.
'          Set Rs1 = db1.OpenRecordset("select * from [ADOX_rate_database]")
'          Rs1.AddNew
'        Else
'          ' LOOK UP AND START EDITING EXISTING RECORD.
'          lngRecID = lstRecords.ItemData(lstRecords.ListIndex)
'          If (False = Database_TestForCriteria( _
'            db1, _
'            Rs1, _
'            "select * from [DB_EqParams] " & _
'            "where [lngRecId]=" & Trim$(Str$(lngRecID)))) Then
'            GoTo exit_err_ThisFunc
'          End If
'          Rs1.Edit
'        End If
'        '
'        ' STORE THE RECORD.
'        '
'        Rs1("strSearchType") = .strSearchType
'        Rs1("strChemicalNameName") = .strChemicalName
'        Rs1("dblEqValue") = .dblEqValue
'        Rs1("strPresaturantIon") = .strPresaturantIon
'        Rs1("strEqValueSource") = .strEqValueSource
'        Rs1("strResinName") = .strResinName
'        Rs1("strResinType") = .strResinType
'        Rs1("strResinManufacturer") = .strResinManufacturer
'        '
'        ' COMPLETE THE STORAGE.
'        '
'        Rs1.Update
'        Rs1.Close
'        If (.DB_Mode = DB_Mode_ADDNEW) Then
'          ' DO NOTHING.
'        Else
'          ' DO NOTHING.
'        End If
'      End With
'      '
'      ' REFRESH WINDOW.
'      '
'      frmChemDB_Record.DB_Mode = DB_Mode_VIEW
'      Call SetRecordDefaults(True)
'      Call Populate_lstRecords
'      Call Populate_RecordData
'      Call refresh_frmChemDB(frmChemDB_Record)
'    '
'    '////////////////////////////////////////////////////////////////////////
'    '////////////////////////////////////////////////////////////////////////
'    Case 60:      'CANCEL CHANGES.
'      frmChemDB_Record.DB_Mode = DB_Mode_VIEW
'      Call Populate_RecordData
'      Call refresh_frmChemDB(frmChemDB_Record)
'  End Select
'exit_normally_ThisFunc:
'  'mnuRecordItem_Click = True
'  Exit Sub
'exit_err_ThisFunc:
'  'mnuRecordItem_Click = False
'  Exit Sub
'err_ThisFunc:
'  Call Show_Trapped_Error("mnuRecordItem_Click")
'  Resume exit_err_ThisFunc
'End Sub




Private Sub txtdata_GotFocus(Index As Integer)
Dim ctl As Control
Set ctl = txtData(Index)
Dim StatusMessagePanel As String
  If (ctl.Locked = True) Then Exit Sub
  If (frmChemDB_Record.DB_Mode <> DB_Mode_VIEW) Then
    Call unitsys_control_txtx_gotfocus(ctl)
  End If
End Sub
Private Sub txtdata_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = Global_NumericKeyPress(KeyAscii)
End Sub
Private Sub txtdata_LostFocus(Index As Integer)
Dim NewValue_Okay As Integer
Dim NewValue As Double
Dim ctl As Control
Set ctl = txtData(Index)
Dim Val_Low As Double
Dim Val_High As Double
Dim Raise_Dirty_Flag As Boolean
Dim Too_Small As Integer
Dim SetNoVal As Boolean
  If (ctl.Locked = True) Then
    Call refresh_frmChemDB(NowTargetCompound)
    Exit Sub
  End If
  'NOTE: LOW AND HIGH VALUES IN BASE UNITS
  Select Case Index
    Case Else:
      Val_Low = -1E+20: Val_High = 1E+20
  End Select
  NewValue_Okay = False
  If (Trim$(ctl.Text) = "") Then
    SetNoVal = True
    NewValue_Okay = True
    Raise_Dirty_Flag = True
  Else
    SetNoVal = False
    If (unitsys_control_txtx_lostfocus_validate( _
        ctl, _
        Val_Low, _
        Val_High, _
        NewValue, _
        Raise_Dirty_Flag)) Then
      NewValue_Okay = True
    End If
  End If
  Call unitsys_control_txtx_lostfocus(ctl, NewValue)
''''  Call Local_GenericStatus_Set("")
  If (NewValue_Okay) Then
    If (Raise_Dirty_Flag) Then
      '
      ' STORE TO MEMORY.
      '
''''      With NowTargetCompound
''''        Select Case Index
''''
''''          Case 0: .Duct(.iDuct_Displayed).dblDiameter = IIf(SetNoVal, NoValue_dbl, NewValue)
''''          Case 1: .Duct(.iDuct_Displayed).dblDepth = IIf(SetNoVal, NoValue_dbl, NewValue)
''''          Case 2: .Duct(.iDuct_Displayed).iNumPortsForDuct = IIf(SetNoVal, NoValue_i, CInt(NewValue))
''''          Case 3: '''' OUTPUT ONLY.
''''        End Select
''''      End With
    End If
  End If
  Call refresh_frmChemDB(NowTargetCompound)
End Sub

Function CHEMDB_MDB_Open( _
    Db1 As Database) _
    As Boolean
On Error GoTo err_ThisFunc
'  Set Db1 = OpenDatabase(fn_database)
  Set Db1 = OpenDatabase(fn_database, _
                True, _
                False, _
                ";pwd=" & Encrypted_User_Password)
exit_normally_ThisFunc:
  CHEMDB_MDB_Open = True
  Exit Function
exit_err_ThisFunc:
  CHEMDB_MDB_Open = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("CHEMDB_MDB_Open")
  Resume exit_err_ThisFunc
End Function
Function CHEMDB_MDB_Close( _
    Db1 As Database) _
    As Boolean
On Error GoTo err_ThisFunc
  Db1.Close
exit_normally_ThisFunc:
  CHEMDB_MDB_Close = True
  Exit Function
exit_err_ThisFunc:
  CHEMDB_MDB_Close = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("CHEMDB_MDB_Close")
  Resume exit_err_ThisFunc
End Function



Public Function frmChemDB_DoEdit(tcNum As Integer) As Integer

Dim is_aborted As Integer
Dim name_new As String

  'IMPORT THIS PROJECT FROM MEMORY TO THE FORM.
  TempProj = NowProj
  
  'SHOW THE FORM.
  frmChemDB.Show 1
  
  'UPDATE MEMORY.
  If (Not USER_HIT_CANCEL) Then
    NowProj = TempProj
  End If
  'RETURN TO MAIN WINDOW.
  frmChemDB_DoEdit = Not USER_HIT_CANCEL

  
End Function

Private Sub txtDataStr_GotFocus(Index As Integer)
  Dim txtctl As Control
  Set txtctl = txtDataStr(Index)
  Call DisplayDataEntryError
  Call Global_GotFocus(txtctl)
End Sub
Private Sub txtDataStr_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = Global_TextKeyPress(KeyAscii)
End Sub
Private Sub txtDataStr_LostFocus(Index As Integer)
  Dim strSearchString As String
  Dim txtctl As Control
  Set txtctl = txtDataStr(Index)
  Dim ok_to_save As Integer
  Dim refresh_type As Integer
  
  ok_to_save = False
  If (txtctl.Text <> txtctl.Tag) Then
    ok_to_save = True
  End If
  If (ok_to_save) Then
    'DATA LOOKS OKAY, LET'S GO AHEAD AND SAVE IT.
    refresh_type = 1
    Select Case Index
      Case 0: strSearchString = Trim$(txtctl.Text)
    End Select
    
    Call AssignTextAndTag(txtctl, txtctl.Text)
    
    'THROW DIRTY FLAG, AND REFRESH EVERY WINDOWS.
'    Call DirtyFlag_Throw(TempProj)
    
    Select Case refresh_type
      Case 1:   'JUST THE PHOTOCAT WINDOW.
'        Call refresh_frmPhotoChem(TempProj)
    End Select
  End If
  Call Global_LostFocus(txtctl)
End Sub
Sub Local_DirtyStatus_Set(DirtyFlag As Boolean, NewSetting As Boolean)
  Call Global_DirtyStatus_Set(Me, DirtyFlag, NewSetting)
End Sub

Sub Local_GenericStatus_Set(NewString As String)
  Call Global_GenericStatus_Set(Me, NewString)
End Sub
