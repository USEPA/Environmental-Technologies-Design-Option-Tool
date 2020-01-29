VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmDB_EqParams 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Equilibrium Parameters Database"
   ClientHeight    =   7305
   ClientLeft      =   495
   ClientTop       =   1545
   ClientWidth     =   11295
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   11295
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdButton 
      Caption         =   "{Cancel/Exit}"
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
      Left            =   7980
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   6450
      Width           =   3165
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   6945
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3405
      _Version        =   65536
      _ExtentX        =   6006
      _ExtentY        =   12250
      _StockProps     =   14
      Caption         =   "Select a Record:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ComboBox cboIonType 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1170
         Style           =   2  'Dropdown List
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   360
         Width           =   2115
      End
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
         Height          =   5580
         Left            =   120
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   930
         Width           =   3165
      End
      Begin VB.Label lbl_cboIonType 
         Alignment       =   1  'Right Justify
         Caption         =   "Ion Type:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   420
         Width           =   945
      End
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Impor&t This Record ..."
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
      Left            =   3630
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   6450
      Width           =   4335
   End
   Begin Threed.SSFrame fraRecord 
      Height          =   5895
      Left            =   3510
      TabIndex        =   5
      Top             =   120
      Width           =   7635
      _Version        =   65536
      _ExtentX        =   13467
      _ExtentY        =   10398
      _StockProps     =   14
      Caption         =   "{caption set by code}"
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
      Begin VB.TextBox txtDataStr 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   2160
         MaxLength       =   250
         TabIndex        =   20
         Text            =   "txtDataStr()"
         Top             =   3120
         Width           =   4365
      End
      Begin VB.TextBox txtDataStr 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   2160
         MaxLength       =   250
         TabIndex        =   18
         Text            =   "txtDataStr()"
         Top             =   2640
         Width           =   4365
      End
      Begin VB.TextBox txtDataStr 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   2160
         MaxLength       =   250
         TabIndex        =   16
         Text            =   "txtDataStr()"
         Top             =   2190
         Width           =   4365
      End
      Begin VB.TextBox txtDataStr 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   2160
         MaxLength       =   250
         TabIndex        =   14
         Text            =   "txtDataStr()"
         Top             =   1740
         Width           =   4365
      End
      Begin VB.TextBox txtDataStr 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   2160
         MaxLength       =   250
         TabIndex        =   12
         Text            =   "txtDataStr()"
         Top             =   1290
         Width           =   4365
      End
      Begin VB.ComboBox cboUnits 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   3810
         Style           =   2  'Dropdown List
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   825
         Width           =   1545
      End
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   2160
         TabIndex        =   1
         Text            =   "txtData()"
         Top             =   840
         Width           =   1600
      End
      Begin VB.TextBox txtDataStr 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   2160
         MaxLength       =   250
         TabIndex        =   0
         Text            =   "txtDataStr()"
         Top             =   390
         Width           =   3165
      End
      Begin VB.Label lblDataStr 
         Alignment       =   1  'Right Justify
         Caption         =   "Resin Manufacturer:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   240
         TabIndex        =   21
         Top             =   3165
         Width           =   1845
      End
      Begin VB.Label lblDataStr 
         Alignment       =   1  'Right Justify
         Caption         =   "Resin Type:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   240
         TabIndex        =   19
         Top             =   2685
         Width           =   1845
      End
      Begin VB.Label lblDataStr 
         Alignment       =   1  'Right Justify
         Caption         =   "Resin Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   240
         TabIndex        =   17
         Top             =   2235
         Width           =   1845
      End
      Begin VB.Label lblDataStr 
         Alignment       =   1  'Right Justify
         Caption         =   "Source of Value:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   240
         TabIndex        =   15
         Top             =   1785
         Width           =   1845
      End
      Begin VB.Label lblDataStr 
         Alignment       =   1  'Right Justify
         Caption         =   "Presaturant Ion:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   13
         Top             =   1335
         Width           =   1845
      End
      Begin VB.Label lblData 
         Alignment       =   1  'Right Justify
         Caption         =   "Parameter Value:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   885
         Width           =   1845
      End
      Begin VB.Label lblDataStr 
         Alignment       =   1  'Right Justify
         Caption         =   "Ion Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   435
         Width           =   1845
      End
   End
   Begin VB.Menu mnuRecord 
      Caption         =   "&Record"
      Begin VB.Menu mnuRecordItem 
         Caption         =   "&Add ..."
         Index           =   10
      End
      Begin VB.Menu mnuRecordItem 
         Caption         =   "&Edit ..."
         Index           =   20
      End
      Begin VB.Menu mnuRecordItem 
         Caption         =   "&Delete ..."
         Index           =   30
      End
      Begin VB.Menu mnuRecordItem 
         Caption         =   "-"
         Index           =   40
      End
      Begin VB.Menu mnuRecordItem 
         Caption         =   "&Save Changes ..."
         Index           =   50
      End
      Begin VB.Menu mnuRecordItem 
         Caption         =   "&Cancel Changes ..."
         Index           =   60
      End
   End
End
Attribute VB_Name = "frmDB_EqParams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iFormMode As Integer
Const iFormMode_IMPORT = 1
Const iFormMode_EDIT = 2

Dim USER_HIT_CANCEL As Boolean
Dim db1 As Database
Public HALT_ALL_CONTROLS As Boolean
Dim strIonType_FORCE As String






Const frmDB_EqParams_decl_end = True


Sub frmDB_EqParams_IMPORT_MODE( _
    out_HitCancel As Boolean, _
    IN_strIonType_FORCE As String)
  iFormMode = iFormMode_IMPORT
  strIonType_FORCE = IN_strIonType_FORCE
  frmDB_EqParams.Show 1
  out_HitCancel = USER_HIT_CANCEL
End Sub
Sub frmDB_EqParams_EDIT_MODE( _
    )
  iFormMode = iFormMode_EDIT
  strIonType_FORCE = ""
  frmDB_EqParams.Show 1
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
  With frmDB_EqParams_Record
    If (IsBlank = True) Then
      .strIonType = cboIonType.List(cboIonType.ListIndex)
      .strIonName = ""
      .dblEqValue = 0#
      .strPresaturantIon = ""
      .strEqValueSource = ""
      .strResinName = ""
      .strResinType = ""
      .strResinManufacturer = ""
    Else
      .strIonType = cboIonType.List(cboIonType.ListIndex)
      .strIonName = "New " & cboIonType.List(cboIonType.ListIndex)
      .dblEqValue = 1#
      .strPresaturantIon = ""
      .strEqValueSource = "n/a"
      .strResinName = "n/a"
      .strResinType = "n/a"
      .strResinManufacturer = "n/a"
    End If
  End With
End Sub


Sub Populate_cboIonType()
Dim i As Integer
  With cboIonType
    Me.HALT_ALL_CONTROLS = True
    .Clear
    .AddItem "Anion"
    .AddItem "Cation"
    .Locked = False
    If (strIonType_FORCE <> "") Then
      For i = 0 To .ListCount - 1
        If (Trim$(UCase$(.List(i))) = Trim$(UCase$(strIonType_FORCE))) Then
          .ListIndex = i
          .Locked = True
          Exit For
        End If
      Next i
    Else
      .ListIndex = 0
    End If
    Me.HALT_ALL_CONTROLS = False
  End With
End Sub


Sub Populate_frmDB_EqParams_Units()
Dim Frm As Form
Set Frm = Me
  '------------------------------------------------------------------------------------------------------------------------
  '
  ' PART ONE.
  '
  Call unitsys_register(Frm, lblData(0), txtData(0), cboUnits(0), "dimensionless", _
      "dim'less", "dim'less", "", "", 100#, True)
End Sub


Function Populate_lstRecords() _
    As Boolean
  On Error GoTo err_ThisFunc
Dim Rs1 As Recordset
Dim strResinName As String
Dim strIonType As String
Dim strIonName As String
Dim lngRecID As Long
Dim strSearchCriteria As String
Dim strIonType_USE_THIS As String
  lstRecords.Clear
  strIonType_USE_THIS = cboIonType.List(cboIonType.ListIndex)
  ''''strSearchCriteria = _
      "select * from [DB_EqParams] " & _
      "where [strIonType]='" & strIonType_USE_THIS & "' " & _
      "order by [strResinName], [strIonType], [strIonName]"
  strSearchCriteria = _
      "select * from [DB_EqParams] " & _
      "where [strIonType]='" & strIonType_USE_THIS & "' " & _
      "order by [strIonName], [strIonType], [strResinName]"
  If (False = Database_TestForCriteria( _
    db1, _
    Rs1, _
    strSearchCriteria)) Then
    GoTo exit_err_ThisFunc
  End If
  Do While (Not Rs1.EOF)
    strResinName = Database_Get_String(Rs1, "strResinName")
    strIonType = Database_Get_String(Rs1, "strIonType")
    strIonName = Database_Get_String(Rs1, "strIonName")
    lngRecID = Database_Get_Long(Rs1, "lngRecID")
    ''''''''lstRecords.AddItem strResinName & " : " & strIonType & " : " & strIonName
    ''''lstRecords.AddItem strResinName & " : " & strIonName
    lstRecords.AddItem strIonName & " : " & strResinName
    lstRecords.ItemData(lstRecords.NewIndex) = lngRecID
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
Dim lngRecID As Long
'Dim sName_Duct As String
'Dim memNote_Duct As String
  If (lstRecords.ListCount <= 0) Or _
      (lstRecords.ListIndex < 0) Then
    GoTo exit_err_ThisFunc
  End If
  lngRecID = lstRecords.ItemData(lstRecords.ListIndex)
  If (False = Database_TestForCriteria( _
    db1, _
    Rs1, _
    "select * from [DB_EqParams] " & _
    "where [lngRecID]=" & Trim$(Str$(lngRecID)))) Then
    GoTo exit_err_ThisFunc
  End If
  With frmDB_EqParams_Record
    .strIonType = Database_Get_String(Rs1, "strIonType")
    .strIonName = Database_Get_String(Rs1, "strIonName")
    .dblEqValue = Database_Get_Double(Rs1, "dblEqValue")
    .strPresaturantIon = Database_Get_String(Rs1, "strPresaturantIon")
    .strEqValueSource = Database_Get_String(Rs1, "strEqValueSource")
    .strResinName = Database_Get_String(Rs1, "strResinName")
    .strResinType = Database_Get_String(Rs1, "strResinType")
    .strResinManufacturer = Database_Get_String(Rs1, "strResinManufacturer")
  End With
  Call frmDB_EqParams_Refresh
  Rs1.Close
exit_normally_ThisFunc:
  Populate_RecordData = True
  Exit Function
exit_err_ThisFunc:
  '
  ' DISPLAY BLANK DATA.
  '
  Call SetRecordDefaults(True)
  Call frmDB_EqParams_Refresh
  ''''txtDataStr(0).Text = ""
  ''''txtDataStr(1).Text = ""
  Populate_RecordData = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("Populate_RecordData")
  Resume exit_err_ThisFunc
End Function


Private Sub cboIonType_Click()
  '
  ' POPULATE RECORDS FROM DATABASE, AND INITIAL (BLANK) RECORD DATA.
  '
  Call Populate_lstRecords
  Call Populate_RecordData
  Call frmDB_EqParams_Refresh
End Sub


''''Private Sub cbo_iDuctType_Click()
''''Dim Ctl As Control
''''Set Ctl = cbo_iDuctType
''''  If (HALT_ALL_CONTROLS = True) Then Exit Sub
''''  If (Val(Ctl.Tag) = Ctl.ListIndex) Then Exit Sub
''''''''  With NowProj
''''''''    .Duct(.iDuct_Displayed).iDuctType = Ctl.ItemData(Ctl.ListIndex)
''''''''  End With
''''  With frmDB_EqParams_Record
''''    .iDuctType = Ctl.ItemData(Ctl.ListIndex)
''''  End With
''''  '
''''  ' RAISE DIRTY FLAG AND REFRESH WINDOW.
''''''''  Call Local_DirtyStatus_Set(Project_Is_Dirty, True)
''''  Call frmDB_EqParams_Refresh
''''End Sub
''''Private Sub cbo_iHotCold_Click()
''''Dim Ctl As Control
''''Set Ctl = cbo_iHotCold
''''  If (HALT_ALL_CONTROLS = True) Then Exit Sub
''''  If (Val(Ctl.Tag) = Ctl.ListIndex) Then Exit Sub
''''''''  With NowProj
''''''''    .Duct(.iDuct_Displayed).iHotCold = Ctl.ItemData(Ctl.ListIndex)
''''''''  End With
''''  With frmDB_EqParams_Record
''''    .iHotCold = Ctl.ItemData(Ctl.ListIndex)
''''  End With
''''  '
''''  ' RAISE DIRTY FLAG AND REFRESH WINDOW.
''''''''  Call Local_DirtyStatus_Set(Project_Is_Dirty, True)
''''  Call frmDB_EqParams_Refresh
''''End Sub
''''Private Sub cbo_sManometer_Click()
''''Dim Ctl As Control
''''Set Ctl = cbo_sManometer
''''  If (HALT_ALL_CONTROLS = True) Then Exit Sub
''''  If (Val(Ctl.Tag) = Ctl.ListIndex) Then Exit Sub
''''''''  With NowProj
''''''''    .Duct(.iDuct_Displayed).sManometer = Ctl.List(Ctl.ListIndex)
''''''''  End With
''''  With frmDB_EqParams_Record
''''    .sManometer = Ctl.List(Ctl.ListIndex)
''''  End With
''''  '
''''  ' RAISE DIRTY FLAG AND REFRESH WINDOW.
''''''''  Call Local_DirtyStatus_Set(Project_Is_Dirty, True)
''''  Call frmDB_EqParams_Refresh
''''End Sub


Private Sub cboUnits_Click(Index As Integer)
Dim Ctl As Control
Set Ctl = cboUnits(Index)
  Call unitsys_control_cbox_click(Ctl)
  Call frmDB_EqParams_Refresh
End Sub
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


Private Sub Form_Load()
  '
  ' ------- OPEN THE DATABASE. --------------------------------------
  '
  Call MFBMODEL_MDB_Open(db1)
  '
  ' POPULATE cboIonType.
  '
  Call Populate_cboIonType
  '
  ' MISC INITS.
  '
  Call CenterOnScreen(Me)
  Call frmDB_EqParams_Refresh_FirstFormLoad
  Call Global_GotFocus_ResetColors("ALL_WHITE")
  Select Case iFormMode
    Case iFormMode_IMPORT:
      cmdButton(0).Visible = True
      cmdButton(1).Caption = "&Cancel"
    Case iFormMode_EDIT:
      cmdButton(0).Visible = False
      cmdButton(1).Caption = "&Close"
  End Select
  frmDB_EqParams_Record.DB_Mode = DB_Mode_VIEW
  '
  ' POPULATE UNITS INTO SCROLLBOX CONTROLS.
  '
  Call Populate_frmDB_EqParams_Units
  '
  ' POPULATE RECORDS FROM DATABASE, AND INITIAL (BLANK) RECORD DATA.
  '
  Call Populate_lstRecords
  Call Populate_RecordData
  Call frmDB_EqParams_Refresh
End Sub
Private Sub Form_Unload(Cancel As Integer)
  '
  ' DEREGISTER UNIT CONTROLS.
  '
  Call unitsys_unregister_all_on_form(Me)
  '
  ' ------- CLOSE THE DATABASE. --------------------------------------
  '
  Call MFBMODEL_MDB_Close(db1)
  '
  ' MISC UNLOAD STUFF.
  '
  Call Global_GotFocus_ResetColors("NORMAL")
End Sub


Private Sub lstRecords_Click()
  Call Populate_RecordData
End Sub


Private Sub mnuRecordItem_Click(Index As Integer)
  On Error GoTo err_ThisFunc
Dim sThisName As String
Dim sMsg As String
Dim RetVal As Integer
Dim lngRecID As Long
Dim Rs1 As Recordset
  Select Case Index
    '
    '////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////
    Case 10:      'NEW.
      frmDB_EqParams_Record.DB_Mode = DB_Mode_ADDNEW
      '
      ' SET DEFAULTS.
      Call SetRecordDefaults(False)
      Call frmDB_EqParams_Refresh
    '
    '////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////
    Case 20:      'EDIT.
      If (IsRecordSelected() = False) Then
        GoTo exit_err_ThisFunc
      End If
      frmDB_EqParams_Record.DB_Mode = DB_Mode_EDIT
      Call frmDB_EqParams_Refresh
    '
    '////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////
    Case 30:      'DELETE.
      If (IsRecordSelected() = False) Then
        GoTo exit_err_ThisFunc
      End If
      sThisName = Trim$(lstRecords.List(lstRecords.ListIndex))
      sMsg = "Do you really want to delete record '" & _
          sThisName & "' from the database ?"
      RetVal = MsgBox(sMsg, vbCritical + vbYesNo, AppName_For_Display_Long)
      If (RetVal = vbNo) Then Exit Sub
      '
      ' DELETE THIS RECORD.
      '
      lngRecID = lstRecords.ItemData(lstRecords.ListIndex)
      If (False = Database_TestForCriteria( _
        db1, _
        Rs1, _
        "select * from [DB_EqParams] " & _
        "where [lngRecId]=" & Trim$(Str$(lngRecID)))) Then
        GoTo exit_err_ThisFunc
      End If
      Rs1.Delete
      Rs1.Close
      '
      ' REFRESH THE WINDOW.
      '
      Call SetRecordDefaults(True)
      Call Populate_lstRecords
      Call Populate_RecordData
      Call frmDB_EqParams_Refresh
    '
    '////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////
    Case 50:      'SAVE CHANGES.
      With frmDB_EqParams_Record
        If (.DB_Mode = DB_Mode_ADDNEW) Then
          ' START TO ADD NEW RECORD.
          Set Rs1 = db1.OpenRecordset("select * from [DB_EqParams]")
          Rs1.AddNew
        Else
          ' LOOK UP AND START EDITING EXISTING RECORD.
          lngRecID = lstRecords.ItemData(lstRecords.ListIndex)
          If (False = Database_TestForCriteria( _
            db1, _
            Rs1, _
            "select * from [DB_EqParams] " & _
            "where [lngRecId]=" & Trim$(Str$(lngRecID)))) Then
            GoTo exit_err_ThisFunc
          End If
          Rs1.Edit
        End If
        '
        ' STORE THE RECORD.
        '
        Rs1("strIonType") = .strIonType
        Rs1("strIonName") = .strIonName
        Rs1("dblEqValue") = .dblEqValue
        Rs1("strPresaturantIon") = .strPresaturantIon
        Rs1("strEqValueSource") = .strEqValueSource
        Rs1("strResinName") = .strResinName
        Rs1("strResinType") = .strResinType
        Rs1("strResinManufacturer") = .strResinManufacturer
        '
        ' COMPLETE THE STORAGE.
        '
        Rs1.Update
        Rs1.Close
        If (.DB_Mode = DB_Mode_ADDNEW) Then
          ' DO NOTHING.
        Else
          ' DO NOTHING.
        End If
      End With
      '
      ' REFRESH WINDOW.
      '
      frmDB_EqParams_Record.DB_Mode = DB_Mode_VIEW
      Call SetRecordDefaults(True)
      Call Populate_lstRecords
      Call Populate_RecordData
      Call frmDB_EqParams_Refresh
    '
    '////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////
    Case 60:      'CANCEL CHANGES.
      frmDB_EqParams_Record.DB_Mode = DB_Mode_VIEW
      Call Populate_RecordData
      Call frmDB_EqParams_Refresh
  End Select
exit_normally_ThisFunc:
  'mnuRecordItem_Click = True
  Exit Sub
exit_err_ThisFunc:
  'mnuRecordItem_Click = False
  Exit Sub
err_ThisFunc:
  Call Show_Trapped_Error("mnuRecordItem_Click")
  Resume exit_err_ThisFunc
End Sub


Private Sub txtData_GotFocus(Index As Integer)
Dim Ctl As Control
Set Ctl = txtData(Index)
Dim StatusMessagePanel As String
  If (Ctl.Locked = True) Then Exit Sub
  If (frmDB_EqParams_Record.DB_Mode <> DB_Mode_VIEW) Then
    Call unitsys_control_txtx_gotfocus(Ctl)
  End If
  Select Case Index
'    Case 0
'      StatusMessagePanel = "Type in the bed diameter"
'    Case 1
'      StatusMessagePanel = "Type in the bed length"
'    Case 2
'      StatusMessagePanel = "Type in the mass of adsorbent in the bed"
'    Case 3
'      StatusMessagePanel = "Type in the inlet flowrate"
  End Select
''''  Call Local_GenericStatus_Set(StatusMessagePanel)
End Sub
Private Sub txtData_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = Global_NumericKeyPress(KeyAscii)
End Sub
Private Sub txtData_LostFocus(Index As Integer)
Dim NewValue_Okay As Integer
Dim NewValue As Double
Dim Ctl As Control
Set Ctl = txtData(Index)
Dim Val_Low As Double
Dim Val_High As Double
Dim Raise_Dirty_Flag As Boolean
Dim Too_Small As Integer
Dim SetNoVal As Boolean
  If (Ctl.Locked = True) Then
    Call frmDB_EqParams_Refresh
    Exit Sub
  End If
  'NOTE: LOW AND HIGH VALUES IN BASE UNITS
  Select Case Index
    Case Else:
      Val_Low = -1E+20: Val_High = 1E+20
  End Select
  NewValue_Okay = False
  If (Trim$(Ctl.Text) = "") Then
    SetNoVal = True
    NewValue_Okay = True
    Raise_Dirty_Flag = True
  Else
    SetNoVal = False
    If (unitsys_control_txtx_lostfocus_validate( _
        Ctl, _
        Val_Low, _
        Val_High, _
        NewValue, _
        Raise_Dirty_Flag)) Then
      NewValue_Okay = True
    End If
  End If
  Call unitsys_control_txtx_lostfocus(Ctl, NewValue)
''''  Call Local_GenericStatus_Set("")
  If (NewValue_Okay) Then
    If (Raise_Dirty_Flag) Then
      '
      ' STORE TO MEMORY.
      '
      ''''With NowProj
      With frmDB_EqParams_Record
        Select Case Index
          '------------------------------------------------------------------------------------------------------------------------
          '
          ' PART ONE.
          '
''''          Case 0: .Duct(.iDuct_Displayed).dblDiameter = IIf(SetNoVal, NoValue_dbl, NewValue)
''''          Case 1: .Duct(.iDuct_Displayed).dblDepth = IIf(SetNoVal, NoValue_dbl, NewValue)
''''          Case 2: .Duct(.iDuct_Displayed).iNumPortsForDuct = IIf(SetNoVal, NoValue_i, CInt(NewValue))
''''          Case 3: '''' OUTPUT ONLY.
          Case 0: .dblEqValue = NewValue
        End Select
      End With
''''      If (Raise_Dirty_Flag) Then
''''        'THROW DIRTY FLAG.
''''        Call Local_DirtyStatus_Set(Project_Is_Dirty, True)
''''      End If
    End If
  End If
  'REFRESH WINDOW.
  Call frmDB_EqParams_Refresh
End Sub


Private Sub txtDataStr_GotFocus(Index As Integer)
Dim Ctl As Control
Set Ctl = txtDataStr(Index)
Dim StatusMessagePanel As String
  If (Ctl.Locked = True) Then Exit Sub
  Ctl.Text = Ctl.Tag
  If (frmDB_EqParams_Record.DB_Mode <> DB_Mode_VIEW) Then
    Call Global_GotFocus(Ctl)
  End If
  Select Case Index
    Case 4:
      StatusMessagePanel = "Type in the duct name"
  End Select
''''  Call Local_GenericStatus_Set(StatusMessagePanel)
End Sub
Private Sub txtDataStr_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = Global_TextKeyPress(KeyAscii)
End Sub
Private Sub txtDataStr_LostFocus(Index As Integer)
Dim Ctl As Control
Set Ctl = txtDataStr(Index)
Dim SetNoVal As Boolean
Dim Raise_Dirty_Flag As Boolean
Dim NewStr As String
  If (Ctl.Locked = True) Then
    Call frmDB_EqParams_Refresh
    Exit Sub
  End If
  SetNoVal = False
  Raise_Dirty_Flag = False
  If (Trim$(Ctl.Text) = "") Then
    SetNoVal = True
    'NewValue_Okay = True
    If (Trim$(Ctl.Tag) <> "") Then
      Raise_Dirty_Flag = True
    End If
  Else
    If (Trim$(Ctl.Tag) <> Trim$(Ctl.Text)) Then
      Raise_Dirty_Flag = True
    End If
  End If
  Call Global_LostFocus(Ctl)
''''  Call Local_GenericStatus_Set("")
  If (Raise_Dirty_Flag = True) Then
    NewStr = Ctl.Text
    '
    ' STORE TO MEMORY.
    ''''With NowProj
    With frmDB_EqParams_Record
      Select Case Index
        '------------------------------------------------------------------------------------------------------------------------
        '
        ' PART ONE.
        '
        Case 0: .strIonName = NewStr
        Case 1: .strPresaturantIon = NewStr
        Case 2: .strEqValueSource = NewStr
        Case 3: .strResinName = NewStr
        Case 4: .strResinType = NewStr
        Case 5: .strResinManufacturer = NewStr
      End Select
    End With
''''    '
''''    ' THROW DIRTY FLAG.
''''    Call Local_DirtyStatus_Set(Project_Is_Dirty, True)
  End If
  '
  ' REFRESH WINDOW.
  Call frmDB_EqParams_Refresh
End Sub




