VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmchembrowse 
   Caption         =   "Chemical Selection"
   ClientHeight    =   5550
   ClientLeft      =   1110
   ClientTop       =   1905
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5550
   ScaleWidth      =   7065
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   0
      Left            =   4860
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   270
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   14
      Top             =   5085
      Width           =   1815
   End
   Begin VB.CommandButton cmdaccept 
      Caption         =   "&Accept"
      Height          =   375
      Left            =   1440
      TabIndex        =   11
      Top             =   5085
      Width           =   1815
   End
   Begin VB.Frame frsort 
      Caption         =   "Sort by"
      Height          =   735
      Index           =   4
      Left            =   240
      TabIndex        =   13
      Top             =   4200
      Width           =   6495
      Begin VB.OptionButton optdesc 
         Caption         =   "descending"
         Height          =   375
         Left            =   3960
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optasc 
         Caption         =   "ascending"
         Height          =   375
         Left            =   2640
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.ComboBox cbosort 
         Height          =   315
         Left            =   240
         TabIndex        =   7
         Text            =   "Combo3"
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdoksort 
         Caption         =   "&Ok"
         Height          =   375
         Left            =   5640
         TabIndex        =   10
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame frfilter 
      Caption         =   "Filter"
      Height          =   735
      Index           =   3
      Left            =   240
      TabIndex        =   12
      Top             =   3360
      Width           =   6495
      Begin VB.ComboBox cbofiltercat 
         Height          =   315
         Left            =   2400
         TabIndex        =   5
         Text            =   "Combo2"
         Top             =   240
         Width           =   3015
      End
      Begin VB.ComboBox cbofilterfield 
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdokfilter 
         Caption         =   "&Ok"
         Height          =   375
         Left            =   5640
         TabIndex        =   6
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame frfind 
      Caption         =   "Find"
      Height          =   855
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   2400
      Width           =   6495
      Begin VB.CommandButton cmdokfind 
         Caption         =   "&Ok"
         Height          =   375
         Left            =   5640
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
      Begin VB.ComboBox cbofind 
         Height          =   315
         Left            =   240
         TabIndex        =   0
         Text            =   "Combo1"
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox tbxfind 
         Height          =   375
         Left            =   2400
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   360
         Width           =   3015
      End
   End
   Begin MSDBGrid.DBGrid grdchemlist 
      Bindings        =   "Frmchbr.frx":0000
      Height          =   1995
      Index           =   0
      Left            =   315
      OleObjectBlob   =   "Frmchbr.frx":0013
      TabIndex        =   15
      Top             =   225
      Width           =   6405
   End
End
Attribute VB_Name = "frmchembrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cbofiltercat_Click()

     cmdokfind.Default = False
    cmdoksort.Default = False
    cmdaccept.Default = False
    cmdokfilter.Default = True
    
End Sub


Private Sub cbofiltercat_GotFocus()

    cmdokfilter.Default = True
End Sub


Private Sub cbofiltercat_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyLeft Then
        cbofilterfield.SetFocus
    ElseIf KeyCode = vbKeyRight Then
        cmdokfilter.SetFocus
    ElseIf KeyCode = vbKeyDown Then
        cbosort.SetFocus
        cmdoksort.Default = True
    ElseIf KeyCode = vbKeyUp Then
        cbofind.SetFocus
        cmdokfind.Default = True
    End If
    
End Sub


Private Sub cbofilterfield_Click()

    cmdokfind.Default = False
    cmdoksort.Default = False
    cmdaccept.Default = False
    cmdokfilter.Default = True
    Select Case Trim(cbofilterfield.Text)
        Case "Source"
            cbofiltercat.Clear
            cbofiltercat.AddItem "CAA"
            cbofiltercat.AddItem "General"
            cbofiltercat.AddItem "EPA-RMP"
            cbofiltercat.AddItem "OSHA"
            cbofiltercat.Text = cbofiltercat.List(0)
        Case "Chemical Family"
            cbofiltercat.Clear
            cbofiltercat.AddItem "AA - N-ALKANES"
            cbofiltercat.AddItem "MA - METHYLALKANES"
            cbofiltercat.AddItem "DA - DIMETHYLALKANES"
            cbofiltercat.AddItem "OA - OTHER ALKANES"
            cbofiltercat.AddItem "CA - CYCLOALKANES"
            cbofiltercat.AddItem "GA - ALKYLCYCLOPENTANES"
            cbofiltercat.AddItem "HA - ALKYLCYCLOHEXANES"
            cbofiltercat.AddItem "RA - MULTIRING CYCLOALKANES"
            cbofiltercat.AddItem "AE - 1-ALKENES"
            cbofiltercat.AddItem "BE - 2,3,4-ALKENES"
            cbofiltercat.AddItem "ME - METHYLALKENES"
            cbofiltercat.AddItem "EE - ETHYL & HIGHER ALKENES"
            cbofiltercat.AddItem "CE - CYCLOALKENES"
            cbofiltercat.AddItem "DE - DIALKENES"
            cbofiltercat.AddItem "YY - ALKYNES"
            cbofiltercat.AddItem "AR - N-ALKYLBENZENES"
            cbofiltercat.AddItem "BR - OTHER ALKYLBENZENES"
            cbofiltercat.AddItem "MR - OTHER MONOAROMATICS"
            cbofiltercat.AddItem "DR - NAPTHALENES"
            cbofiltercat.AddItem "OR - OTHER CONDENSED RINGS"
            cbofiltercat.AddItem "PR - DIPHENYL/POLYAROMATICS"
            cbofiltercat.AddItem "TR - TERPENES"
            cbofiltercat.AddItem "HR - OTHER HYDROCARBON RINGS"
            cbofiltercat.AddItem "IG - INORGANIC GASES"
            cbofiltercat.AddItem "LD - ALDEHYDES"
            cbofiltercat.AddItem "KK - KETONES"
            cbofiltercat.AddItem "AL - N-ALCOHOLS"
            cbofiltercat.AddItem "OL - OTHER ALIPHATIC ALCOHOLS"
            cbofiltercat.AddItem "CL - CYCLOALIPHATIC ALCOHOLS"
            cbofiltercat.AddItem "RL - AROMATIC ALCOHOLS"
            cbofiltercat.AddItem "PL - POLYOLS"
            cbofiltercat.AddItem "AC - N-ALIPHATIC ACIDS"
            cbofiltercat.AddItem "OC - OTHER ALIPHATIC ACIDS"
            cbofiltercat.AddItem "DC - DICARBOXYLIC ACIDS"
            cbofiltercat.AddItem "RC - AROMATIC CARBOXYLIC ACIDS"
            cbofiltercat.AddItem "HD - ANHYDRIDES"
            cbofiltercat.AddItem "FS - FORMATES"
            cbofiltercat.AddItem "ES - ACETATES"
            cbofiltercat.AddItem "BS - PROPIONATES AND BUTYRATES"
            cbofiltercat.AddItem "SS - OTHER SATURATED ALIPHATIC ESTERS"
            cbofiltercat.AddItem "US - UNSATURATED ALIPHATIC ESTERS"
            cbofiltercat.AddItem "RS - AROMATIC ESTERS"
            cbofiltercat.AddItem "AT - ALIPHATIC ETHERS"
            cbofiltercat.AddItem "OT - OTHER ETHERS/DIETHERS"
            cbofiltercat.AddItem "CT - EPOXIDES"
            cbofiltercat.AddItem "TT - PEROXIDES"
            cbofiltercat.AddItem "AH - C1/C2 ALIPHATIC CHLORIDES"
            cbofiltercat.AddItem "HH - C3 & HIGHER ALIPHATIC CHLORIDES"
            cbofiltercat.AddItem "RH - AROMATIC CHLORIDES"
            cbofiltercat.AddItem "VH - C,H,BR COMPOUNDS"
            cbofiltercat.AddItem "WH - C,H,I COMPOUNDS"
            cbofiltercat.AddItem "FH - C,H,F COMPOUNDS"
            cbofiltercat.AddItem "PH - C,H, MULTIHALOGEN COMPOUNDS"
            cbofiltercat.AddItem "AM - N-ALIPHATIC PRIMARY AMINES"
            cbofiltercat.AddItem "BM - OTHER ALIPHATIC AMINES"
            cbofiltercat.AddItem "RM - AROMATIC AMINES"
            cbofiltercat.AddItem "OM - OTHER AMINES, IMINES"
            cbofiltercat.AddItem "NX - NITRILES"
            cbofiltercat.AddItem "TN - C,H,NO2 COMPOUNDS"
            cbofiltercat.AddItem "YN - ISOCYANATES/DIISOCYANATES"
            cbofiltercat.AddItem "SD - MERCAPTANS"
            cbofiltercat.AddItem "SF - SULFIDES/THIOPHENES"
            cbofiltercat.AddItem "PC - POLYFUNCTIONAL ACIDS"
            cbofiltercat.AddItem "PS - POLYFUNCTIONAL ESTERS"
            cbofiltercat.AddItem "PO - OTHER POLYFUNCTIONAL C,H,O"
            cbofiltercat.AddItem "NP - POLYFUNCTIONAL NITRILES"
            cbofiltercat.AddItem "TM - NITROAMINES"
            cbofiltercat.AddItem "PM - POLYFUNCTIONAL AMIDES/AMINES"
            cbofiltercat.AddItem "PN - POLYFUNCTIONAL C,H,O,N"
            cbofiltercat.AddItem "SP - POLYFUNCTIONAL C,H,O,S"
            cbofiltercat.AddItem "HP - POLYFUNCTIONAL C,H,O,HALIDE"
            cbofiltercat.AddItem "BP - POLYFUNCTIONAL C,H,N,HALIDE,(O)"
            cbofiltercat.AddItem "OP - OTHER POLYFUNCTIONAL ORGANICS"
            cbofiltercat.AddItem "LX - ELEMENTS"
            cbofiltercat.AddItem "SX - SILANES/SILOXANES"
            cbofiltercat.AddItem "GI - ORGANIC/INORGANIC COMPOUNDS"
            cbofiltercat.AddItem "IC - INORGANIC ACIDS"
            cbofiltercat.AddItem "IB - INORGANIC BASES"
            cbofiltercat.AddItem "GS - ORGANIC SALTS"
            cbofiltercat.AddItem "XS - SODIUM SALTS"
            cbofiltercat.AddItem "OS - OTHER INORGANIC SALTS"
            cbofiltercat.AddItem "IH - INORGANIC HALIDES"
            cbofiltercat.AddItem "OI - OTHER INORGANICS"
            cbofiltercat.Text = cbofiltercat.List(0)
        Case "user input (Name)"
            cbofiltercat.Clear
           
    End Select
End Sub


Private Sub cbofilterfield_GotFocus()

    cmdokfilter.Default = True
End Sub


Private Sub cbofilterfield_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyRight Then
        cbofiltercat.SetFocus
    ElseIf KeyCode = vbKeyDown Then
        cbosort.SetFocus
    ElseIf KeyCode = vbKeyUp Then
        cbofind.SetFocus
    End If
End Sub


Private Sub cbofind_Click()

    cmdokfilter.Default = False
    cmdoksort.Default = False
    cmdaccept.Default = False
    cmdokfind.Default = True
    tbxfind.Text = ""
End Sub


Private Sub cbofind_GotFocus()

    cmdokfind.Default = True
End Sub


Private Sub cbofind_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyRight Then
        tbxfind.SetFocus
    ElseIf KeyCode = vbKeyUp Then
            grdchemlist(0).SetFocus
    ElseIf KeyCode = vbKeyDown Then
        cbofilterfield.SetFocus
    End If
End Sub


Private Sub cbosort_Click()

    cmdoksort.Default = True
End Sub

Private Sub cbosort_GotFocus()

    cmdoksort.Default = True
End Sub

Private Sub cbosort_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyRight Then
        optasc.SetFocus
    ElseIf KeyCode = vbKeyDown Then
        cmdaccept.SetFocus
    ElseIf KeyCode = vbKeyUp Then
        cbofilterfield.SetFocus
    End If
End Sub


Private Sub cmdaccept_Click()

    ' for now we'll just show the selected chem info and exit
    'On Error GoTo none_selected
    On Error Resume Next
    selected_cas = CLng(Data1(MASTER).Recordset("CAS"))
    selected_name = Data1(MASTER).Recordset("Name")
    selected_smiles = Trim(Data1(MASTER).Recordset("Smiles"))
    selected_structure = Data1(MASTER).Recordset("Formula")
    selected_family = Trim(Data1(MASTER).Recordset("Chemical Family"))
    selected_temperature = 25
    selected_temp_units = "C"
    selected_rings = -1 ' still don't have this
    Data1(MASTER).Recordset.Close
    On Error GoTo none_selected
    If Len(CStr(selected_cas)) > 0 Then
        frmchembrowse.Hide
    Else
        GoTo none_selected
    End If
    
    
    Exit Sub
none_selected:
    
    
    'MsgBox ("no chemical selected")
    
End Sub

Private Sub cmdaccept_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyUp Then
        cbosort.SetFocus
    ElseIf KeyCode = vbKeyDown Then
        grdchemlist(0).SetFocus
    End If
End Sub


Private Sub cmdexit_Click()

    On Error Resume Next
    'chembrowsedb.Close
    frmchembrowse.Hide
    Unload Me
End Sub

Private Sub cmdokfilter_Click()
    Dim Criteria As String
    Dim Field As String
    Dim RChar As Integer
    Dim LChar As Integer
    Criteria = Trim(cbofiltercat.Text)
    Field = Trim(cbofilterfield.Text)
    If Field = "Chemical Family" Then
        Criteria = Mid(Criteria, 1, 2)
    ElseIf Field = "user input (Name)" Then
            Field = "Name"
    End If
    
    ' to restore to full list
    If Trim(cbofiltercat.Text) = "" Then
        Data1(MASTER).RecordSource = "PEARLS List"
        Data1(MASTER).Refresh
        cmdaccept.Default = True
        Exit Sub
    End If
    If Field = "CAS" Then
        Data1(MASTER).RecordSource = "SELECT * FROM [PEARLS List] WHERE [" & Field & "] = " & Val(Criteria)
    Else
        Data1(MASTER).RecordSource = "SELECT * FROM [PEARLS List] WHERE [" & Field & "] Like " & "'*" & Criteria & "*'"
    
    End If
    'FRMMain!TABViewProp.CurrTab = 6
    Data1(MASTER).Refresh
    cmdaccept.Default = True
End Sub

Private Sub cmdokfilter_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyLeft Then
        cbofiltercat.SetFocus
    ElseIf KeyCode = vbKeyUp Then
        cbofind.SetFocus
    ElseIf KeyCode = vbKeyDown Then
        cbosort.SetFocus
    End If
        
End Sub


Private Sub cmdokfind_Click()

    Dim searchtable As Recordset
    Dim Criteria As String
    Dim Field As String
    Dim RChar As Integer
    Dim LChar As Integer
    Dim DBTbl As Recordset
    
    Screen.MousePointer = 11
    ' code that will indicate which grid and data object we're using
    ' the criteria from the browser form
    Criteria = tbxfind.Text
    Field = cbofind.Text
    
    If Field = "Chemical Family" Then
        Criteria = Mid(Criteria, 1, 2)
    End If
    
    If Trim(tbxfind.Text) = "" Then
        MsgBox "No criteria specified", 48, "No Criteria"
        Screen.MousePointer = 1
        Exit Sub
    End If
    
    RChar = Asc(Right(Criteria, 1))
    LChar = Asc(Left(Criteria, 1))
    
    'Search only CAS if a number was entered
    If LChar > 47 And LChar < 59 And RChar > 47 And RChar < 59 Then
        Data1(MASTER).Recordset.FindFirst "CAS = " & Val(Criteria)
        If Not Data1(MASTER).Recordset.NoMatch Then
            Screen.MousePointer = 1
            cmdaccept.Default = True
            Exit Sub
        Else
            Data1(MASTER).Refresh
            frmchembrowse.Refresh
            MsgBox (Criteria & " not found, doing synonym search...")
        End If
        
    Else
        ' do a search based on the criteria the user had entered
        Data1(MASTER).Recordset.FindNext "[" & Field & "] LIKE '*" & Criteria & "*'"
        If Data1(MASTER).Recordset.NoMatch Then
            If Field = "Name" Then
                    Set searchtable = chembrowsedb.OpenRecordset("Synonym List", dbOpenSnapshot)
                    searchtable.FindNext "[Name] LIKE '*" & Criteria & "*'"
                    If Not searchtable.NoMatch Then
                        MsgBox "Synonym match found", 48, "Synonym Match"
                        Data1(MASTER).Recordset.FindNext "[CAS] = '*" & searchtable("CAS") & "*'"
                        searchtable.Close
                        
                        Screen.MousePointer = 1
                        If Not Data1(MASTER).Recordset.NoMatch Then
                            cmdaccept.Default = True
                            Exit Sub
                        Else
                            Data1(MASTER).Refresh
                            frmchembrowse.Refresh
                            MsgBox (Criteria & " not found, doing synonym search...")
                        End If
                    End If
                    Set searchtable = chembrowsedb.OpenRecordset("DIPPR801", dbOpenSnapshot)
                    searchtable.FindNext "[INAM] LIKE '*" & Criteria & "*'"
                    If Not searchtable.NoMatch Then
                        MsgBox "Synonym match found", 48, "Synonym Match"
                        Data1(MASTER).Recordset.FindNext "[CAS] = '*" & searchtable("CASN") & "*'"
                        
                        searchtable.Close
                        Screen.MousePointer = 1
                        cmdaccept.Default = True
                        Exit Sub
                    End If
                    
                    searchtable.FindNext "[CNAM] LIKE '*" & Criteria & "*'"
                    If Not searchtable.NoMatch Then
                        MsgBox "Synonym match found", 48, "Synonym Match"
                        Data1(MASTER).Recordset.FindNext "[CAS] = '*" & searchtable("CASN") & "*'"
                        
                        searchtable.Close
                        Screen.MousePointer = 1
                        cmdaccept.Default = True
                        Exit Sub
                    End If
                    searchtable.Close
            Else
                'Data1(whichgrid).Refresh
                'grdchemlist(whichgrid).Refresh
                'frmchembrowse.Refresh
                Screen.MousePointer = 1
                cmdaccept.Default = True
                Exit Sub
            End If
        End If
    End If
    
    
    Screen.MousePointer = 1
End Sub

Private Sub cmdokfind_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyLeft Then
        tbxfind.SetFocus
    ElseIf KeyCode = vbKeyUp Then
        grdchemlist(0).SetFocus
    ElseIf KeyCode = vbKeyDown Then
        cbofilterfield.SetFocus
    End If
End Sub


Private Sub cmdoksort_Click()

    Dim TempCAS As Long
    Dim Field As String
    Dim was_selected As Boolean
    was_selected = True
    Field = Trim(cbosort.Text)
    On Error GoTo set_error_flag
    TempCAS = Data1(MASTER).Recordset("CAS")
    
    Screen.MousePointer = 11
    If optasc.value = True Then
        Data1(MASTER).RecordSource = "SELECT * FROM [PEARLS List] ORDER BY [" & Field & "] ASC"
        If was_selected = True Then
            Data1(MASTER).Recordset.FindFirst "CAS = " & TempCAS
        End If
        Data1(MASTER).Refresh
    ElseIf optdesc.value = True Then
        Data1(MASTER).RecordSource = "SELECT * FROM [PEARLS List] ORDER BY [" & Field & "] DESC"
        If was_selected = True Then
            Data1(MASTER).Recordset.FindFirst "CAS = " & TempCAS
        End If
        Data1(MASTER).Refresh
        'frmchembrowse.Refresh
    End If
    Screen.MousePointer = 1
    cmdaccept.Default = True
    Exit Sub
set_error_flag:
    was_selected = False
    Resume Next
End Sub

Private Sub cmdoksort_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyDown Then
        cmdaccept.SetFocus
    ElseIf KeyCode = vbKeyUp Then
        cbofilterfield.SetFocus
    ElseIf KeyCode = vbKeyLeft Then
        optdesc.SetFocus
    End If
    
End Sub


Private Sub Form_Load()

    'Call load_chem_browse_info
    optasc.value = True
    tbxfind.Text = ""
    Call cbofilterfield_Click
    cmdaccept.Default = True
    'Data1.Refresh
End Sub


Private Sub frfilter_Click(Index As Integer)

     cmdokfind.Default = False
    cmdoksort.Default = False
    cmdaccept.Default = False
    cmdokfilter.Default = True
End Sub

Private Sub frfind_Click(Index As Integer)

    cmdokfilter.Default = False
    cmdoksort.Default = False
    cmdaccept.Default = False
    cmdokfind.Default = True
End Sub

Private Sub grdchemlist_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyDown Then
        cbofind.SetFocus
    End If
    Data1(Index).Refresh
End Sub

Private Sub optasc_GotFocus()

    cmdoksort.Default = True
End Sub

Private Sub optasc_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyRight Then
        cmdoksort.SetFocus
    ElseIf KeyCode = vbKeyLeft Then
        cbosort.SetFocus
    ElseIf KeyCode = vbKeyDown Then
        cmdaccept.SetFocus
    ElseIf KeyCode = vbKeyUp Then
        cbofilterfield.SetFocus
    End If
End Sub


Private Sub optdesc_GotFocus()

    cmdoksort.Default = True
End Sub

Private Sub optdesc_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyRight Then
        cmdoksort.SetFocus
    ElseIf KeyCode = vbKeyLeft Then
        optasc.SetFocus
    ElseIf KeyCode = vbKeyDown Then
        cmdaccept.SetFocus
    ElseIf KeyCode = vbKeyUp Then
        cbofilterfield.SetFocus
    End If
End Sub

Private Sub tbxfind_Click()

    cmdokfilter.Default = False
    cmdoksort.Default = False
    cmdaccept.Default = False
    cmdokfind.Default = True
End Sub


Private Sub tbxfind_GotFocus()

    cmdokfind.Default = True
End Sub


Private Sub tbxfind_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyLeft Then
        cbofind.SetFocus
    ElseIf KeyCode = vbKeyRight Then
        cmdokfind.SetFocus
    ElseIf KeyCode = vbKeyUp Then
        grdchemlist(0).SetFocus
    ElseIf KeyCode = vbKeyDown Then
        cbofilterfield.SetFocus
    End If
End Sub


