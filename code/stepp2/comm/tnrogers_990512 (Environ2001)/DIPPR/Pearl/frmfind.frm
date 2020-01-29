VERSION 5.00
Begin VB.Form frmfind 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find"
   ClientHeight    =   1590
   ClientLeft      =   1575
   ClientTop       =   1785
   ClientWidth     =   5910
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   Icon            =   "frmfind.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1590
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CMDRestore_List 
      Caption         =   "&Restore List"
      Height          =   375
      Left            =   1560
      TabIndex        =   8
      Top             =   1080
      Width           =   1275
   End
   Begin VB.ComboBox CMBFindStr 
      Height          =   315
      ItemData        =   "frmfind.frx":030A
      Left            =   1800
      List            =   "frmfind.frx":030C
      TabIndex        =   7
      Text            =   "CMBFindStr"
      Top             =   120
      Width           =   3975
   End
   Begin VB.CommandButton CMDFilter 
      Caption         =   "&Filter"
      Height          =   375
      Left            =   2940
      TabIndex        =   6
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton CMDFindPrevious 
      Caption         =   "Find &Previous"
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   1080
      Width           =   1335
   End
   Begin VB.ComboBox CMBFind 
      Height          =   315
      Left            =   1800
      TabIndex        =   3
      Text            =   "CMBFind"
      Top             =   600
      Width           =   2535
   End
   Begin VB.CommandButton CMDFindNext 
      Caption         =   "Find &Next"
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   540
      Width           =   1335
   End
   Begin VB.CommandButton CMDClose 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label LBLSearchCol 
      Caption         =   "&Search/Apply To:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   645
      Width           =   1575
   End
   Begin VB.Label LBLFind 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Find/Filter &Criteria:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   1695
   End
End
Attribute VB_Name = "frmfind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMBFind_Click()

    Select Case CMBFind.Text
        Case "Source"
            CMBFindStr.AddItem "CAA"
            CMBFindStr.AddItem "General"
            CMBFindStr.AddItem "EPA-RMP"
            CMBFindStr.AddItem "OSHA"
        Case "Chemical Family"
            CMBFindStr.AddItem "AA - N-ALKANES"
            CMBFindStr.AddItem "MA - METHYLALKANES"
            CMBFindStr.AddItem "DA - DIMETHYLALKANES"
            CMBFindStr.AddItem "OA - OTHER ALKANES"
            CMBFindStr.AddItem "CA - CYCLOALKANES"
            CMBFindStr.AddItem "GA - ALKYLCYCLOPENTANES"
            CMBFindStr.AddItem "HA - ALKYLCYCLOHEXANES"
            CMBFindStr.AddItem "RA - MULTIRING CYCLOALKANES"
            CMBFindStr.AddItem "AE - 1-ALKENES"
            CMBFindStr.AddItem "BE - 2,3,4-ALKENES"
            CMBFindStr.AddItem "ME - METHYLALKENES"
            CMBFindStr.AddItem "EE - ETHYL & HIGHER ALKENES"
            CMBFindStr.AddItem "CE - CYCLOALKENES"
            CMBFindStr.AddItem "DE - DIALKENES"
            CMBFindStr.AddItem "YY - ALKYNES"
            CMBFindStr.AddItem "AR - N-ALKYLBENZENES"
            CMBFindStr.AddItem "BR - OTHER ALKYLBENZENES"
            CMBFindStr.AddItem "MR - OTHER MONOAROMATICS"
            CMBFindStr.AddItem "DR - NAPTHALENES"
            CMBFindStr.AddItem "OR - OTHER CONDENSED RINGS"
            CMBFindStr.AddItem "PR - DIPHENYL/POLYAROMATICS"
            CMBFindStr.AddItem "TR - TERPENES"
            CMBFindStr.AddItem "HR - OTHER HYDROCARBON RINGS"
            CMBFindStr.AddItem "IG - INORGANIC GASES"
            CMBFindStr.AddItem "LD - ALDEHYDES"
            CMBFindStr.AddItem "KK - KETONES"
            CMBFindStr.AddItem "AL - N-ALCOHOLS"
            CMBFindStr.AddItem "OL - OTHER ALIPHATIC ALCOHOLS"
            CMBFindStr.AddItem "CL - CYCLOALIPHATIC ALCOHOLS"
            CMBFindStr.AddItem "RL - AROMATIC ALCOHOLS"
            CMBFindStr.AddItem "PL - POLYOLS"
            CMBFindStr.AddItem "AC - N-ALIPHATIC ACIDS"
            CMBFindStr.AddItem "OC - OTHER ALIPHATIC ACIDS"
            CMBFindStr.AddItem "DC - DICARBOXYLIC ACIDS"
            CMBFindStr.AddItem "RC - AROMATIC CARBOXYLIC ACIDS"
            CMBFindStr.AddItem "HD - ANHYDRIDES"
            CMBFindStr.AddItem "FS - FORMATES"
            CMBFindStr.AddItem "ES - ACETATES"
            CMBFindStr.AddItem "BS - PROPIONATES AND BUTYRATES"
            CMBFindStr.AddItem "SS - OTHER SATURATED ALIPHATIC ESTERS"
            CMBFindStr.AddItem "US - UNSATURATED ALIPHATIC ESTERS"
            CMBFindStr.AddItem "RS - AROMATIC ESTERS"
            CMBFindStr.AddItem "AT - ALIPHATIC ETHERS"
            CMBFindStr.AddItem "OT - OTHER ETHERS/DIETHERS"
            CMBFindStr.AddItem "CT - EPOXIDES"
            CMBFindStr.AddItem "TT - PEROXIDES"
            CMBFindStr.AddItem "AH - C1/C2 ALIPHATIC CHLORIDES"
            CMBFindStr.AddItem "HH - C3 & HIGHER ALIPHATIC CHLORIDES"
            CMBFindStr.AddItem "RH - AROMATIC CHLORIDES"
            CMBFindStr.AddItem "VH - C,H,BR COMPOUNDS"
            CMBFindStr.AddItem "WH - C,H,I COMPOUNDS"
            CMBFindStr.AddItem "FH - C,H,F COMPOUNDS"
            CMBFindStr.AddItem "PH - C,H, MULTIHALOGEN COMPOUNDS"
            CMBFindStr.AddItem "AM - N-ALIPHATIC PRIMARY AMINES"
            CMBFindStr.AddItem "BM - OTHER ALIPHATIC AMINES"
            CMBFindStr.AddItem "RM - AROMATIC AMINES"
            CMBFindStr.AddItem "OM - OTHER AMINES, IMINES"
            CMBFindStr.AddItem "NX - NITRILES"
            CMBFindStr.AddItem "TN - C,H,NO2 COMPOUNDS"
            CMBFindStr.AddItem "YN - ISOCYANATES/DIISOCYANATES"
            CMBFindStr.AddItem "SD - MERCAPTANS"
            CMBFindStr.AddItem "SF - SULFIDES/THIOPHENES"
            CMBFindStr.AddItem "PC - POLYFUNCTIONAL ACIDS"
            CMBFindStr.AddItem "PS - POLYFUNCTIONAL ESTERS"
            CMBFindStr.AddItem "PO - OTHER POLYFUNCTIONAL C,H,O"
            CMBFindStr.AddItem "NP - POLYFUNCTIONAL NITRILES"
            CMBFindStr.AddItem "TM - NITROAMINES"
            CMBFindStr.AddItem "PM - POLYFUNCTIONAL AMIDES/AMINES"
            CMBFindStr.AddItem "PN - POLYFUNCTIONAL C,H,O,N"
            CMBFindStr.AddItem "SP - POLYFUNCTIONAL C,H,O,S"
            CMBFindStr.AddItem "HP - POLYFUNCTIONAL C,H,O,HALIDE"
            CMBFindStr.AddItem "BP - POLYFUNCTIONAL C,H,N,HALIDE,(O)"
            CMBFindStr.AddItem "OP - OTHER POLYFUNCTIONAL ORGANICS"
            CMBFindStr.AddItem "LX - ELEMENTS"
            CMBFindStr.AddItem "SX - SILANES/SILOXANES"
            CMBFindStr.AddItem "GI - ORGANIC/INORGANIC COMPOUNDS"
            CMBFindStr.AddItem "IC - INORGANIC ACIDS"
            CMBFindStr.AddItem "IB - INORGANIC BASES"
            CMBFindStr.AddItem "GS - ORGANIC SALTS"
            CMBFindStr.AddItem "XS - SODIUM SALTS"
            CMBFindStr.AddItem "OS - OTHER INORGANIC SALTS"
            CMBFindStr.AddItem "IH - INORGANIC HALIDES"
            CMBFindStr.AddItem "OI - OTHER INORGANICS"
    End Select
    
End Sub


Private Sub CMDClose_Click()
    
    Unload FRMFind
    
End Sub

Private Sub CMDFilter_Click()
    
    Dim Criteria As String
    Dim Field As String
    Dim RChar As Integer
    Dim LChar As Integer
    
    Criteria = CMBFindStr.Text
    Field = CMBFind.Text
    
    If Field = "Chemical Family" Then
        Criteria = Mid(Criteria, 1, 2)
    End If
               
    If CMBFindStr.Text = "" Then
        FRMMain!Data1.RecordSource = "PEARLS List"
        FRMMain!Data1.Refresh
        Exit Sub
    End If
               
    If Field = "CAS" Then
        FRMMain!Data1.RecordSource = "SELECT * FROM [PEARLS List] WHERE [" & Field & "] = " & Val(Criteria)
    Else
        FRMMain!Data1.RecordSource = "SELECT * FROM [PEARLS List] WHERE [" & Field & "] ='" & Criteria & "'"
    End If
    FRMMain!TABViewProp.CurrTab = 6
    FRMMain!Data1.Refresh
   
End Sub

Private Sub CMDFindNext_Click()
    
    ' this doesn't seem to be closing the tables DENISE check
    Dim Criteria As String
    Dim Field As String
    Dim RChar As Integer
    Dim LChar As Integer
    Dim DBTbl As Recordset
    
    Screen.MousePointer = 11
    
    Criteria = CMBFindStr.Text
    Field = CMBFind.Text
    
    If Field = "Chemical Family" Then
        Criteria = Mid(Criteria, 1, 2)
    End If
    
    If CMBFindStr.Text = "" Then
        MsgBox "No criteria specified", 48, "No Criteria"
        Screen.MousePointer = 1
        Exit Sub
    End If
    
    RChar = Asc(Right(Criteria, 1))
    LChar = Asc(Left(Criteria, 1))
    
    'Search only CAS if a number was entered
    If LChar > 47 And LChar < 59 And RChar > 47 And RChar < 59 Then
        FRMMain!Data1.Recordset.FindFirst "CAS = " & Val(Criteria)
        If Not FRMMain!Data1.Recordset.NoMatch Then
            ' reset the data on the main form and refresh
            FRMMain!LSTSelList.Text = FRMMain!Data1.Recordset("Name")
            FRMMain!TXTFamily.Text = GetFamilyGroup(FRMMain!Data1.Recordset("Chemical Family"))
            FRMMain!TABViewProp.CurrTab = 6
            FRMMain.Refresh
            Screen.MousePointer = 1
            Exit Sub
        End If
        
    Else
        ' do a search based on the criteria the user had entered
        FRMMain!Data1.Recordset.FindNext "[" & Field & "] LIKE '*" & Criteria & "*'"
        If FRMMain!Data1.Recordset.NoMatch Then
            If Field = "Name" Then
                Set DBTbl = DBJetMaster.OpenRecordset("Synonym List", dbOpenSnapshot)
                DBTbl.FindNext "[Name] LIKE '*" & Criteria & "*'"
                If Not FRMMain!Data1.Recordset.NoMatch Then
                    MsgBox "Synonym match found", 48, "Synonym Match"
                    FRMMain!Data1.Recordset.FindNext "[CAS] = '*" & DBTbl("CAS") & "*'"
                    ' denise added 3/27/97
                    If Not FRMMain!Data1.Recordset.NoMatch Then
                        ' reset the data on the main form and refresh
                        FRMMain!LSTSelList.Text = FRMMain!Data1.Recordset("Name")
                        FRMMain!TXTFamily.Text = GetFamilyGroup(FRMMain!Data1.Recordset("Chemical Family"))
                        FRMMain!TABViewProp.CurrTab = 6
                        FRMMain.Refresh
                        Screen.MousePointer = 1
                        Exit Sub
                    End If
                End If
                Set DBTbl = DBJetMaster.OpenRecordset("DIPPR801", dbOpenSnapshot)
                DBTbl.FindNext "[INAM] LIKE '*" & Criteria & "*'"
                If Not FRMMain!Data1.Recordset.NoMatch Then
                    MsgBox "Synonym match found", 48, "Synonym Match"
                    FRMMain!Data1.Recordset.FindNext "[CAS] = '*" & DBTbl("CASN") & "*'"
                    ' denise added 3/27/97
                    
                    FRMMain!TXTFamily.Text = GetFamilyGroup(FRMMain!Data1.Recordset("Chemical Family"))
                    FRMMain!TABViewProp.CurrTab = 6
                    FRMMain.Refresh
                    Screen.MousePointer = 1
                    Exit Sub
                End If
                DBTbl.FindNext "[CNAM] LIKE '*" & Criteria & "*'"
                If Not FRMMain!Data1.Recordset.NoMatch Then
                    ' reset the info on the main form and refresh
                    MsgBox "Synonym match found", 48, "Synonym Match"
                    FRMMain!Data1.Recordset.FindNext "[CAS] = '*" & DBTbl("CASN") & "*'"
                    FRMMain!TXTFamily.Text = GetFamilyGroup(FRMMain!Data1.Recordset("Chemical Family"))
                    FRMMain!TABViewProp.CurrTab = 6
                    FRMMain.Refresh
                    Screen.MousePointer = 1
                    Exit Sub
                End If
            End If
            MsgBox "No more matches found", 48, "No Match"
        Else
            ' make sure the data on the main form is the same
            FRMMain!LSTSelList.Text = FRMMain!Data1.Recordset("Name")
            FRMMain!TXTFamily.Text = GetFamilyGroup(FRMMain!Data1.Recordset("Chemical Family"))
            FRMMain!TABViewProp.CurrTab = 6
            FRMMain.Refresh
            Screen.MousePointer = 1
            Exit Sub
        End If
    End If
    
    ' in case cur_info has changed
    
    FRMMain!Data1.Recordset.FindFirst "CAS = " & Cur_Info.CAS

On Error GoTo Blank_Data1_LIST:
    FRMMain!TXTFamily.Text = GetFamilyGroup(FRMMain!Data1.Recordset("Chemical Family"))
    Cur_Info.name = FRMMain!Data1.Recordset("Name")
    Cur_Info.Formula = FRMMain!Data1.Recordset("Formula")
    Cur_Info.source = FRMMain!Data1.Recordset("Source")
    Cur_Info.Family = FRMMain!Data1.Recordset("Chemical Family")
    Cur_Info.SMILES = FRMMain!Data1.Recordset("Smiles")
    FRMMain!TABViewProp.CurrTab = 6
    FRMMain.Refresh
    
    Screen.MousePointer = 1
    Exit Sub

Blank_Data1_LIST:
    If Err = 3021 Then  'user is searching blank list
        FRMMain!TABViewProp.CurrTab = 6
    Else    'serious unknown screwup
        MsgBox "Prgrom Error " & Err & " occured", 48, "Synonym Match"
        FRMMain!TABViewProp.CurrTab = 6
    End If
    
    Screen.MousePointer = 1
        
End Sub


Private Sub CMDFindPrevious_Click()
    
    Dim Criteria As String
    Dim Field As String
    Dim RChar As Integer
    Dim LChar As Integer
    Dim DBTbl As Recordset
    
    Screen.MousePointer = 11
    
    Criteria = CMBFindStr.Text
    Field = CMBFind.Text
 
    If Field = "Chemical Family" Then
        Criteria = Mid(Criteria, 1, 2)
    End If
       
    If CMBFindStr.Text = "" Then
        MsgBox "No criteria specified", 48, "No Criteria"
        Screen.MousePointer = 1
        Exit Sub
    End If
    
    RChar = Asc(Right(Criteria, 1))
    LChar = Asc(Left(Criteria, 1))
    
    'Search only CAS if a number was entered
    If LChar > 47 And LChar < 59 And RChar > 47 And RChar < 59 Then
        FRMMain!Data1.Recordset.FindPrevious "CAS = " & Val(Criteria)
        ' denise added 3/27/97
        FRMMain!TXTFamily.Text = GetFamilyGroup(FRMMain!Data1.Recordset("Chemical Family"))

    Else
        FRMMain!Data1.Recordset.FindPrevious "[" & Field & "] LIKE '*" & Criteria & "*'"
        If FRMMain!Data1.Recordset.NoMatch Then
            If Field = "Name" Then
                Set DBTbl = DBJetMaster.OpenRecordset("Synonym List", dbOpenSnapshot)
                DBTbl.FindPrevious "[Name] LIKE '*" & Criteria & "*'"
                If Not FRMMain!Data1.Recordset.NoMatch Then
                    MsgBox "Synonym match found", 48, "Synonym Match"
                    FRMMain!Data1.Recordset.FindPrevious "[CAS] = '*" & DBTbl("CAS") & "*'"
                    ' denise added 3/27/97
                    FRMMain!TXTFamily.Text = GetFamilyGroup(FRMMain!Data1.Recordset("Chemical Family"))

                    FRMMain!LSTSelList.Text = FRMMain!Data1.Recordset("Name")
                    Screen.MousePointer = 1
                    Exit Sub
                End If
                Set DBTbl = DBJetMaster.OpenRecordset("DIPPR801", dbOpenSnapshot)
                DBTbl.FindPrevious "[INAM] LIKE '*" & Criteria & "*'"
                If Not FRMMain!Data1.Recordset.NoMatch Then
                    MsgBox "Synonym match found", 48, "Synonym Match"
                    FRMMain!Data1.Recordset.FindPrevious "[CAS] = '*" & DBTbl("CASN") & "*'"
                    ' denise added 3/27/97
                    FRMMain!TXTFamily.Text = GetFamilyGroup(FRMMain!Data1.Recordset("Chemical Family"))

                    FRMMain!LSTSelList.Text = FRMMain!Data1.Recordset("Name")
                    Screen.MousePointer = 1
                    Exit Sub
                End If
                DBTbl.FindNext "[CNAM] LIKE '*" & Criteria & "*'"
                If Not FRMMain!Data1.Recordset.NoMatch Then
                    MsgBox "Synonym match found", 48, "Synonym Match"
                    FRMMain!Data1.Recordset.FindPrevious "[CAS] = '*" & DBTbl("CASN") & "*'"
                    ' denise added 3/27/97
                    FRMMain!TXTFamily.Text = GetFamilyGroup(FRMMain!Data1.Recordset("Chemical Family"))

                    FRMMain!LSTSelList.Text = FRMMain!Data1.Recordset("Name")
                    Screen.MousePointer = 1
                    Exit Sub
                End If
            End If
            MsgBox "No more matches found", 48, "No Match"
        Else
            FRMMain!LSTSelList.Text = FRMMain!Data1.Recordset("Name")
            Screen.MousePointer = 1
            Exit Sub
        End If
    End If
        
    Screen.MousePointer = 1
        
End Sub



Private Sub Command1_Click()

End Sub

Private Sub CMDRestore_List_Click()

FRMMain!Data1.RecordSource = "PEARLS List"
FRMMain!Data1.Refresh

End Sub

Private Sub Form_Deactivate()

    FRMFind.Hide
    
    If FRMMain!Data1.Recordset("CAS") = Cur_Info.CAS Then Exit Sub

    Cur_Info.CAS = FRMMain!Data1.Recordset("CAS")
    Call Recalculate
    Call DisplayProps
    
End Sub

Private Sub Form_Load()
        
    CenterForm Me
    
    'Clear search string criteria box
    CMBFindStr.Clear
    
    'Load all possible search fields
    CMBFind.AddItem "CAS"
    CMBFind.AddItem "Name"
    CMBFind.AddItem "Formula"
    CMBFind.AddItem "Source"
    CMBFind.AddItem "Chemical Family"
    CMBFind.AddItem "Smiles"

    CMBFind.ListIndex = 0
     
End Sub


Private Sub TXTFindStr_KeyPress(KeyAscii As Integer)
    
    'Check for return key
    If KeyAscii = 13 Then
       
       CMDFindNext.SetFocus
    
    End If
    
End Sub


Private Sub Restore_List_Click()
        
FRMMain!Data1.RecordSource = "PEARLS List"
FRMMain!Data1.Refresh

End Sub


