VERSION 5.00
Begin VB.Form frmstruct 
   Caption         =   "Structure Dissasembly Software"
   ClientHeight    =   5010
   ClientLeft      =   885
   ClientTop       =   2385
   ClientWidth     =   8220
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5010
   ScaleWidth      =   8220
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdadd 
      Caption         =   "&Add to database"
      Height          =   375
      Left            =   3120
      TabIndex        =   50
      Top             =   4560
      Width           =   2175
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "&Done"
      Height          =   375
      Left            =   5760
      TabIndex        =   5
      Top             =   4560
      Width           =   2055
   End
   Begin VB.CommandButton cmdrun 
      Caption         =   "&Run"
      Height          =   375
      Left            =   630
      TabIndex        =   4
      Top             =   4560
      Width           =   2055
   End
   Begin VB.Frame frgroups 
      Caption         =   "Groups Found"
      Height          =   2175
      Left            =   240
      TabIndex        =   1
      Top             =   2280
      Width           =   7815
      Begin VB.Label lblgrno 
         Caption         =   "Label1"
         Height          =   255
         Index           =   19
         Left            =   7200
         TabIndex        =   48
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label lblgrno 
         Caption         =   "Label1"
         Height          =   255
         Index           =   18
         Left            =   7200
         TabIndex        =   47
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblgrno 
         Caption         =   "Label1"
         Height          =   255
         Index           =   17
         Left            =   7200
         TabIndex        =   46
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lblgrno 
         Caption         =   "Label1"
         Height          =   255
         Index           =   16
         Left            =   7200
         TabIndex        =   45
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lblgrno 
         Caption         =   "Label1"
         Height          =   255
         Index           =   15
         Left            =   7200
         TabIndex        =   44
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblgrno 
         Caption         =   "Label1"
         Height          =   255
         Index           =   14
         Left            =   7200
         TabIndex        =   43
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblgrno 
         Caption         =   "Label1"
         Height          =   255
         Index           =   13
         Left            =   4680
         TabIndex        =   42
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label lblgrno 
         Caption         =   "Label1"
         Height          =   255
         Index           =   12
         Left            =   4680
         TabIndex        =   41
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label lblgrno 
         Caption         =   "Label1"
         Height          =   255
         Index           =   11
         Left            =   4680
         TabIndex        =   40
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblgrno 
         Caption         =   "Label1"
         Height          =   255
         Index           =   10
         Left            =   4680
         TabIndex        =   39
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lblgrno 
         Caption         =   "Label1"
         Height          =   255
         Index           =   9
         Left            =   4680
         TabIndex        =   38
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lblgrno 
         Caption         =   "Label1"
         Height          =   255
         Index           =   8
         Left            =   4680
         TabIndex        =   37
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblgrno 
         Caption         =   "Label1"
         Height          =   255
         Index           =   7
         Left            =   4680
         TabIndex        =   36
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblgrno 
         Caption         =   "Label1"
         Height          =   255
         Index           =   6
         Left            =   2040
         TabIndex        =   35
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label lblgrno 
         Caption         =   "Label1"
         Height          =   255
         Index           =   5
         Left            =   2040
         TabIndex        =   34
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label lblgrno 
         Caption         =   "Label1"
         Height          =   255
         Index           =   4
         Left            =   2040
         TabIndex        =   33
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblgrno 
         Caption         =   "Label1"
         Height          =   255
         Index           =   3
         Left            =   2040
         TabIndex        =   32
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lblgrno 
         Caption         =   "Label1"
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   31
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lblgrno 
         Caption         =   "Label1"
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   30
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblgrno 
         Caption         =   "Label1"
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   29
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lbldisplay 
         Caption         =   "junk"
         Height          =   255
         Index           =   19
         Left            =   5280
         TabIndex        =   28
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label lbldisplay 
         Caption         =   "junk"
         Height          =   255
         Index           =   18
         Left            =   5280
         TabIndex        =   27
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label lbldisplay 
         Caption         =   "junk"
         Height          =   255
         Index           =   17
         Left            =   5280
         TabIndex        =   26
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lbldisplay 
         Caption         =   "junk"
         Height          =   255
         Index           =   16
         Left            =   5280
         TabIndex        =   25
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label lbldisplay 
         Caption         =   "junk"
         Height          =   255
         Index           =   15
         Left            =   5280
         TabIndex        =   24
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lbldisplay 
         Caption         =   "junk"
         Height          =   255
         Index           =   14
         Left            =   5280
         TabIndex        =   23
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lbldisplay 
         Caption         =   "junk"
         Height          =   255
         Index           =   13
         Left            =   2760
         TabIndex        =   22
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label lbldisplay 
         Caption         =   "junk"
         Height          =   255
         Index           =   12
         Left            =   2760
         TabIndex        =   21
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label lbldisplay 
         Caption         =   "junk"
         Height          =   255
         Index           =   11
         Left            =   2760
         TabIndex        =   20
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label lbldisplay 
         Caption         =   "junk"
         Height          =   255
         Index           =   10
         Left            =   2760
         TabIndex        =   19
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lbldisplay 
         Caption         =   "junk"
         Height          =   255
         Index           =   9
         Left            =   2760
         TabIndex        =   18
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label lbldisplay 
         Caption         =   "junk"
         Height          =   255
         Index           =   8
         Left            =   2760
         TabIndex        =   17
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lbldisplay 
         Caption         =   "junk"
         Height          =   255
         Index           =   7
         Left            =   2760
         TabIndex        =   16
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lbldisplay 
         Caption         =   "junk"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   15
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label lbldisplay 
         Caption         =   "junk"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   14
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label lbldisplay 
         Caption         =   "junk"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   13
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label lbldisplay 
         Caption         =   "junk"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lbldisplay 
         Caption         =   "junk"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label lbldisplay 
         Caption         =   "junk"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lbldisplay 
         Caption         =   "junk"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame frselection 
      Caption         =   "Disassembly Settings"
      Height          =   1695
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   7815
      Begin VB.TextBox tbxname 
         Height          =   375
         Left            =   1560
         TabIndex        =   49
         Top             =   240
         Width           =   5055
      End
      Begin VB.ComboBox cboschtype 
         Height          =   315
         Left            =   1560
         TabIndex        =   8
         Text            =   "Combo1"
         Top             =   1200
         Width           =   3135
      End
      Begin VB.TextBox tbxsmiles 
         Height          =   375
         Left            =   1560
         TabIndex        =   3
         Top             =   720
         Width           =   5055
      End
      Begin VB.Label lblschtype 
         Alignment       =   1  'Right Justify
         Caption         =   "Search Type"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblname 
         Alignment       =   1  'Right Justify
         Caption         =   "Chemical Name"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblsmiles 
         Alignment       =   1  'Right Justify
         Caption         =   "SMILES Notation"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.Label lblstatus 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   1080
      TabIndex        =   51
      Top             =   120
      Width           =   5895
   End
   Begin VB.Menu mnuchem 
      Caption         =   "&Chemical"
      Begin VB.Menu mnubrowse 
         Caption         =   "&browse"
      End
   End
   Begin VB.Menu mnuedit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuedgr 
         Caption         =   "&groups"
      End
   End
   Begin VB.Menu mnutools 
      Caption         =   "&Tools"
      Begin VB.Menu mnumatch 
         Caption         =   "&match"
         Begin VB.Menu mnumatchsmiles 
            Caption         =   "smiles"
         End
         Begin VB.Menu mnumatchgroups 
            Caption         =   "groups"
         End
      End
   End
End
Attribute VB_Name = "frmstruct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdadd_Click()

    ' this function adds the groups to the database for this chemical
    ' the only table in the database that's
    ' designed for this is UNIFAC, so we'll only
    ' do it for those groups
    
    Dim localtable As Dynaset
    Dim i As Integer
    Dim fieldname As String
    On Error GoTo error_in_update
    Set localtable = chembrowsedb.OpenRecordset("UNIFAC Groups", dbOpenDynaset)
    localtable.FindFirst "CAS = " & Val(selected_cas)
    If Not localtable.NoMatch Then
        localtable.Edit
        If selected_rings <> -1 Then
            localtable("RG") = selected_rings
        Else
            Call set_rings
            If selected_rings = -1 Then
                MsgBox ("Warning: number of rings not updated")
            End If
            
        End If
        i = 0
        While cur_chem_groups(i) <> -1 And i < 9
            fieldname = "G" & CStr(i + 1)
            localtable(fieldname) = cur_chem_groups(i)
            fieldname = "N" & CStr(i + 1)
            localtable(fieldname) = num_cur_chem_groups(i)
            i = i + 1
        Wend
        localtable("MX") = i
        localtable.Update
    Else
        localtable.AddNew
        localtable("CAS") = selected_cas
        If selected_rings <> -1 Then
            localtable("RG") = selected_rings
        Else
            Call set_rings
            If selected_rings = -1 Then
                MsgBox ("Warning: number of rings not updated")
            End If
        End If
        While cur_chem_groups(i) <> -1 And i < 9
            fieldname = "G" & CStr(i + 1)
            localtable(fieldname) = cur_chem_groups(i)
            fieldname = "N" & CStr(i + 1)
            localtable(fieldname) = num_cur_chem_groups(i)
            i = i + 1
        Wend
        localtable("MX") = i
        localtable.Update
    End If
    
    localtable.Close
            
    MsgBox ("Database successfully updated")
    Exit Sub
error_in_update:
    MsgBox ("An error occurred updating the database")
    End Sub

Public Sub set_rings()
Dim MyValue As String
Dim i As Integer

    MyValue = InputBox("Please enter the number of rings for this chemical", selected_name, "")
        If Trim(MyValue) <> "" And IsNumeric(MyValue) Then
            selected_rings = CInt(MyValue)
            ' update the edit wizard form if necessary
            For i = 0 To MAX_INPUTS_EACH - 1
                If frmeditwizard!lblinputprop(i).Caption Like "*ring*" Then
                    frmeditwizard!tbxinputprop(i).Text = selected_rings
                End If
            Next i
        Else
            selected_rings = -1
        End If
End Sub


Private Sub cmdexit_Click()

    Unload Me
End Sub

Private Sub cmdmatch_Click()

Dim match_smiles_name As String
Dim match_group_name As String
Dim match_choice As String
Dim match_smiles_smiles As String
Dim match_group_smiles As String
Dim match_smiles_cas As Long
Dim match_group_cas As Long

Dim group_match As Boolean
Dim smiles_match As Boolean
group_match = False
smiles_match = False

' first find out whether we're matching the smiles or
' the groups

If Trim(frmstruct.tbxsmiles.Text) = "" Then
    If Trim(frmstruct.lbldisplay(0).Caption) <> "" Then
        ' find any chemical that matches the groups displayed
        'group_match = match_groups(match_group_name, match_group_cas, match_group_smiles)
    Else
        MsgBox ("this function requires a smiles string and/or groups")
        Exit Sub
    End If
Else
    If Trim(frmstruct.lbldisplay(0).Caption) <> "" Then
        ' try matching both smiles and groups
        'smiles_match = match_smiles(match_smiles_name, match_smiles_cas, match_smiles_smiles)
        'group_match = match_groups(match_group_name, match_group_cas, match_group_smiles)
    Else
        ' try matching the smiles
        'smiles_match = match_smiles(match_smiles_name, match_smiles_cas, match_smiles_smiles)
    End If
End If
' now report what we found
If smiles_match = True And group_match = True Then
    If match_smiles_cas = match_group_cas Then
        ' match is good, just update the screen and globals
    Else
        ' different matches, find out which we should use
        match_choice = InputBox("Two different matches found, enter s to match smiles, g to match groups", "Select a criteria for match", "g", 100, 100)
        If Trim(match_choice) = "s" Then
            ' match the smiles
            'cur_cas = match_smiles_cas
            'cur_name = match_smiles_name
            'cur_smiles = match_smiles_smiles
            tbxsmiles = selected_smiles
        Else
            ' match the groups
           ' cur_cas = match_group_cas
           ' cur_name = match_group_name
           ' cur_smiles = match_group_smiles
            tbxsmiles = selected_smiles
        End If
    End If
ElseIf smiles_match = True Then
    ' then update name etc
    'cur_name = match_smiles_name
   ' cur_cas = match_smiles_cas
   ' cur_smiles = match_smiles_smiles
    tbxsmiles = selected_smiles
ElseIf group_match = True Then
    ' then update name etc.
    'cur_name = match_group_name
    'cur_cas = match_group_cas
    'cur_smiles = match_group_smiles
    tbxsmiles = selected_smiles
Else
    MsgBox ("no matches found")
End If
End Sub

Private Sub cmdrun_Click()
Dim SearchResult As Byte
Dim spaces As String
Dim Message As String
Dim i As Integer
    Screen.MousePointer = 11
    
    'Call set_group_array
    'return_status = do_structure_disassembly(search_type)

    SearchResult = Run_Mosdap(Trim(frmstruct!tbxsmiles.Text), global_groupfile, cboschtype.ListIndex)
    
    Select Case SearchResult
        Case 0
            Message = "Unable to disassemble " & Trim(frmstruct!tbxsmiles.Text)
        Case 1
            Message = "Successfully disassembled"
        Case 2
            Message = "Partially disassembled"
        Case Else
            Message = "An error occurred in the code while disassembling"
    End Select
    
    For i = 0 To MAX_GROUPS_PER_CHEM - 1
        If cur_chem_groups(i) > 0 Then
            If cur_chem_groups(i) < 10 Then
                spaces = "  "
            ElseIf cur_chem_groups(i) < 100 Then
                spaces = " "
            End If
            lbldisplay(count) = spaces & cur_chem_groups(i) & ". " & group_smiles(cur_chem_groups(i) - 1)
            lblgrno(count) = num_cur_chem_groups(i)
        Else
            lbldisplay(i) = ""
            lblgrno(i) = ""
        End If
    Next i
    lblstatus.Caption = Message
    Screen.MousePointer = 1
End Sub


Private Sub mnubrowse_Click()
   
    Call load_chem_browse_info
    frmchembrowse!cmdAccept.Enabled = False
    frmchembrowse.cmdexit.Caption = "&Done"
    frmchembrowse.Show 1
    frmchembrowse.cmdexit.Caption = "&Cancel"
    frmchembrowse!cmdAccept.Enabled = True
End Sub


Private Sub mnuedgr_Click()
    Screen.MousePointer = 11
    Call load_edit_groups_form
    Screen.MousePointer = 1
    frmeditgr.Show 1
End Sub

Private Sub mnumatchgroups_Click()

' this option will look for a group match in the database
' if found, the rename screen will appear to allow the user
' to accept the chemical name (if not already the same)
Dim match_groups_name As String
Dim match_groups_smiles As String
Dim match_groups_cas As Long
Dim groups_match As Boolean

    groups_match = False

' first check that there's a smiles string there
    If Trim(frmstruct!lbldisplay(0).Caption) = "" Then
        MsgBox ("no current groups")
        Exit Sub
    End If
    'groups_match = match_groups(match_groups_name, match_groups_cas, match_groups_smiles)
    If groups_match = True Then
        'frmrename.tbxchemname = match_groups_name
        'frmrename.tbxcas = match_groups_cas
        'frmrename.Show 1
    Else
        MsgBox ("no matches found")
    End If
          
End Sub

Private Sub mnumatchsmiles_Click()

' this function will try to find a chemical in the database
' with the current smiles
' if found, the rename screen will be called to allow the user to
' change the name of the current chemical (if not already the same)
Dim match_smiles_name As String
Dim match_smiles_smiles As String
Dim match_smiles_cas As Long
Dim smiles_match As Boolean

    smiles_match = False

' first check that there's a smiles string there
    If Trim(frmstruct!tbxsmiles.Text) = "" Then
        MsgBox ("no current smiles string")
        Exit Sub
    End If
    'smiles_match = match_smiles(match_smiles_name, match_smiles_cas, match_smiles_smiles)
    If smiles_match = True Then
        'frmrename.tbxchemname = match_smiles_name
        'frmrename.tbxcas = match_smiles_cas
        'frmrename.Show 1
    Else
        MsgBox ("no matches found")
    End If
          
End Sub



