VERSION 5.00
Begin VB.Form frmedit 
   Caption         =   "Database Editor"
   ClientHeight    =   5310
   ClientLeft      =   1545
   ClientTop       =   1995
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5310
   ScaleWidth      =   6690
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frenter 
      Caption         =   "Enter field data"
      Height          =   2415
      Left            =   240
      TabIndex        =   13
      Top             =   2040
      Width           =   6255
      Begin VB.TextBox tbxdata 
         Height          =   375
         Left            =   1080
         TabIndex        =   21
         Top             =   1680
         Width           =   4815
      End
      Begin VB.ComboBox cboopt 
         Height          =   315
         Left            =   1560
         TabIndex        =   20
         Top             =   1200
         Width           =   3735
      End
      Begin VB.ComboBox cbofield 
         Height          =   315
         Left            =   1560
         TabIndex        =   19
         Top             =   720
         Width           =   3735
      End
      Begin VB.ComboBox cbotable 
         Height          =   315
         Left            =   1560
         TabIndex        =   18
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label lbldata 
         Caption         =   "Field Data"
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label lblopt 
         Alignment       =   1  'Right Justify
         Caption         =   "editing options"
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lblfield 
         Alignment       =   1  'Right Justify
         Caption         =   "field"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lbltable 
         Alignment       =   1  'Right Justify
         Caption         =   "table"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdaccept 
      Caption         =   "&Accept entry"
      Height          =   495
      Left            =   840
      TabIndex        =   8
      Top             =   4680
      Width           =   1815
   End
   Begin VB.CommandButton cmdone 
      Caption         =   "&Done"
      Height          =   495
      Left            =   3960
      TabIndex        =   7
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Frame fredit 
      Caption         =   "Editing..."
      Height          =   735
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   6255
      Begin VB.Label lblpromptchem 
         Alignment       =   2  'Center
         Caption         =   "use the chemical browser to select a chemical or 'new' to enter a new chemical  name"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   5775
      End
      Begin VB.Label lblselname 
         Caption         =   "lblchemname"
         Height          =   375
         Left            =   1680
         TabIndex        =   11
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label lblselcas 
         Caption         =   "lblselcas"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame fradd 
      Caption         =   "Enter Chemical Information"
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      Begin VB.TextBox tbxfamily 
         Height          =   375
         Left            =   3240
         TabIndex        =   26
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox tbxstructure 
         Height          =   375
         Left            =   840
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CommandButton cmdacceptadd 
         Caption         =   "add chemical"
         Height          =   375
         Left            =   4560
         TabIndex        =   22
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox tbxsmiles 
         Height          =   375
         Left            =   3120
         TabIndex        =   3
         Text            =   "Text3"
         Top             =   360
         Width           =   3015
      End
      Begin VB.TextBox tbxname 
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Text            =   "Text2"
         Top             =   840
         Width           =   4815
      End
      Begin VB.TextBox tbxcas 
         Height          =   375
         Left            =   720
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblfamily 
         Caption         =   "Family"
         Height          =   255
         Left            =   2640
         TabIndex        =   25
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label lblstructure 
         Caption         =   "Formula"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label lblsmiles 
         Caption         =   "SMILES"
         Height          =   255
         Left            =   2400
         TabIndex        =   6
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblname 
         Caption         =   "Chemical Name"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblcas 
         Caption         =   "CAS #"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.Menu mnuchem 
      Caption         =   "&Chemical"
      Begin VB.Menu mnunew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuchembr 
         Caption         =   "&Browse"
      End
   End
   Begin VB.Menu mnutools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuedwiz 
         Caption         =   "&Edit wizard"
      End
   End
End
Attribute VB_Name = "frmedit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboprop_Change()

End Sub

Private Sub cbofield_Click()

    Call update_data_box
End Sub


Private Sub cbotable_Click()

    Call update_field_options
    Call update_data_box
    
End Sub


Private Sub cmdaccept_Click()

    ' for now the only option is user input
   ' Dim localdb As Database
    Dim localtable As Recordset
    Dim idfield As String
    Dim type_string As String
    Dim Criteria As Long
    Dim expected_type As Integer
    Dim entered_type As Integer
    Dim found_field As Integer
    Dim I As Integer
    Dim match As Boolean
    
    If frmedit!fradd.Visible = True Then
        If Trim(frmedit!tbxcas.Text) = "" Then
            MsgBox ("select a chemical to edit")
            Exit Sub
        ElseIf Trim(frmedit!tbxdata.Text) = "" Then
            MsgBox ("enter data for field")
            Exit Sub
        Else
            Criteria = CLng(frmedit.tbxcas.Text)
        End If
    Else
        If Not IsNumeric(Trim(frmedit!lblselcas)) Then
            MsgBox ("select a chemical to edit")
            Exit Sub
        Else
            Criteria = CLng(Trim(frmedit!lblselcas.Caption))
        End If
    End If
    ' first, if it's a global identifier, update that
    If Trim(frmedit!cbofield.Text) = "CAS" Then
        selected_cas = CLng(frmedit!tbxdata.Text)
    ElseIf Trim(frmedit!cbofield.Text) = "Name" Then
        selected_name = frmedit!tbxdata.Text
    ElseIf Trim(frmedit!cbofield.Text) = "Formula" Then
        selected_structure = frmedit!tbxdata.Text
    ElseIf Trim(frmedit!cbofield.Text) = "Smiles" Then
        selected_smiles = frmedit!tbxdata.Text
    End If
    
    If frmedit!cbotable.Text = "DIPPR911" Then
        idfield = "Cas #"
    Else
        idfield = "CAS"
    End If
    ' if it's not in the main list (ie PEARLS List or fexp2) add it there
    ' add code here DENISE
    Select Case Trim(frmedit!cboopt.Text)
        Case "user input"
            'Set localdb = OpenDatabase(dbpath & "\" & curname, False, False)
            Set localtable = chembrowsedb.OpenRecordset(Trim(frmedit!cbotable.Text), dbOpenDynaset)
            ' first find the chemical if its there
            If idfield = "Cas #" Then
                localtable.FindFirst "'" & idfield & "' = " & Val(Criteria)
            Else
                localtable.FindFirst idfield & " = " & Val(Criteria)
            End If
            ' check and make sure the data type is compatible
            For I = 0 To localtable.Fields.count - 1
                If Trim(localtable.Fields(I).name) = Trim(frmedit!cbofield.Text) Then
                    found_field = I
                    Exit For
                End If
            Next I
                
            expected_type = localtable.Fields(found_field).Type
            match = confirm_type(expected_type, entered_type, frmedit!tbxdata.Text)
            If match = False Then
                GoTo wrong_data_type
            End If
            If localtable.NoMatch Then
                If Trim(frmedit!tbxcas.Text) = "" Then
                    MsgBox ("select a chemical to edit")
                    Exit Sub
                End If
                ' add a new one
                localtable.AddNew
                localtable.Fields(idfield) = Criteria
                localtable.Fields(found_field) = frmedit!tbxdata.Text
                localtable.Update
                
            Else
                localtable.Edit
                localtable.Fields(found_field) = frmedit!tbxdata.Text
                localtable.Update
            End If
    End Select
    localtable.Close
    'localdb.Close
    ' if it's one of the global chem describers, update them
    
    Call update_data_box
    MsgBox ("database successfully updated")
    Exit Sub
error_message:
    MsgBox ("error in updating database")
    Exit Sub
wrong_data_type:
    Select Case expected_type
        Case dbDate
            type_string = "Date"
        Case dbInteger
            type_string = "Integer"
        Case dbLong
            type_string = "Long"
        Case dbDouble
            type_string = "Double"
        Case dbText
            type_string = "Text"
        Case dbMemo
            type_string = "Memo"
        Case Else
            type_string = "???"
    End Select
    MsgBox ("incompatible data type entered, this field requires a " & type_string)
    Exit Sub
End Sub

Private Sub cmdacceptadd_Click()

    Dim success As Boolean
    Dim answer As Integer
    If Trim(tbxcas.Text) = "" Or Trim(tbxname.Text) = "" Then
        MsgBox ("The database requires a name and CAS number for each chemical")
        Exit Sub
    ElseIf Not IsNumeric(tbxcas.Text) Then
        MsgBox ("The database requires a numeric value for the CAS field")
        Exit Sub
    ElseIf Trim(tbxsmiles.Text) = "" Then
        MsgBox ("The database requires a smiles string for each chemical")
        Exit Sub
    ElseIf Trim(tbxstructure.Text) = "" Then
        MsgBox ("The database requires a formula for each chemical")
        Exit Sub
    ElseIf Trim(tbxfamily.Text) = "" Then
        answer = MsgBox("Warning:  no chemical family code has been entered, continue?", vbYesNo)
        If answer = vbNo Then
            Exit Sub
        End If
    End If
    selected_cas = CLng(frmedit!tbxcas.Text)
    selected_name = Trim(frmedit!tbxname.Text)
    selected_smiles = Trim(frmedit!tbxsmiles.Text)
    selected_structure = Trim(frmedit!tbxstructure.Text)
    selected_family = Trim(frmedit!tbxfamily.Text)
    
    Screen.MousePointer = 11
    success = add_chemical_info
    If success = False Then
        MsgBox ("an error occurred adding the chemical to the database")
    Else
        ' if it was successful, enable the editing options
        frmedit.Height = 4810
        frmedit!frenter.Top = 960
        frmedit!cmdaccept.Top = 3550
        frmedit!cmdone.Top = 3550
        frmedit!cbotable.Enabled = True
        frmedit!cbofield.Enabled = True
        frmedit!cboopt.Enabled = True
        frmedit!tbxdata.Enabled = True
        frmedit!cmdaccept.Enabled = True
        frmedit!lbltable.Enabled = True
        frmedit!lblfield.Enabled = True
        frmedit!lblopt.Enabled = True
        frmedit!lbldata.Enabled = True
        frmedit!fradd.Visible = False
        
        frmedit!fredit.Visible = True
        frmedit!lblpromptchem.Visible = False
        frmedit!lblselname.Visible = True
        frmedit!lblselcas.Visible = True
        frmedit!lblselname.Caption = selected_name
        frmedit!lblselcas.Caption = CStr(selected_cas)
    End If
    Screen.MousePointer = 1
End Sub

Private Sub cmdone_Click()

    On Error GoTo after_close   ' takes care of if db not open
    selected_name = ""
    selected_cas = 0
    selected_structure = ""
    selected_smiles = ""
    selected_rings = -1
    dbstatus = STATUS_CHANGED
after_close:
    
    frmedit.Hide
    'frmdbman!frwait.Caption = "Ready"
    'frmdbman!lblwait.Caption = "use menu items to edit " & curname
    'frmdbman!frwait.Visible = True
    'frmdbman.Show
    Unload Me
End Sub


Private Sub Data1_Validate(Action As Integer, Save As Integer)

End Sub

Private Sub dbgrdedit_Click()

End Sub

Private Sub Combo2_Change()

End Sub

Private Sub mnuchembr_Click()

'    If Trim(curapp) <> "" Then
        Screen.MousePointer = 11
        Call load_chem_browse_info
        Screen.MousePointer = 1
        frmchembrowse.Show 1
        If frmedit!fredit.Visible = False Then
            frmedit.Height = 5000
            frmedit!frenter.Top = 1100
            frmedit!cmdaccept.Top = 3680
            frmedit!cmdone.Top = 3680
            frmedit.fradd.Visible = False
            frmedit!cbotable.Enabled = True
            frmedit!cbofield.Enabled = True
            frmedit!cboopt.Enabled = True
            frmedit!tbxdata.Enabled = True
            frmedit!lbltable.Enabled = True
            frmedit!lblfield.Enabled = True
            frmedit!lblopt.Enabled = True
            frmedit!lbldata.Enabled = True
            frmedit!cmdaccept.Enabled = True
            frmedit!fredit.Visible = True
            frmedit!fradd.Visible = False
            frmedit.Caption = "Database Editor: " & curpath & curname
        End If
        If Trim(selected_name) <> "" Then
            frmedit!lblselname.Visible = True
            frmedit!lblselcas.Visible = True
            frmedit!lblselname.Caption = selected_name
            frmedit!lblselcas.Caption = CStr(selected_cas)
            frmedit!lblpromptchem.Visible = False
            frmedit.Refresh
        End If
'    Else
'        MsgBox ("chemical browser only available for pearls databases")
'    End If
End Sub

Private Sub mnuchemimport_Click()

'On Error GoTo cancel_error

    'Show user available files
    'frmedit!cdbrowse.DialogTitle = "Import Database Selection"
    'frmedit!cdbrowse.CancelError = True
    'frmedit!cdbrowse.Filter = "(*.mdb)|*.mdb"
    'frmedit!cdbrowse.FilterIndex = 1
    'frmedit!cdbrowse.InitDir = App.path
    'frmedit!cdbrowse.DefaultExt = "mdb"
    'frmedit!cdbrowse.Action = 1
            
    ' if they chose something, use it for the master
    ' separate the name from the path
    'tempname = Trim(frmedit!cdbrowse.filename)
    'Call do_chemical_import
    
'general_import_error:
 '   If Error = cdlCancel Then
 '       Exit Sub
  '  End If
  '  MsgBox ("Import feature not yet implemented")
   ' Exit Sub
    
End Sub

Private Sub mnuchemrename_Click()

End Sub

Private Sub mnuedwiz_Click()

    ' first check that a chemical has been selected
    If Trim(selected_name) = "" Or selected_cas = 0 Then
        MsgBox ("select a chemical to edit")
        Exit Sub
    End If
    
    Call load_edit_wizard_form
    frmeditwizard.Show 1
    
End Sub

Private Sub mnuimport_Click()

   
   ' Select Case curapp
    '    Case "PEARLS", "Pearls"
            
    '        frmimport.Show
    '    Case "Block 5", "block 5"
    '        frmimport.Show
    'End Select
End Sub

Private Sub mnunew_Click()

    Call load_add_chem_info
    
End Sub

Private Sub mnutools_Click()

    If Trim(tbxname.Text) = "" Then
        tbxname.Text = "unnamed"
    End If
End Sub



Public Function get_cnum(cas_string As String) As Long
    ' just a dummy value for now DENISE change
    get_cnum = 1

End Function
