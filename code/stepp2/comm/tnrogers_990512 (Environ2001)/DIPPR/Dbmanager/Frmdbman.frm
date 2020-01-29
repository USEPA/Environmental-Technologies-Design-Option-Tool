VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmdbman 
   Caption         =   "Database Manager"
   ClientHeight    =   2730
   ClientLeft      =   1230
   ClientTop       =   3795
   ClientWidth     =   8205
   Icon            =   "Frmdbman.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2730
   ScaleWidth      =   8205
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frwait 
      Caption         =   "Please Wait..."
      Height          =   1875
      Left            =   3480
      TabIndex        =   5
      Top             =   120
      Width           =   4575
      Begin VB.Label lblwait 
         Alignment       =   2  'Center
         Height          =   1335
         Left            =   330
         TabIndex        =   6
         Top             =   360
         Width           =   3945
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame frsoftware 
      Caption         =   "PPMS Software"
      Height          =   735
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   3135
      Begin VB.Label lblapps 
         Caption         =   "PEARLS"
         Height          =   285
         Left            =   315
         TabIndex        =   7
         Top             =   315
         Width           =   1320
      End
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   5025
      TabIndex        =   3
      Top             =   2160
      Width           =   1800
   End
   Begin VB.Frame fravail 
      Caption         =   "Available Databases"
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Width           =   3135
      Begin VB.ComboBox cbonames 
         Height          =   315
         Left            =   270
         TabIndex        =   2
         Text            =   "Combo2"
         Top             =   270
         Width           =   2610
      End
   End
   Begin VB.Frame frformat 
      Caption         =   "Database Format"
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   3135
      Begin VB.Label lblformats 
         Caption         =   "MASTER"
         Height          =   330
         Left            =   315
         TabIndex        =   8
         Top             =   315
         Width           =   1140
      End
   End
   Begin MSComDlg.CommonDialog cddbman 
      Left            =   2760
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuaddexistdb 
         Caption         =   "Add &Existing Database"
      End
      Begin VB.Menu mnuadddb 
         Caption         =   "Add &New Database"
      End
      Begin VB.Menu mnuremovedb 
         Caption         =   "&Remove Database"
      End
      Begin VB.Menu mnuspace1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuchemical 
      Caption         =   "&Chemical"
      Begin VB.Menu mnubrowse 
         Caption         =   "&Browse"
      End
      Begin VB.Menu mnudash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuadd 
         Caption         =   "&Add Chemical"
      End
      Begin VB.Menu mnurmchem 
         Caption         =   "&Remove Chemical"
      End
      Begin VB.Menu mnuedchem 
         Caption         =   "&Edit Chemical "
      End
      Begin VB.Menu mnuspace2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditWiz 
         Caption         =   "Edit &Wizard"
      End
      Begin VB.Menu mnuShred 
         Caption         =   "&Method Shredder"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuabout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmdbman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub update_names_box()
    Dim i As Integer
    cbonames.Clear
    For i = 0 To dbman_apps - 1
        cbonames.AddItem dbman_(i, 0)
    Next i
    cbonames.ListIndex = 0
End Sub

Private Sub cbonames_Click()
' *paul
    If Trim(curname) = "" Then
        dbstatus = STATUS_CLOSED
    ElseIf Trim(curname) <> Trim(cbonames.Text) & ".mdb" Then
        dbstatus = STATUS_CHANGED
    End If
    Call Set_CurName(cbonames.Text)
    
    frwait.Caption = "Ready"
    lblwait.Caption = "Use menu items to edit " & curname
    frwait.Visible = True
End Sub

Private Sub cmdexit_Click()
' *paul
    Unload Me
End Sub

Private Sub Form_Load()
' *paul
    Call update_names_box
End Sub

Private Sub Form_Unload(Cancel As Integer)
' *paul
    Dim answer As Integer
    answer = MsgBox("Quit Database Manager?", vbYesNo)
    If answer = vbYes Then
        On Error Resume Next
        If dbstatus = STATUS_OPEN Or dbstatus = STATUS_CHANGED Then
            On Error Resume Next
            chembrowsedb.Close
            dbstatus = STATUS_CLOSED
        End If
        Call write_dbman_file
    Else
        Cancel = True
    End If
End Sub

Function Load_File(Caption As String) As String
' *paul
    On Error GoTo Cancel_Error
    
    cddbman.CancelError = True
    cddbman.Filter = "*.mdb, *dbm|*.mdb;*.dbm"
    cddbman.FilterIndex = 1
    cddbman.DialogTitle = Caption
    cddbman.InitDir = AppPath
    cddbman.filename = ""
    cddbman.DefaultExt = "mdb"
    cddbman.Action = 2
    Load_File = cddbman.filename
    Exit Function
    
Cancel_Error:
    Load_File = ""
End Function

Private Sub mnuabout_Click()
    frmSplash.Show 1
End Sub

Private Sub mnuadd_Click()
' *paul
    Dim db_extension As String
    If LCase(curname) = LCase("master.mdb") Then
        MsgBox ("You can't add a chemical to " & curname)
        Exit Sub
    End If
    
    Call load_add_chem_info
    Call load_form_edit_info
    frmedit.Show
End Sub

Private Sub mnuadddb_Click()
' *paul
    Dim newfilename As String
    Dim file_extension As String
    Dim tempchar As String
    Dim success As Boolean
    
    If MsgBox("Create a new database with the format of the PEARLS MASTER database?", vbYesNo) = vbYes Then
        
        ' call browser to set the name of the new one (extension dbm)
        newfilename = Load_File("Set Name of New Database")
        If newfilename = "" Then
            Exit Sub
        End If
        
        frwait.Caption = "Please Wait..."
        If create_database_copy(newfilename) = True Then
            ' curname is now set by previous function
            Call Set_CurName(newfilename)
            Call update_globals(1)
            Call update_names_box
        Else
            MsgBox ("unable to add database")
        End If
        
    End If
    Exit Sub
    
Cancel_Error:
'    frmdbman!frwait.Caption = "Error"
'    frmdbman!lblwait.Caption = "Error adding " & newfilename
'    frmdbman!frwait.Visible = True
'    Exit Sub
End Sub

Private Sub mnuaddexistdb_Click()
' *paul
    Dim name As String
    
    name = Load_File("Select Name of Database to Add")
    If name = "" Then
        Exit Sub
    End If
    frwait.Caption = "Please Wait"
    lblwait.Caption = "adding " & name & "..."
    frwait.Visible = True
    frwait.Refresh
    ' set curname and dbpath ??
    Call Set_CurName(name)
    Call update_globals(1)
    Call update_names_box
End Sub

Private Sub mnubrowse_Click()

    ' here we're not browsing from the edit form so disable the accept button
    
    Call load_chem_browse_info
    frmchembrowse!cmdAccept.Enabled = False
    frmchembrowse.cmdexit.Caption = "&Done"
    frmchembrowse.Show 1
    frmchembrowse.cmdexit.Caption = "&Cancel"
    frmchembrowse!cmdAccept.Enabled = True
End Sub

Private Sub mnuedchem_Click()
    Dim db_extension As String
    If LCase(curname) = LCase("master.mdb") Then
        MsgBox ("You can't edit a chemical in " & curname)
        Exit Sub
    End If
    
    Call load_edit_chem_info
    Call load_form_edit_info
    frmedit!fredit.Refresh
    frmedit.Show
End Sub

Private Sub mnuEditWiz_Click()
    ' first check that a chemical has been selected
'    If LCase(curname) = LCase("master.mdb") Then
'        MsgBox ("You can't edit the " & curname)
'        Exit Sub
'    End If
    
    Call load_chem_browse_info
    frmchembrowse.Show 1
    
    If Trim(selected_name) = "" Or selected_cas = 0 Then
        Exit Sub
    End If
    
    Call load_edit_wizard_form
    frmeditwizard.Show 1
End Sub

Private Sub mnuexit_Click()
' *paul
    Unload Me
End Sub

Private Sub mnuremovedb_Click()
    Dim newfilename As String
    Dim answer As Integer
    Dim success As Boolean
    Dim file_extension As String
    If LCase(curname) = LCase("master.mdb") Then
        MsgBox ("You can't remove the file " & curname)
        Exit Sub
    End If
    answer = MsgBox("Remove " & curname & " ?", vbYesNo)
    
'    newfilename = Trim(frmdbman!cbonames.Text)
    def_modified = True
'    curname = newfilename
    Call update_globals(-1)
    Call update_names_box
    
End Sub

Private Sub mnurmchem_Click()
    Dim db_extension As String
    If LCase(curname) = LCase("master.mdb") Then
        MsgBox ("You can't remove a chemical from " & curname)
        Exit Sub
    End If
    Call load_remove_chem_info
    Call load_form_edit_info
    frmedit.Show
End Sub

Private Sub mnuShred_Click()
    ' first check that a chemical has been selected
    Call load_chem_browse_info
    frmchembrowse.Show 1

    If Trim(selected_name) = "" Or selected_cas = 0 Then
        Exit Sub
    End If
    frmShredDisp.Show 1
End Sub
