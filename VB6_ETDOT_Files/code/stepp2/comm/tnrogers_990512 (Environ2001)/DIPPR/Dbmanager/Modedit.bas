Attribute VB_Name = "modedit"
Option Explicit

Public Function create_database_copy(name As String) As Boolean
' *paul
'Dim CurrentDatabase As Database
Dim NewDatabase As Database
Dim i As Integer
Dim tempname As String
    On Error GoTo create_db_error
    ' indicate that the def file is modified (to be written when app closed)
    def_modified = True
'    If Not Left(dbpath, 2) Like "*:" Then
'        dbpath = Left(App.path, 2) & dbpath
'    End If
    tempname = AppPath & "~tempdb.mdb"
    frmdbman!frwait.Caption = "Please Wait"
    frmdbman!lblwait.Caption = "building:  " & name & "..."
    frmdbman!frwait.Visible = True
    frmdbman.Refresh
    Screen.MousePointer = 11
    ' create the copy
    DBEngine.CompactDatabase curpath & curname, Trim(tempname), dbLangGeneral
    ' delete all the data in it
    Set NewDatabase = DBEngine.OpenDatabase(Trim(tempname), False, False)
    For i = 0 To NewDatabase.TableDefs.count - 1
        If NewDatabase.TableDefs(i).Attributes = dbSystemObject _
            Or NewDatabase.TableDefs(i).Attributes = -2147483648# _
            Or NewDatabase.TableDefs(i).Attributes = 2 _
            Or NewDatabase.TableDefs(i).Updatable = False _
            Or is_persistant_table(NewDatabase.TableDefs(i).name) = True Then
            GoTo after_delete
        End If
        NewDatabase.Execute "DELETE * FROM [" & NewDatabase.TableDefs(i).name & "]"
after_delete:
    
    Next i
    NewDatabase.Close
    
    DBEngine.CompactDatabase Trim(tempname), Trim(name), dbLangGeneral
    Kill tempname
    Kill AppPath & "~tempdb.ldb"
    
    Screen.MousePointer = 1
    ' set curname to the new database
    create_database_copy = True
Exit Function

create_db_error:
    Screen.MousePointer = 1
    create_database_copy = False

End Function

Public Sub load_add_chem_info()

    Call clear_form_text_boxes
    frmedit!fredit.Visible = False
    frmedit.Height = 6000
    frmedit!frenter.Top = 2040
    frmedit!cmdAccept.Top = 4680
    frmedit!cmdone.Top = 4680
    frmedit!fradd.Visible = True
    frmedit!cmdacceptadd.Default = True
    ' don't let the user enter this until they've added the chem info
    frmedit!cbotable.Enabled = False
    frmedit!cbofield.Enabled = False
    frmedit!cboopt.Enabled = False
    frmedit!tbxdata.Enabled = False
    frmedit!lbltable.Enabled = False
    frmedit!lblfield.Enabled = False
    frmedit!lblopt.Enabled = False
    frmedit!lbldata.Enabled = False
    frmedit!cmdAccept.Enabled = False
    frmedit.Caption = "Database Editor: " & curpath & curname
    If selected_name <> "" Then
        frmedit!lblselname.Caption = selected_name
        frmedit!lblselcas.Caption = selected_cas
        frmedit!lblselname.Visible = True
        frmedit!lblselcas.Visible = True
        frmedit!lblpromptchem.Visible = False
    Else
        frmedit!lblselname.Visible = False
        frmedit!lblselcas.Visible = False
        frmedit!lblpromptchem.Visible = True
        selected_cas = 0
        selected_name = ""
        selected_smiles = ""
        selected_family = ""
    End If
    frmedit!fradd.Visible = True
    frmedit!fredit.Visible = False
    
End Sub

Public Sub load_edit_chem_info()

    Call clear_form_text_boxes
    
    frmedit.Height = 5000
    frmedit!frenter.Top = 1100
    frmedit!cmdAccept.Top = 3680
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
    frmedit!cmdAccept.Enabled = True
    frmedit!fredit.Visible = True
    frmedit!fradd.Visible = False
    frmedit.Caption = "Database Editor: " & curpath & curname
    If Trim(selected_name) <> "" Then
        frmedit!lblselname.Caption = selected_name
        frmedit!lblselcas.Caption = selected_cas
        frmedit!lblselname.Visible = True
        frmedit!lblselcas.Visible = True
        frmedit!lblpromptchem.Visible = False
    Else
        frmedit!lblselname.Caption = ""
        frmedit!lblselcas.Caption = ""
        frmedit!lblselname.Visible = False
        frmedit!lblselcas.Visible = False
        frmedit!lblpromptchem.Visible = True
    End If
    
End Sub


Public Sub load_remove_chem_info()

    clear_form_text_boxes
    frmedit.Height = 5310
    frmedit!frenter.Top = 2040
    frmedit!cmdAccept.Top = 4680
    frmedit!cmdone.Top = 4680
    frmedit.fradd.Visible = False
    frmedit.cbotable.Enabled = True
    frmedit.cbofield.Enabled = True
    frmedit.cboopt.Enabled = True
    frmedit.tbxdata.Enabled = True
    frmedit!lbltable.Enabled = True
    frmedit!lblfield.Enabled = True
    frmedit!lblopt.Enabled = True
    frmedit!lbldata.Enabled = True
    frmedit!cmdAccept.Enabled = True
    frmedit.fredit.Visible = True
    frmedit.Caption = "Database Editor: " & curpath & curname
    If selected_name <> "" Then
        frmedit.lblselname.Caption = selected_name
        frmedit.lblselcas.Caption = selected_cas
        frmedit.lblselname.Visible = True
        frmedit.lblselcas.Visible = True
        frmedit.lblpromptchem.Visible = False
    Else
        frmedit.lblselname.Visible = False
        frmedit.lblselcas.Visible = False
        frmedit.lblpromptchem.Visible = True
    End If
End Sub

Public Sub clear_form_text_boxes()

    
    frmedit!tbxname.Text = ""
    frmedit!tbxcas.Text = ""
    frmedit!tbxsmiles.Text = ""
    frmedit!tbxstructure.Text = ""
    frmedit!tbxfamily.Text = ""
End Sub

Public Sub load_form_edit_info()

    'Dim localdb As Database
    Dim i As Integer
    Dim j As Integer
    If dbstatus = STATUS_CLOSED Then
        Set chembrowsedb = DBEngine.OpenDatabase(Trim(curpath & curname), False, False)
        dbstatus = STATUS_OPEN
    ElseIf dbstatus = STATUS_CHANGED Then
        On Error Resume Next
        chembrowsedb.Close
        Set chembrowsedb = DBEngine.OpenDatabase(Trim(curpath & curname), False, False)
        dbstatus = STATUS_OPEN
    End If
    On Error GoTo next_table
    For i = 0 To chembrowsedb.TableDefs.count - 1
        If Left(Trim(chembrowsedb.TableDefs(i).name), 4) <> "MSys" Then
            frmedit!cbotable.AddItem chembrowsedb.TableDefs(i).name
        End If
next_table:
    Next i
    frmedit!cbotable.ListIndex = 0
    'localdb.Close
    'frmedit!fradd.Visible = False
    'frmedit!fredit.Visible = True
    Call update_field_options
    Call update_editing_options
    Call update_data_box
End Sub

Public Sub update_editing_options()

    ' for now it will just be user input manually
    frmedit!cboopt.Clear
    frmedit!cboopt.AddItem "user input"
    frmedit!cboopt.ListIndex = 0
    
End Sub

Public Sub update_field_options()

    'Dim localdb As Database
    Dim localtable As Recordset
    Dim i As Integer
    
   ' Set localdb = OpenDatabase(dbpath & "\" & curname, False, False)
    Set localtable = chembrowsedb.OpenRecordset(Trim(frmedit!cbotable.Text), dbOpenTable)
    
    frmedit!cbofield.Clear
    For i = 0 To localtable.Fields.count - 1
        ' don't put identifier fields in the box, don't let user edit name or cas
        If localtable.Fields(i).name <> "Name" And localtable.Fields(i).name <> "name" And localtable.Fields(i).name <> "CAS" And localtable.Fields(i).name <> "NCAS" And localtable.Fields(i).name <> "Cas #" And localtable.Fields(i).name <> "Cas#" Then
            frmedit!cbofield.AddItem localtable.Fields(i).name
        End If
    Next i
    If frmedit!cbofield.ListCount > 0 Then
        frmedit!cbofield.ListIndex = 0
    End If
End Sub

Public Sub update_data_box()
    'Dim localdb As Database
    Dim localtable As Recordset
    Dim idfield As String
    Dim localfield As String
    Dim Criteria As String
    Dim i As Integer
    
    If selected_cas = 0 Or Trim(selected_name) = "" Then
        frmedit!tbxdata.Text = ""
        Exit Sub
    End If
    Criteria = "CAS = " & Val(selected_cas)
    localfield = Trim(frmedit!cbofield.Text)
    On Error GoTo no_previous_data
    'Set localdb = OpenDatabase(dbpath & "\" & curname, False, False)
    Set localtable = chembrowsedb.OpenRecordset(Trim(frmedit!cbotable.Text), dbOpenSnapshot)
    
    localtable.FindFirst Criteria
    'localtable.Index = "PrimaryKey"    ' Define current index.
    'localtable.Seek "=", Criteria ' Seek record.
        
    If localtable.NoMatch Then
        frmedit!tbxdata.Text = ""
    Else
        frmedit!tbxdata.Text = localtable(localfield)
    End If
    Exit Sub
no_previous_data:
    frmedit!tbxdata.Text = ""
    
End Sub

Public Function confirm_type(data_type As Integer, arg_type As Integer, entered_string As Variant) As Boolean

    Dim data_value As Variant
    On Error GoTo error_in_type
    Select Case data_type
        Case 8      ' dbDate
            data_value = CDate(entered_string)
            arg_type = dbDate
            confirm_type = True
        Case 10     'dbText
            data_value = CStr(entered_string)
            arg_type = dbText
            confirm_type = True
        Case 12     'dbMemo
            data_value = CStr(entered_string)
            arg_type = dbMemo
            confirm_type = True
        Case 1      'dbBoolean
            data_value = CBool(entered_string)
            arg_type = dbBoolean
            confirm_type = True
        Case 3      'dbInteger
            data_value = CInt(entered_string)
            arg_type = dbInteger
            confirm_type = True
        Case 4      'dbLong
            data_value = CLng(entered_string)
            arg_type = dbLong
            confirm_type = True
        Case 7      'dbDouble
            data_value = CDbl(entered_string)
            arg_type = dbDouble
            confirm_type = True
        Case Else
            confirm_type = False
    End Select
    
    Exit Function
error_in_type:
    confirm_type = False
    Exit Function
End Function

Public Function add_chemical_info() As Boolean

    ' this function adds the chemical to the
    ' database, needs to know which tables to
    ' add it to
    ' master: PEARLS List   (CAS, Name)
    '           Chemical Name   (CAS, Name)
    '           Family (CAS, Chemical Family)
    '           Formula (CAS, Formula)
    '           SMILES Indices (CAS, SMILES)
    '           Source (CAS, Source)
    ' block5:   fexp2   (NCAS, Name)
    
    Dim localtable As Dynaset
    Dim answer As Integer
    Dim modified As Boolean
    Dim newname As String
    Dim newcas As Long
    On Error GoTo error_adding_chemical
    
        modified = False
        newname = selected_name
        newcas = selected_cas
        ' first just check if this is going to conflict with any name/cas in the db
        Set localtable = chembrowsedb.OpenRecordset("PEARLS List", dbOpenDynaset)
        localtable.FindFirst "CAS = " & Val(selected_cas)
        If localtable.NoMatch Then
            newcas = 0
        End If
        localtable.FindFirst "Name = " & "'" & selected_name & "'"
        If localtable.NoMatch Then
            newname = ""
        End If
        localtable.Close
        If newcas <> 0 Then
                answer = MsgBox("CAS already exists in PEARLS List")
                add_chemical_info = False
                Exit Function
        ElseIf newname <> "" Then
            answer = MsgBox("Chemical already exists in PEARLS List. Modify entry?", vbYesNo)
            If answer = vbYes Then
                modified = True
                Call replace_in_master_format(newname)
                add_chemical_info = True
                Exit Function
            End If
        Else
            Call global_add_chem
            add_chemical_info = True
            Exit Function
        End If
                       
    'localdb.Close
    add_chemical_info = True
    Exit Function
    
error_adding_chemical:
    On Error Resume Next
    add_chemical_info = False
End Function

Public Function is_persistant_table(arg_table As String)

    ' this function indicates whether a database table should
    ' be copied with data to a new database during the copy
    ' process
    Dim persistant As Boolean
            Select Case Trim(arg_table)
                Case "Othmer fragments"
                    persistant = True
                Case "reference chemicals"
                    persistant = True
                Case "Schroeder Values"
                    persistant = True
                Case "UNIFAC"
                    persistant = True
                Case "Lydersen"
                    persistant = True
                Case "Benson"
                    persistant = True
                Case "Pintar"
                    persistant = True
                Case "Hine & Mookerjee"
                    persistant = True
                Case "AENV"
                    persistant = True
                Case "AGLB"
                    persistant = True
                Case "ALLE"
                    persistant = True
                Case "AVLE"
                    persistant = True
                Case "Synonym List"
                    persistant = True
                Case "Unit Conversions"
                    persistant = True
                Case Else
                    persistant = False
            End Select
            
        is_persistant_table = persistant
End Function

Public Sub global_add_chem()

    Dim localtable As Dynaset
        ' start editing tables
        Set localtable = chembrowsedb.OpenRecordset("Chemical Name", dbOpenDynaset)
        localtable.FindFirst "Name = " & "'" & selected_name & "'"
        If Not localtable.NoMatch Then
            localtable.Edit
            localtable("CAS") = selected_cas
            localtable("Name") = selected_name
            localtable.Update
        Else
            localtable.AddNew
            localtable("CAS") = selected_cas
            localtable("Name") = selected_name
            localtable.Update
        End If
        localtable.Close
        Set localtable = chembrowsedb.OpenRecordset("PEARLS List", dbOpenDynaset)
        localtable.FindFirst "Name = " & "'" & selected_name & "'"
        If Not localtable.NoMatch Then
            localtable.Edit
            localtable("CAS") = selected_cas
            localtable("Name") = selected_name
            localtable("SMILES") = selected_smiles
            localtable("Formula") = selected_structure
            localtable("Chemical Family") = selected_family
            localtable("Source") = "user input from PPMS DBM"
            localtable.Update
        Else
            localtable.AddNew
            localtable("CAS") = selected_cas
            localtable("Name") = selected_name
            localtable("SMILES") = selected_smiles
            localtable("Formula") = selected_structure
            localtable("Chemical Family") = selected_family
            localtable("Source") = "user input from PPMS DBM"
            localtable.Update
        End If
        localtable.Close
        Set localtable = chembrowsedb.OpenRecordset("Family", dbOpenDynaset)
        localtable.FindFirst "CAS = " & Val(selected_cas)
        If Not localtable.NoMatch Then
            localtable.Edit
            localtable("CAS") = selected_cas
            localtable("Chemical Family") = selected_family
            localtable.Update
        Else
            localtable.AddNew
            localtable("CAS") = selected_cas
            localtable("Chemical Family") = selected_family
            localtable.Update
        End If
        localtable.Close
        Set localtable = chembrowsedb.OpenRecordset("Formula", dbOpenDynaset)
        localtable.FindFirst "CAS = " & Val(selected_cas)
        If Not localtable.NoMatch Then
            localtable.Edit
            localtable("CAS") = selected_cas
            localtable("Formula") = selected_structure
            localtable.Update
        Else
            localtable.AddNew
            localtable("CAS") = selected_cas
            localtable("Formula") = selected_structure
            localtable.Update
        End If
        localtable.Close
        Set localtable = chembrowsedb.OpenRecordset("SMILES Indices", dbOpenDynaset)
        localtable.FindFirst "CAS = " & Val(selected_cas)
        If Not localtable.NoMatch Then
            localtable.Edit
            localtable("CAS") = selected_cas
            localtable("SMILES") = selected_smiles
            localtable.Update
        Else
            localtable.AddNew
            localtable("CAS") = selected_cas
            localtable("SMILES") = selected_smiles
            localtable.Update
        End If
        localtable.Close
        Set localtable = chembrowsedb.OpenRecordset("Source", dbOpenDynaset)
        localtable.FindFirst "CAS = " & Val(selected_cas)
        If Not localtable.NoMatch Then
            localtable.Edit
            localtable("CAS") = selected_cas
            localtable("Source") = "user input from PPMS DBM"
            localtable.Update
        Else
            localtable.AddNew
            localtable("CAS") = selected_cas
            localtable("Source") = "user input from PPMS DBM"
            localtable.Update
        End If
        localtable.Close

End Sub

Public Sub replace_in_master_format(newname As String)

    ' a function to globally replace a name in the master format database
    Dim localtable As Dynaset
    If newname <> "" Then
        Set localtable = chembrowsedb.OpenRecordset("PEARLS List", dbOpenDynaset)
        localtable.FindFirst "CAS = " & Val(selected_cas)
        If Not localtable.NoMatch Then
            localtable.Edit
            localtable("Name") = newname
            localtable.Update
        End If
        localtable.Close
        Set localtable = chembrowsedb.OpenRecordset("Chemical Name", dbOpenDynaset)
        localtable.FindFirst "CAS = " & Val(selected_cas)
        If Not localtable.NoMatch Then
            localtable.Edit
            localtable("Name") = newname
            localtable.Update
        End If
        localtable.Close
   End If
        
        
End Sub
