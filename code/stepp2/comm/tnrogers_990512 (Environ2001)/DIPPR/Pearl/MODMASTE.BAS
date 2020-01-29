Attribute VB_Name = "modmaster"

Public Sub load_frm_master_info()

    If PathMaster <> NULLPATH Then
        frmmaster!txtpath(0).caption = PathMaster
        frmmaster!txtpath(0).Visible = True
    Else
        frmmaster!txtpath(0).caption = "null"
        frmmaster!txtpath(0).Visible = True
    End If
    If PathSave <> NULLPATH Then
        frmmaster!txtpath(1).caption = PathSave
        frmmaster!txtpath(1).Visible = True
    Else
        frmmaster!txtpath(1).caption = "null"
        frmmaster!txtpath(1).Visible = True
    End If
    
    If PathBlock5 <> NULLPATH Then
        frmmaster!txtpath(2).caption = PathBlock5
        frmmaster!txtpath(2).Visible = True
    Else
        frmmaster!txtpath(2).caption = "null"
        frmmaster!txtpath(2).Visible = True
    End If
    ' now set the user list selected as default
    frmmaster!txtname(1).Font.Bold = True
    frmmaster!txtpath(1).Font.Bold = True
End Sub

Public Function load_selected_save() As Boolean
     ' this is the mnuload_click function without the file selection
    ' used when file has already been selected (as in after browser)
    Dim DBTbl As Recordset
    Dim Response As Integer
    Dim CLCAS As Long
    Dim ULCAS As Long
    Dim LastList As Integer
    
    'Check for existing selections
    'On Error GoTo skip_save_warning
    If FRMMain!Data2.Recordset.RecordCount >= 1 And WorkModified = True Then
        Response = MsgBox("Save current template?", 3, "Save")
        If Response = 6 Then
            Call FRMMain.MNUSave_Click
        ElseIf Response = 2 Then
            Exit Function
        End If
    End If
   ' DBJetUser.Close
skip_save_warning:
    On Error GoTo LoadError
        
    'Set mousepointer to hourglass (wait mode)
    Screen.MousePointer = 11
    
    'Refresh main form
    FRMMain.Refresh
    On Error Resume Next
    DBJetUser.Close
    On Error GoTo LoadError
    ' temporarily attack the temp db to the control
    FRMMain!Data2.databasename = App.path & "\temp.mdb"
    FRMMain!Data2.RecordSource = "User List"
    FRMMain!Data2.Refresh
    On Error Resume Next
    
    Kill PathUser
    On Error GoTo LoadError
    ' set the savefile name and copy it to dbuser so we're not directly editing the saved file
    SaveFileName = PathSave
    FileCopy PathSave, PathUser
    Set DBJetUser = OpenDatabase(PathUser, False, False)
    ' now reset the control to our user file
    UserDBName = PathUser
    FRMMain!Data2.databasename = PathUser
    FRMMain!Data2.RecordSource = "User List"
    FRMMain!Data2.RecordsetType = 2
    FRMMain!Data2.Refresh
    On Error Resume Next
    Set DBTbl = DBJetUser.OpenRecordset("Last CAS", dbOpenTable)
        
    Cur_Info.CAS = DBTbl("LastCAS")
    CLCAS = DBTbl("CLCAS")
    ULCAS = DBTbl("ULCAS")
    LastList = DBTbl("LastList")
    
    DBTbl.Close
    
    FRMMain!Data1.Recordset.FindFirst "CAS =" & Cur_Info.CAS
    ' try moving this here
    Call LoadUserPreferences
    If GetUserData = True Then
        If LastList = 1 And CLCAS = Cur_Info.CAS Then
            'Call LoadUserPreferences
            Call Recalculate
            Call DisplayProps
            Call TabFolderEnable(True)
        End If
        If LastList = 2 And ULCAS = Cur_Info.CAS Then
            'Call LoadUserPreferences
            Call Recalculate
            Call DisplayProps
            Call TabFolderEnable(True)
        End If
    End If
        
    FRMMain!Data1.Recordset.FindFirst "CAS =" & CLCAS
    FRMMain!LSTSelList.Text = FRMMain!Data1.Recordset("Name")
    
    FRMMain!Data2.Recordset.FindFirst "CAS =" & ULCAS
    FRMMain!LSTUserList.Text = FRMMain!Data2.Recordset("Name")
        
    If LastList = 1 Then
        FRMMain!LSTSelList.SetFocus
    Else
        FRMMain!LSTUserList.SetFocus
    End If
        
    'Set mousepointer to arrow (normal mode)
    Screen.MousePointer = 1
    'Reset modified flag
    WorkModified = False
    ' make sure the file name fits on the main form
    If Len(SaveFileName) > 45 Then
        FRMMain.caption = "PEARLS:  ..." & Right(SaveFileName, 45)
    Else
        FRMMain.caption = "PEARLS:  " & SaveFileName
    End If
    load_selected_save = True
    Exit Function
       
LoadError:
    FRMMain!Data2.databasename = App.path & "\temp.mdb"
    FRMMain!Data2.RecordSource = "User List"
    FRMMain!Data2.Refresh
    Screen.MousePointer = 1
    If Err <> 32755 Then
        MsgBox "Error loading PEARLS file", 48, "Error"
    End If
    load_selected_save = False
End Function
