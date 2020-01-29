Attribute VB_Name = "modchembrowse"
Option Explicit

Public Sub load_chem_browse_info()

    import_flag = False
    If dbstatus = STATUS_CLOSED Then
        Set chembrowsedb = OpenDatabase(curpath & curname, False, False)
        dbstatus = STATUS_OPEN
    ElseIf dbstatus = STATUS_CHANGED Then
        On Error Resume Next
        chembrowsedb.Close
        Set chembrowsedb = OpenDatabase(curpath & curname, False, False)
        dbstatus = STATUS_OPEN
    End If
    
    ' for now just hard code the path in
    ' open the database used for searching, this should be in main
    frmchembrowse!grdchemlist(0).Columns(0).Width = 1200
    frmchembrowse!grdchemlist(0).Columns(1).Width = 4400
    frmchembrowse!Data1(MASTER).DatabaseName = curpath & curname
    frmchembrowse!Data1(MASTER).RecordSource = "PEARLS List"
    frmchembrowse!Data1(MASTER).RecordsetType = 2
    frmchembrowse!Data1(MASTER).Refresh
    frmchembrowse!Data1(MASTER).Recordset.MoveLast
    frmchembrowse!Data1(MASTER).Recordset.MoveFirst
    
    frmchembrowse!grdchemlist(0).Visible = True
    frmchembrowse!grdchemlist(0).Refresh
     ' set the find stuff
    frmchembrowse!cbofind.Clear
    frmchembrowse!cbofind.AddItem "CAS"
    frmchembrowse!cbofind.AddItem "Name"
    frmchembrowse!cbofind.AddItem "Formula"
    frmchembrowse!cbofind.AddItem "Source"
    frmchembrowse!cbofind.AddItem "Chemical Family"
    frmchembrowse!cbofind.AddItem "Smiles"
    frmchembrowse!cbofind.ListIndex = 0
    ' set the filter stuff
    frmchembrowse!cbofilterfield.Clear
    frmchembrowse!cbofilterfield.AddItem "Chemical Family"
    frmchembrowse!cbofilterfield.AddItem "Source"
    frmchembrowse!cbofilterfield.AddItem "user input (Name)"
    frmchembrowse!cbofilterfield.ListIndex = 0
    ' set the sorting stuff
    frmchembrowse!cbosort.Clear
    frmchembrowse!cbosort.AddItem "CAS"
    frmchembrowse!cbosort.AddItem "Name"
    frmchembrowse!cbosort.ListIndex = 0
    frmchembrowse.Refresh
End Sub

Public Sub load_chem_browse_special(filearg As String)

    ' this is basically the same as the regular
    ' load chem browse except:
    '   1.  it takes a database name as an argument
    '       and opens that one instead of the global database

import_flag = True
If dbstatus = STATUS_CLOSED Then
        Set chembrowsedb = OpenDatabase(filearg, False, False)
        dbstatus = STATUS_OPEN
    ElseIf dbstatus = STATUS_CHANGED Then
        On Error Resume Next
        chembrowsedb.Close
        Set chembrowsedb = OpenDatabase(filearg, False, False)
        dbstatus = STATUS_OPEN
    End If
    
    ' for now just hard code the path in
    ' open the database used for searching, this should be in main
    frmchembrowse!grdchemlist(0).Columns(0).Width = 1000
    frmchembrowse!grdchemlist(0).Columns(1).Width = 4300
    frmchembrowse!Data1(MASTER).DatabaseName = curpath & curname
    frmchembrowse!Data1(MASTER).RecordSource = "PEARLS List"
    frmchembrowse!Data1(MASTER).RecordsetType = 2
    frmchembrowse!grdchemlist(0).Visible = True
    frmchembrowse!Data1(MASTER).Refresh
    frmchembrowse!grdchemlist(0).Refresh
     ' set the find stuff
    frmchembrowse!cbofind.Clear
    frmchembrowse!cbofind.AddItem "CAS"
    frmchembrowse!cbofind.AddItem "Name"
    frmchembrowse!cbofind.AddItem "Formula"
    frmchembrowse!cbofind.AddItem "Source"
    frmchembrowse!cbofind.AddItem "Chemical Family"
    frmchembrowse!cbofind.AddItem "Smiles"
    frmchembrowse!cbofind.ListIndex = 0
    ' set the filter stuff
    frmchembrowse!cbofilterfield.Clear
    frmchembrowse!cbofilterfield.AddItem "Chemical Family"
    frmchembrowse!cbofilterfield.AddItem "Source"
    frmchembrowse!cbofilterfield.AddItem "user input (Name)"
    frmchembrowse!cbofilterfield.ListIndex = 0
    frmchembrowse!cbofiltercat.Clear
    frmchembrowse!cbofiltercat.Text = ""
    ' set the sorting stuff
    frmchembrowse!cbosort.Clear
    frmchembrowse!cbosort.AddItem "CAS"
    frmchembrowse!cbosort.AddItem "Name"
    frmchembrowse!cbosort.ListIndex = 0
    frmchembrowse.Refresh
End Sub

