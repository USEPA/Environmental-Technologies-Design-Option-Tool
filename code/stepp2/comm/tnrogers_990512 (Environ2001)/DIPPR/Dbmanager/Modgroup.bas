Attribute VB_Name = "modgroups"
Option Explicit

Public Sub reset_groups(new_grouptype As String)

Dim i As Integer
Dim j As Integer
    ' first, if it's not a group change, or there is no current group, don't do anything
    If UCase(Trim(global_grouptype)) = UCase(Trim(new_grouptype)) Then
        Exit Sub
    End If
    If Len(Trim(new_grouptype)) < 4 Then
        Exit Sub
    End If
    'clear the groups for the current chemical
    For i = 0 To MAX_GROUPS_PER_CHEM - 1
        cur_chem_groups(i) = -1
        num_cur_chem_groups(i) = 0
    Next i
    'clear the array holding current group type groups
    If Trim(group_smiles(0)) <> "" Then
        Call clear_group_array
    End If
    global_grouptype = new_grouptype
    
    Call set_group_array
    ' give the form the appropriate caption
    frmeditgr.Caption = "Editing " & global_grouptype & " groups"
    ' make sure the frgroups has the same one
    frmeditwizard!frgroups.Caption = global_grouptype & " Groups"
    
    ' update the actual grid (clear it)
    For j = 1 To frmeditwizard!grdgroups.Rows
        frmeditwizard!grdgroups.Row = j
        frmeditwizard!grdgroups.Col = 1
        If frmeditwizard!grdgroups.Text <> "" Then
            frmeditwizard!grdgroups.Text = ""
            frmeditwizard!grdgroups.Col = 2
            frmeditwizard!grdgroups.Text = ""
            frmeditwizard!grdgroups.Col = 3
            frmeditwizard!grdgroups.Text = ""
        Else
                Exit For
        End If
    Next j
End Sub

Public Sub update_groups()

Dim i As Integer
Dim j As Integer
    ' fill in the grid containing the current chemical groups
    i = 0
    frmeditwizard!grdgroups.ColWidth(0) = 300
    frmeditwizard!grdgroups.ColWidth(1) = 800
    frmeditwizard!grdgroups.ColWidth(2) = 3100
    frmeditwizard!grdgroups.ColWidth(3) = 800
    
    While cur_chem_groups(i) <> -1 And i < MAX_GROUPS_PER_CHEM - 1
        frmeditwizard!grdgroups.Row = i + 1
        frmeditwizard.grdgroups.Col = 1
        frmeditwizard!grdgroups.Text = cur_chem_groups(i)
        frmeditwizard!grdgroups.Col = 2
        frmeditwizard!grdgroups.Text = group_smiles(cur_chem_groups(i))
        frmeditwizard!grdgroups.Col = 3
        frmeditwizard!grdgroups.Text = num_cur_chem_groups(i)
        i = i + 1
    Wend
    ' erase the rest of the rows
    For j = i + 1 To MAX_GROUPS_PER_CHEM - 1
        frmeditwizard!grdgroups.Row = j
        frmeditwizard!grdgroups.Col = 1
        frmeditwizard!grdgroups.Text = ""
        frmeditwizard!grdgroups.Col = 2
        frmeditwizard!grdgroups.Text = ""
        frmeditwizard!grdgroups.Col = 3
        frmeditwizard!grdgroups.Text = ""
    Next j
    frmeditwizard!grdgroups.Refresh
End Sub
