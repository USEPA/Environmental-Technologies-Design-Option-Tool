Attribute VB_Name = "modstruct"
Option Explicit

Public Sub load_form_struct()

    Dim i As Integer
    
    ' the only type of groups that get stored
    ' in the db so far are UNIFAC
    If global_grouptype <> "UNIFAC" Then
        frmstruct!cmdadd.Enabled = False
    Else
        frmstruct!cmdadd.Enabled = True
    End If
    frmstruct!cboschtype.Clear
    frmstruct!cboschtype.AddItem "Sequential, Non-Truncating"
    frmstruct!cboschtype.AddItem "Sequential, Truncating"
    frmstruct!cboschtype.AddItem "Combinatorial, Truncating"
    frmstruct!cboschtype.ListIndex = 2
    
    frmstruct!tbxname.Text = selected_name
    frmstruct!tbxsmiles.Text = selected_smiles
    For i = 0 To 19
        frmstruct!lbldisplay(i).Caption = ""
        frmstruct!lblgrno(i).Caption = ""
    Next i
    frmstruct!lblstatus.Caption = ""
End Sub


Public Sub set_group_array()

    Dim groupset As Recordset
    Dim i As Integer
    Dim group_index As Integer
    On Error Resume Next
    ' now read the group strings from the file
    Select Case Trim(global_grouptype)
        Case "UNIFAC"
           ' Set groupdb = OpenDatabase(dbpath & "\" & dbname, False, False)
            Set groupset = chembrowsedb.OpenRecordset("UNIFAC", dbOpenSnapshot)
            If Not groupset.EOF Then
                groupset.MoveFirst
                While Not groupset.EOF
                    group_index = groupset("Sub Group") - 1
                    group_smiles(group_index) = groupset("Sub Group Structure")
                    groupset.MoveNext
                Wend
            End If
            groupset.Close
        Case "Pintar"
            Set groupset = chembrowsedb.OpenRecordset("Pintar", dbOpenSnapshot)
            If Not groupset.EOF Then
                groupset.MoveFirst
                While Not groupset.EOF
                    group_index = groupset("Group ID") - 1
                    group_smiles(group_index) = groupset("Fragment")
                    groupset.MoveNext
                Wend
            End If
            groupset.Close
       Case "Lydersen"
       MsgBox ("Unable to calculate properties using Lydersen Groups")
       Exit Sub
        Set groupset = chembrowsedb.OpenRecordset("Lydersen", dbOpenSnapshot)
            If Not groupset.EOF Then
                groupset.MoveFirst
                While Not groupset.EOF
                    group_index = groupset("Group ID") - 1
                    group_smiles(group_index) = groupset("Fragment")
                    groupset.MoveNext
                Wend
            End If
            groupset.Close
        Case "Benson"
            Set groupset = chembrowsedb.OpenRecordset("Benson", dbOpenSnapshot)
            If Not groupset.EOF Then
                groupset.MoveFirst
                While Not groupset.EOF
                    group_index = groupset("Group ID") - 1
                    group_smiles(group_index) = groupset("Fragment")
                    groupset.MoveNext
                Wend
            End If
            groupset.Close
        Case "Hine and Mookerjee"
            Set groupset = chembrowsedb.OpenRecordset("Hine & Mookerjee", dbOpenSnapshot)
            If Not groupset.EOF Then
                groupset.MoveFirst
                While Not groupset.EOF
                    group_index = groupset("Group ID") - 1
                    group_smiles(group_index) = groupset("Fragment")
                    groupset.MoveNext
                Wend
            End If
            groupset.Close
    End Select
            
End Sub
