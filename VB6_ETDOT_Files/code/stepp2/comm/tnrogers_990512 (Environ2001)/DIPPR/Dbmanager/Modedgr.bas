Attribute VB_Name = "modeditgroups"
Option Explicit

Public Function load_edit_groups_form() As Boolean

    Dim i As Integer
    Dim label As String
    Dim new_grouptype As String
    ' if global grouptype isn't correct get it from the input list
    
        ' figure out based on input
        For i = 0 To MAX_INPUTS_EACH - 1
            If UCase(Right(Trim(frmeditwizard!lblinputprop(i)), 6)) = "GROUPS" Then
                label = Trim(frmeditwizard!lblinputprop(i))
                new_grouptype = Trim(Left(label, Len(label) - 6))
                Exit For
            End If
        Next i
   
    Call reset_groups(new_grouptype)
    
    ' load the available groups part
    ' DENISE should add something here to grey out the tabs that aren't used
    For i = 0 To MAX_GROUPS - 1
        If Trim(group_smiles(i)) <> "" Then
            frmeditgr!Label1(i).Caption = i + 1 & ".  " & group_smiles(i)
        Else
            frmeditgr!Label1(i).Caption = i + 1 & "."
        End If
    Next i
    ' if a calculation has already been done, put those groups into the selected part
    For i = 0 To MAX_GROUPS_PER_CHEM - 1
        If cur_chem_groups(i) > 0 Then
            frmeditgr!lblselindex(i).Caption = cur_chem_groups(i) & "."
            frmeditgr!lblsel(i).Caption = group_smiles(cur_chem_groups(i) - 1)
            frmeditgr!lblselno(i).Caption = num_cur_chem_groups(i)
        Else
            frmeditgr!lblselindex(i).Caption = ""
            frmeditgr!lblsel(i).Caption = ""
            frmeditgr!lblselno(i).Caption = ""
        End If
    Next i
    load_edit_groups_form = True
    
End Function

Public Sub clear_group_array()

    Dim i As Integer
    For i = 0 To MAX_GROUPS - 1
        group_smiles(i) = ""
    Next i
    
End Sub
