Attribute VB_Name = "modeditwiz"
Option Explicit

Public Function check_inputs(property_num As Integer, method_num As Integer) As Boolean

    ' this function needs to check whether we have the necessary
    ' inputs to calculate this method for this property and, if
    ' so, return true, else, return false
    ' this function also sets the selected_temperature and selected_rings for the current chemical in case
    ' they haven't already been set
    Dim I As Integer
    Dim J As Integer
    Dim success As Boolean
    Dim error_message As String
    error_message = "Inputs not found: "
    success = True
    For I = 0 To MAX_INPUTS_EACH - 1
        If wiz_inputs(property_num, method_num, I) = -1 Then
            success = True
            Exit For
        End If
        Select Case Trim(wiz_inputs(property_num, method_num, I))
            Case CONST_U_GROUPS
                If Trim(frmeditwizard!frgroups.Caption) = input_name(CONST_U_GROUPS) And cur_chem_groups(0) <> 0 Then
                    GoTo next_iteration
                Else
                    error_message = error_message & Chr(13) & Chr(9) & input_name(CONST_U_GROUPS)
                    success = False
                    Exit For
                End If
            Case CONST_P_GROUPS
                If Trim(frmeditwizard!frgroups.Caption) = input_name(CONST_P_GROUPS) And cur_chem_groups(0) <> 0 Then
                    GoTo next_iteration
                Else
                    error_message = error_message & Chr(13) & Chr(9) & input_name(CONST_P_GROUPS)
                    success = False
                End If
            Case CONST_B_GROUPS
                If Trim(frmeditwizard!frgroups.Caption) = input_name(CONST_B_GROUPS) And cur_chem_groups(0) <> 0 Then
                    GoTo next_iteration
                Else
                    error_message = error_message & Chr(13) & Chr(9) & input_name(CONST_B_GROUPS)
                    success = False
                End If
            Case CONST_HM_GROUPS
                If Trim(frmeditwizard!frgroups.Caption) = input_name(CONST_HM_GROUPS) And cur_chem_groups(0) <> 0 Then
                    GoTo next_iteration
                Else
                    error_message = error_message & Chr(13) & Chr(9) & input_name(CONST_HM_GROUPS)
                    success = False
                End If
            Case CONST_L_GROUPS
                If Trim(frmeditwizard!frgroups.Caption) = input_name(CONST_L_GROUPS) And cur_chem_groups(0) <> 0 Then
                    GoTo next_iteration
                Else
                    error_message = error_message & Chr(13) & Chr(9) & input_name(CONST_L_GROUPS)
                    success = False
                End If
            Case CONST_ELEMENTS
                frmeditwizard!grdelements.Row = 1
                frmeditwizard!grdelements.Col = 1
                If Trim(frmeditwizard!grdelements.Text) <> "" Then
                    GoTo next_iteration
                Else
                    error_message = error_message & Chr(13) & Chr(9) & input_name(CONST_ELEMENTS)
                    success = False
                End If
            Case CONST_TEMP
                For J = 0 To MAX_INPUTS_EACH - 1
                    If Trim(frmeditwizard!lblinputprop(J).Caption) = input_name(CONST_TEMP) Then
                        If Len(Trim(frmeditwizard!tbxinputprop(J).Text)) > 0 And IsNumeric(frmeditwizard!tbxinputprop(J).Text) Then
                            selected_temperature = CDbl(frmeditwizard!tbxinputprop(J).Text)
                            GoTo next_iteration
                        Else
                            selected_temperature = STANDARD_TEMPERATURE
                            frmeditwizard!tbxinputprop(J).Text = selected_temperature
                            error_message = error_message & Chr(13) & Chr(9) & input_name(CONST_TEMP) & "(using " & STANDARD_TEMPERATURE & ")"
                            success = False
                        End If
                    End If
                Next J
            Case CONST_NUM_RINGS
                For J = 0 To MAX_INPUTS_EACH - 1
                    If Trim(frmeditwizard!lblinputprop(J).Caption) = input_name(CONST_NUM_RINGS) Then
                        If Len(Trim(frmeditwizard!tbxinputprop(J).Text)) > 0 And IsNumeric(frmeditwizard!tbxinputprop(J).Text) Then
                            selected_rings = CInt(frmeditwizard!tbxinputprop(J).Text)
                            GoTo next_iteration
                        Else
                            error_message = error_message & Chr(13) & Chr(9) & input_name(CONST_NUM_RINGS)
                            success = False
                        End If
                    End If
                Next J
            Case Else
                error_message = error_message & Chr(13) & Chr(9) & input_name(wiz_inputs(property_num, method_num, I))
                success = False
        End Select
next_iteration:
    
    Next I
after_iteration:
        If success = True Then
            check_inputs = True
        Else
            MsgBox (error_message)
            check_inputs = False
        End If
        
    

End Function

Public Sub load_edit_wizard_form()

    Dim I As Integer
    Dim J As Integer
    Dim K As Integer
    Dim found As Integer
    Dim method_selected As Integer
    ' DENISE for now hard code this in
    global_cur_property = -1
    
    ' set up the element grid
    frmeditwizard!grdelements.ColWidth(0) = 200
    frmeditwizard!grdelements.ColWidth(1) = 725
    frmeditwizard!grdelements.ColWidth(2) = 500
    frmeditwizard!grdelements.Row = 0
    frmeditwizard!grdelements.Col = 1
    frmeditwizard!grdelements.Text = "Element"
    frmeditwizard!grdelements.Col = 2
    frmeditwizard!grdelements.Text = "#"
    ' the element frame info
    For I = 1 To MAX_ELEMENTS - 1
        frmeditwizard!grdelements.Row = I
        frmeditwizard!grdelements.Col = 1
        frmeditwizard!grdelements.Text = ""
        frmeditwizard!grdelements.Col = 2
        frmeditwizard!grdelements.Text = ""
    Next I
    For I = 1 To MAX_GROUPS_PER_CHEM - 1
        frmeditwizard!grdgroups.Row = I
        frmeditwizard!grdgroups.Col = 1
        frmeditwizard!grdgroups.Text = ""
        frmeditwizard!grdgroups.Col = 2
        frmeditwizard!grdgroups.Text = ""
    Next I
    Call load_frmselprop_info
    frmselprop.Show 1
    'Call reset_groups
    ' size the grid containing the groups
    frmeditwizard!grdgroups.ColWidth(0) = 250
    frmeditwizard!grdgroups.ColWidth(1) = 800
    frmeditwizard!grdgroups.ColWidth(2) = 3000
    frmeditwizard!grdgroups.ColWidth(3) = 750
    
    'frmeditwizard!lblgrprompt.Visible = True
    frmeditwizard!grdgroups.Row = 0
    frmeditwizard!grdgroups.Col = 1
    frmeditwizard!grdgroups.Text = "Grp ID"
    frmeditwizard!grdgroups.Col = 2
    frmeditwizard!grdgroups.Text = "Group SMILES"
    frmeditwizard!grdgroups.Col = 3
    frmeditwizard!grdgroups.Text = "#"
    frmeditwizard!grdgroups.Refresh
    ' select a current method (first by default)
    
    
End Sub


Public Sub init_wizard_globals()

' this function sets all the information for:
'   -> names of properties for display
'   -> methods available for each property
'   -> inputs needed for each method (by number)
' NOTE: may eventually want to store this in a file and
'       read it in, for now we'll hard code it in
Dim I As Integer
Dim J As Integer
Dim K As Integer
' the path to the dll
'dllpath = "c:\windows\system"
For I = 0 To MAX_PROPERTIES - 1
    input_name(I) = ""
    input_enabled(I) = False
    For J = 0 To MAX_METHODS_EACH - 1
        wiz_methods(I, J) = ""
        For K = 0 To MAX_INPUTS_EACH - 1
            wiz_inputs(I, J, K) = -1
        Next K
    Next J
Next I

' set the input properties
input_name(FP) = "Flashpoint"
input_name(LFL) = "Lower Flammability Limit"
input_name(UFL) = "Upper Flammability Limit"
input_name(AIT) = "Autoignition Temperature"
input_name(VP) = "Vapor Pressure as f(t)"
input_name(LD) = "Liquid Density as f(t)"
input_name(MW) = "Molecular Weight"
'input_name(CV) = "Critical Volume"
input_name(HC) = "Henry's Constant"
input_name(Schem) = "Solubility Chemical in Water"
input_name(Swater) = "Solubility Water in Chemical"
input_name(ACchem) = "Activity Coefficient of Chemical in Water"
input_name(ACwater) = "Activity Coefficient of Water in Chemical"
input_name(logKow) = "log Kow"
input_name(logKoc) = "log Koc"
input_name(BCF) = "Bioconcentration Factor"
input_name(ThODcarb) = "Carbonaceous ThOD"
input_name(ThODcomb) = "Combined ThOD"
input_name(CONST_REF_CHEM) = "Reference Chemical Information"
input_name(CONST_ELEMENTS) = "elements"
input_name(CONST_U_GROUPS) = "UNIFAC Groups"
input_name(CONST_P_GROUPS) = "Pintar Groups"
input_name(CONST_B_GROUPS) = "Benson Groups"
input_name(CONST_L_GROUPS) = "Lydersen Groups"
input_name(CONST_HM_GROUPS) = "Hine and Mookerjee Groups"
input_name(CONST_NUM_RINGS) = "Number of Rings"
input_name(CONST_TEMP) = "Temperature (C)"

' etc

' set the methods
wiz_methods(FP, 0) = "MTU LFL Group Contribution"
wiz_methods(FP, 1) = "Penn State Group Contribution"
wiz_methods(LFL, 0) = "Penn State Group Contribution"
wiz_methods(LFL, 1) = "MTU Combustion Reaction"
wiz_methods(LFL, 2) = "MTU Group Contribution"
wiz_methods(UFL, 0) = "Penn State Group Contribution"
wiz_methods(UFL, 1) = "MTU Combustion Reaction"
wiz_methods(UFL, 2) = "MTU Group Contribution"
wiz_methods(AIT, 0) = "MTU Linear Method"
wiz_methods(AIT, 1) = "MTU Logarithmic Method"
wiz_methods(LD, 0) = "Rogers Method"
'wiz_methods(CV, 0) = "Lydersen Method"
wiz_methods(VP, 0) = "Loll Method"
wiz_methods(MW, 0) = "atomic"
wiz_methods(HC, 0) = "Hine and Mookerjee"
wiz_methods(Schem, 0) = "UNIFAC"
wiz_methods(Swater, 0) = "UNIFAC"
wiz_methods(ACchem, 0) = "UNIFAC"
wiz_methods(ACwater, 0) = "UNIFAC"
wiz_methods(logKow, 0) = "UNIFAC"
wiz_methods(logKoc, 0) = "Baker"
wiz_methods(ThODcarb, 0) = "Baker"
wiz_methods(ThODcomb, 0) = "Baker"
wiz_methods(BCF, 0) = "Kobayshi"
' etc

' set the inputs
wiz_inputs(FP, 0, 0) = CONST_P_GROUPS       ' MTU LFL Gr Cont
wiz_inputs(FP, 0, 1) = CONST_ELEMENTS       ' MTU LFL Gr Cont
wiz_inputs(FP, 1, 0) = CONST_P_GROUPS       ' penn state gr cont
wiz_inputs(FP, 1, 1) = CONST_ELEMENTS       ' penn state gr cont
wiz_inputs(LFL, 0, 0) = CONST_P_GROUPS      ' Penn State
wiz_inputs(LFL, 0, 1) = CONST_ELEMENTS      ' Penn State
wiz_inputs(LFL, 1, 0) = CONST_P_GROUPS      ' combustion reaction method
wiz_inputs(LFL, 1, 1) = CONST_ELEMENTS
wiz_inputs(LFL, 2, 0) = CONST_P_GROUPS      ' mtu group cont
wiz_inputs(LFL, 2, 1) = CONST_ELEMENTS
wiz_inputs(UFL, 0, 0) = CONST_P_GROUPS      ' Penn State
wiz_inputs(UFL, 0, 1) = CONST_ELEMENTS      ' Penn State
wiz_inputs(UFL, 1, 0) = CONST_P_GROUPS      ' comb reaction
wiz_inputs(UFL, 1, 1) = CONST_ELEMENTS
wiz_inputs(UFL, 2, 0) = CONST_P_GROUPS      ' mtu group cont
wiz_inputs(UFL, 2, 1) = CONST_ELEMENTS
wiz_inputs(AIT, 0, 0) = CONST_P_GROUPS      ' MTU linear
wiz_inputs(AIT, 0, 1) = CONST_ELEMENTS      ' MTU linear
wiz_inputs(AIT, 1, 0) = CONST_P_GROUPS      ' mtu logarithmic
wiz_inputs(AIT, 1, 1) = CONST_ELEMENTS
wiz_inputs(VP, 0, 0) = CONST_U_GROUPS       ' loll method
wiz_inputs(VP, 0, 1) = CONST_TEMP           ' loll method

wiz_inputs(LD, 0, 0) = CONST_ELEMENTS       ' rogers method
wiz_inputs(LD, 0, 1) = CONST_U_GROUPS       ' rogers method
wiz_inputs(LD, 0, 2) = CONST_NUM_RINGS      ' rogers method
wiz_inputs(LD, 0, 3) = CONST_TEMP           ' rogers method
'wiz_inputs(CV, 0, 0) = CONST_L_GROUPS       ' lydersen method
'wiz_inputs(CV, 0, 1) = CONST_TEMP
wiz_inputs(MW, 0, 0) = CONST_ELEMENTS       ' molecular weight by atomic
wiz_inputs(HC, 0, 0) = CONST_HM_GROUPS      ' Henry's constant by Hine & Mookerjee
wiz_inputs(Schem, 0, 0) = CONST_U_GROUPS    ' Solubility Chemical in Water by UNIFAC
wiz_inputs(Schem, 0, 1) = CONST_TEMP
wiz_inputs(Swater, 0, 0) = CONST_U_GROUPS   ' Solubility Water in Chemical by UNIFAC
wiz_inputs(Swater, 0, 1) = CONST_TEMP
wiz_inputs(ACchem, 0, 0) = CONST_U_GROUPS   ' Activity Coefficient chemical in water by UNIFAC method
wiz_inputs(ACchem, 0, 1) = CONST_TEMP
wiz_inputs(ACwater, 0, 0) = CONST_U_GROUPS   ' Activity Coefficient water in chemical by UNIFAC method
wiz_inputs(ACwater, 0, 1) = CONST_TEMP
wiz_inputs(logKow, 0, 0) = CONST_U_GROUPS
wiz_inputs(logKow, 0, 1) = CONST_TEMP
wiz_inputs(logKoc, 0, 0) = CONST_TEMP
wiz_inputs(ThODcarb, 0, 0) = CONST_ELEMENTS ' Baker method for ThODcarb
wiz_inputs(ThODcomb, 0, 0) = CONST_ELEMENTS ' Baker method for ThODcomb
wiz_inputs(BCF, 0, 0) = logKow              ' Kobayshi method for BCF


End Sub


Public Sub load_frmselprop_info()

    Dim I As Integer
    frmselprop!cboprop.Clear
    For I = 0 To MAX_DISPLAY_PROPERTIES     ' actual PPMS properties start at 10
        If Trim(input_name(I)) <> "" Then
            frmselprop!cboprop.AddItem input_name(I)
        End If
    Next I
    frmselprop!cboprop.ListIndex = 0
    
End Sub



Public Sub update_edwiz_input_info(methodindex As Integer)
    ' called when the user selects a method on the wizard form
    ' or when the property is changed
    Dim I As Integer
    Dim J As Integer
    Dim calclistposition As Integer
    Dim label As String
    Dim input_groups As String
    Dim local_input_name As String
    input_groups = ""
    local_input_name = ""
    
    ' first blank out the input stuff
    For I = 0 To MAX_INPUTS_EACH - 1
        frmeditwizard!lblinputprop(I) = ""
        frmeditwizard!tbxinputprop(I).Visible = False
    Next I
    For I = 0 To 2
        frmeditwizard!lblinputcalc(I) = ""
    Next I
    ' if we already have groups for this chemical put them in there
    ' this assumes that if the groups in arrayquant weren't valid it would
    ' have been cleared
    
        For I = 1 To frmeditwizard!grdgroups.Rows
            If num_cur_chem_groups(I) > 0 Then
                frmeditwizard!grdgroups.Row = I
                frmeditwizard!grdgroups.Col = 1
                frmeditwizard!grdgroups.Text = cur_chem_groups(I)
                frmeditwizard!grdgroups.Col = 2
                frmeditwizard!grdgroups.Text = group_smiles(cur_chem_groups(I))
                frmeditwizard!grdgroups.Col = 3
                frmeditwizard!grdgroups.Text = num_cur_chem_groups(I)
            Else
                Exit For
            End If
        Next I
   
            
    calclistposition = 0    ' max must be 2
    ' now put in the inputs for this property/this method
    For I = 0 To MAX_INPUTS_EACH - 1
        If wiz_inputs(global_cur_property, methodindex, I) <> -1 Then
            ' if it's a calculable property put it in that list
            ' first get the input name
            local_input_name = Trim(input_name(wiz_inputs(global_cur_property, methodindex, I)))
            If wiz_inputs(global_cur_property, methodindex, I) < 41 And wiz_inputs(global_cur_property, methodindex, I) <> NBP Then
                frmeditwizard!lblinputcalc(calclistposition) = local_input_name
                calclistposition = calclistposition + 1
                GoTo next_i
            End If
            frmeditwizard!lblinputprop(I) = local_input_name
            If wiz_inputs(global_cur_property, methodindex, I) <= CONST_ELEMENTS And wiz_inputs(global_cur_property, methodindex, I) >= CONST_U_GROUPS Then
                frmeditwizard!tbxinputprop(I).Text = ""
                frmeditwizard!tbxinputprop(I).Visible = False
                frmeditwizard!lblinputprop(I).Visible = True
                frmeditwizard!lblinputprop(I).Caption = local_input_name
                ' update the global grouptype if necessary
                If Right(input_name(wiz_inputs(global_cur_property, methodindex, I)), 6) = "Groups" Then
                    input_groups = Trim(Left(local_input_name, Len(local_input_name) - 6))
                    
                End If
            ElseIf wiz_inputs(global_cur_property, methodindex, I) = CONST_NUM_RINGS Then
                frmeditwizard!lblinputprop(I).Visible = True
                frmeditwizard!lblinputprop(I).Caption = local_input_name
                frmeditwizard!tbxinputprop(I).Visible = True
                frmeditwizard!tbxinputprop(I).Text = get_num_rings(selected_cas)
            ElseIf wiz_inputs(global_cur_property, methodindex, I) = CONST_TEMP Then
                frmeditwizard!lblinputprop(I).Visible = True
                frmeditwizard!lblinputprop(I).Caption = local_input_name
                frmeditwizard!tbxinputprop(I).Visible = True
                frmeditwizard!tbxinputprop(I).Text = selected_temperature
            Else
                frmeditwizard!lblinputprop(I).Visible = True
                frmeditwizard!lblinputprop(I).Caption = local_input_name
                frmeditwizard!tbxinputprop(I).Visible = True
                frmeditwizard!tbxinputprop(I).Text = ""
            End If
        Else
            frmeditwizard!lblinputprop(I) = ""
            frmeditwizard!lblinputprop(I).Visible = False
            frmeditwizard!tbxinputprop(I).Text = ""
            frmeditwizard!tbxinputprop(I).Visible = False
        End If
next_i:
        
    Next I

    If input_groups = "" Then
        frmeditwizard!frgroups.Caption = "Groups"
        global_grouptype = ""
    End If
    Call reset_groups(input_groups)
End Sub

Public Sub update_edwiz_method_info()

    Dim I As Integer
    Dim J As Integer
    I = 0
    If global_cur_property > -1 Then
        While I < MAX_METHODS_EACH - 1 And Trim(wiz_methods(global_cur_property, I)) <> ""
            frmeditwizard!optmethod(I).Visible = True
            frmeditwizard!optmethod(I).Caption = wiz_methods(global_cur_property, I)
            I = I + 1
        Wend
        For J = I To MAX_METHODS_EACH - 1
            frmeditwizard!optmethod(J).Caption = ""
            frmeditwizard!optmethod(J).value = False
            frmeditwizard!optmethod(J).Visible = False
        Next J
        ' set the first one as default
        If Trim(frmeditwizard!optmethod(0).Caption) <> "" Then
            frmeditwizard!optmethod(0).value = True
        Else
            MsgBox ("methods for " & input_name(global_cur_property) & " not yet implemented")
        End If
        
    End If
End Sub

Public Function get_num_rings(casarg As Long) As Integer

    If selected_rings <> -1 Then
        get_num_rings = selected_rings
    Else
        get_num_rings = 0
    End If
End Function
