Attribute VB_Name = "modexport"
Option Explicit
Global ex_design_existing As Boolean
Global ex_design_modified As Boolean
Global ex_from_design As Boolean
Global Const MAXFIELDS = 23
Dim delimiter As String
Dim fieldflags(MAXFIELDS) As Boolean
Dim fieldlist(MAXFIELDS) As Integer
Dim fieldcount As Integer
Global Const MAX_FIELDS As Integer = 23
Global Const MAX_DESIGNS As Integer = 20
Dim def_settings(MAX_DESIGNS, MAX_FIELDS) As Integer
Dim design_array(MAX_DESIGNS, MAX_DESIGNS) As String
Global Const CHEM_NAME  As Integer = 0
Global Const C_NUM As Integer = 1
Global Const CAS_NUM As Integer = 2
Global Const property_name As Integer = 3
Global Const PROPERTY_CODE As Integer = 4
Global Const value As Integer = 5
Global Const VALUE_UNITS As Integer = 6
Global Const OP_TEMP As Integer = 7
Global Const OP_TEMP_UNITS As Integer = 8
Global Const OP_PRESSURE As Integer = 9
Global Const OP_PRESSURE_UNITS As Integer = 10
Global Const MIN_TEMP As Integer = 11
Global Const MAX_TEMP As Integer = 12
Global Const TEMP_RANGE_UNITS As Integer = 13
Global Const COEFFICIENTS As Integer = 14
Global Const QUALITY_CHECKED_CODE As Integer = 15
Global Const RANKING_CODE As Integer = 16
Global Const DIPPR_CODE As Integer = 17
Global Const FAMILY_GROUPS As Integer = 18
Global Const REFERENCE As Integer = 19
Global Const CITATION As Integer = 20
Global Const METHOD_NAME As Integer = 21
Global Const ART_NO As Integer = 22
' a few extra constants for properties specific to Aspen
Global Const MV As Integer = 45         ' molar volume in m3/kmol
Global Const omega As Integer = 46      ' acentric factor
Global Const CPIGa As Integer = 47      ' ideal gas heat capacity in J/kmol-k
Global Const PLXANTb As Integer = 48    ' antoine eqn. parameters a, b and c
Global Const ZC As Integer = 49         ' critical compressibility factor


Public Sub do_new(exportfile As String, count As Integer, chemname_list() As String, export_list() As Long, field_list() As Integer, d_char As String)

    Dim J As Integer
    Dim FNum As Integer
    Dim casno As Long
    Dim filename As String
    Dim fullname As String
    Dim write_flag As Integer
    
    FRMExport.Hide
    FRMMain.Refresh
    If d_char = "space" Then
        delimiter = " "
    Else
        delimiter = ", "
    End If
        ' put the fields and the number of fields into the global array
    For J = 0 To MAXFIELDS
        If field_list(J) = -1 Then
            fieldlist(J) = -1
            fieldcount = J
            Exit For
        Else
            fieldlist(J) = field_list(J)
        End If
    Next J
        ' fill the array of flags indicating which fields are used
    Call set_flags
    
    filename = Trim(exportfile)
    On Error GoTo export_error
    Screen.MousePointer = 11
    For J = 0 To (count - 1)
        casno = export_list(J)
        If IsNull(casno) Then
            Exit For
        End If
        Cur_Info.CAS = casno
        Call Recalculate
        
        If J = 0 Then
            write_flag = 1
            If (export_chem(write_flag, filename, chemname_list(J), casno)) = False Then
                MsgBox ("Error exporting " & casno)
            End If
        Else
            write_flag = 2
            If Not (export_chem(write_flag, filename, chemname_list(J), casno)) Then
               MsgBox ("Error exporting " & casno)
            End If
        End If
    Next J
    Screen.MousePointer = 1
    Exit Sub

        
export_error:
    Screen.MousePointer = 1
    MsgBox ("Error exporting chemicals")
    Exit Sub


End Sub

Public Function do_append(exportfile As String, count As Integer, chemname_list() As String, export_list() As Long, field_list() As Integer, d_char As String) As Boolean

    Dim J As Integer
    Dim FNum As Integer
    Dim casno As Long
    Dim filename As String
    Dim fullname As String
    Dim write_flag As Integer
    Dim success As Boolean
    
    write_flag = 2
    FRMExport.Hide
    FRMMain.Refresh
        ' get the correct delimiter
    If d_char = "space" Then
        delimiter = " "
    Else
        delimiter = ","
    End If
        ' put the fields and number of fields into the global array
    For J = 0 To MAXFIELDS
        If field_list(J) = -1 Then
            fieldlist(J) = -1
            fieldcount = J
            Exit For
        Else
            fieldlist(J) = field_list(J)
        End If
    Next J
        ' fill the array of flags indicating which fields are used
    Call set_flags
    filename = Trim(exportfile)
    On Error GoTo export_error
    Screen.MousePointer = 11
    For J = 0 To (count - 1)
        casno = export_list(J)
        If IsNull(casno) Then
            Exit For
        End If
        Cur_Info.CAS = casno
        Call Recalculate
       
        If (export_chem(write_flag, filename, chemname_list(J), casno)) = False Then
            MsgBox ("Error exporting " & casno)
        End If
    Next J
    Screen.MousePointer = 1
    Exit Function

       
export_error:
    Screen.MousePointer = 1
    MsgBox ("Error exporting chemicals")
    Exit Function

End Function


Public Function get_propcode(propnum As Integer, value As String) As Boolean

        Select Case propnum
            Case BOD
                value = "1a"
            Case COD
                value = "1b"
            Case ThODcarb
                value = "1cc"
            Case ThODcomb
                value = "1cn"
            Case logKow
                value = "2a"
            Case logKoc
                value = "2c"
            Case BCF
                value = "2d"
            Case MW
                value = "3a"
            Case LD25
                value = "3b"
            Case LD
                value = "3bt"
            Case Schem
                value = "3c"
            Case mp
                value = "3d"
            Case NBP
                value = "3e"
            Case VP25
                value = "3f"
            Case VP
                value = "3g"
            Case Dair
                value = "3h"
            Case Dwater
                value = "3i"
            Case VV
                value = "3j"
            Case LV
                value = "3k"
            Case ST25
                value = "3l"
            Case ST
                value = "3lt"
            Case LTC
                value = "3ml"
            Case VTC
                value = "3mv"
            Case hfor
                value = "3n"
            Case LHC
                value = "3o"
            Case VHC
                value = "3p"
            Case CT
                value = "3q"
            Case CP
                value = "3r"
            Case CV
                value = "3s"
            Case Hvap25
                value = "3t"
            Case Hvap
                value = "3tt"
            Case HvapNBP
                value = "3tz"
            Case ACchem
                value = "4a"
            Case ACwater
                value = "4b"
            Case HC
                value = "4c"
            Case LFL
                value = "5al"
            Case UFL
                value = "5au"
            Case FP
                value = "5b"
            Case AIT
                value = "5c"
            Case Hcomb
                value = "5d"
            Case Swater
                value = "2b"
            Case Else
                value = " "
                get_propcode = False
                Exit Function
    End Select
    get_propcode = True
    
End Function

Public Function export_chem(write_flag As Integer, fname As String, chemname As String, casno As Long) As Boolean

    Dim FNum As Integer
    Dim count As Long
    Dim J As Integer
    Dim N As Integer
    Dim i As Integer
    Dim K As Integer
    Dim found As Boolean
    Dim outputline As String
    Dim propcode As String
    Dim propname As String
    Dim temp_cas As Long
    Dim temp_value As Double
    Dim temp_units As String
    Dim temp_opt As Double
    Dim temp_opp As Double
    Dim temp_optunit As String
    Dim temp_oppunit As String
    Dim temp_coef(5) As Double
    Dim temp_mint As Double
    Dim temp_maxt As Double
    Dim temp_cnum As Double
    Dim temp_eqnum As Integer
    Dim temp_propcode As String
    Dim temp_artno As Long
    Dim temp_method As String
    Dim temp_rating(10)
    Dim junk As String
    Dim field_header As String
    
        ' first figure out if we're starting a new file or appending to an existing one
        ' if it's a new one, put the field names in there
    On Error GoTo File_Closed_Error
    FNum = FreeFile
    If write_flag = 2 Then  ' its for append
        Open fname For Input As #FNum
            ' add something here to get the fields from the file and check
        Close #FNum
        Open fname For Append As #FNum
    Else
        ' first build an array of fields based on the selected fields for this design
        field_header = build_fields()
        Open fname For Output As #FNum
        Print #FNum, field_header
    End If
    
         On Error GoTo File_Open_Error
        ' found indicates that CurMethod (the method of choice) has been found
        ' first, these will be the same regardless of which method is used
    temp_cas = Cur_Disp.CAS
    temp_opt = Cur_Disp.OpT
    temp_opp = Cur_Disp.OpP
    temp_optunit = Trim(Cur_Disp.OpTUnit)
    temp_oppunit = Trim(Cur_Disp.OpPUnit)
    temp_cnum = 0
                           
    For i = 0 To NumProperties
        found = False
        propname = get_propname(i)
        For J = 1 To NumMethods
            If J = InfoMethod(i).CurMethod Then
                found = True
                temp_method = DispMethod(i).MethodName(J)
                    ' initialize these just in case we can't get them from the database
                temp_artno = 0
                Call get_db_temps(temp_cnum, temp_artno, temp_method, CLng(i))
                    ' resume next takes care of the null cases encountered
                On Error Resume Next
                    ' probably not necessary but here's a check that the property number is valid
                If get_propcode(i, temp_propcode) = False Then
                    GoTo end_export_prop
                End If
                    ' first, these are different depending on whether it's an f(t) property
                
                If is_f_of_t(i) = False Then
                    temp_value = DispMethod(i).value(J)
                    temp_units = Trim(DispMethod(i).Unit)
                    temp_eqnum = 0
                    For N = 0 To 4
                        temp_coef(N) = 0
                    Next N
                    temp_mint = 0
                    temp_maxt = 0
                    ' add more here for getting artno and cnum
                
                   ' Write #FNum, temp_cnum, temp_propcode, temp_artno, J, temp_method, temp_cas, _
                    '    temp_value, temp_units, temp_opt, _
                    '    , , , , , , junk, junk, junk, temp_opp, junk
                    'GoTo end_export_prop
                Else
                    temp_value = 0
                    temp_method = DispMethod(i).MethodName(J)
                    temp_units = Trim(DispMethod(i).Unit)
                    temp_eqnum = DispMethod(i).EqNum(J)
                    For N = 0 To 4
                        temp_coef(N) = DispMethod(i).Coeff(J, 2)
                    Next N
                    temp_mint = DispMethod(i).MinT(J)
                    temp_maxt = DispMethod(i).MaxT(J)
                End If
                    ' add more stuff for getting artno and c#
                    ' now figure out which fields to export
                 outputline = ""
                For K = 0 To fieldcount - 1
                    Select Case fieldlist(K)
                        Case 0:
                            outputline = outputline + Chr(34) + chemname + Chr(34) + delimiter
                        Case 1:
                            outputline = outputline + CStr(temp_cnum) + delimiter
                        Case 2:
                            outputline = outputline + CStr(temp_cas) + delimiter
                        Case 3:
                            outputline = outputline + Chr(34) + propname + Chr(34) + delimiter
                        Case 4:
                            outputline = outputline + Chr(34) + temp_propcode + Chr(34) + delimiter
                        Case 5:
                            outputline = outputline + CStr(temp_value) + delimiter
                        Case 6:
                            outputline = outputline + Chr(34) + temp_units + Chr(34) + delimiter
                        Case 7:
                            outputline = outputline + CStr(temp_opt) + delimiter
                        Case 8:
                            outputline = outputline + Chr(34) + temp_optunit + Chr(34) + delimiter
                        Case 9:
                            outputline = outputline + CStr(temp_opp) + delimiter
                        Case 10:
                            outputline = outputline + Chr(34) + temp_oppunit + delimiter
                        Case 11:
                            outputline = outputline + CStr(temp_mint) + delimiter
                        Case 12:
                            outputline = outputline + CStr(temp_maxt) + delimiter
                        Case 13:
                            outputline = outputline + Chr(34) + temp_optunit + Chr(34) + delimiter
                        Case 14:
                            outputline = outputline + CStr(temp_coef(0)) + delimiter + CStr(temp_coef(1)) + delimiter + CStr(temp_coef(2)) + delimiter + CStr(temp_coef(3)) + delimiter + CStr(temp_coef(4)) + delimiter
                        'Case 15:
                        '    Write #FNum, temp_qu_code, delimiter
                        'Case 16:
                        '    Write #FNum, temp_rank_code, delimiter
                        'Case 17:
                         '   Write #FNum, temp_dippr_code, delimiter
                        'Case 18:
                         '   Write #FNum, temp_family, delimiter
                        'Case 19:
                         '   Write #FNum, temp_ref, delimiter
                        'Case 20:
                         '   Write #FNum, temp_citation, delimiter
                        'Case 21:
                         '   Write #FNum, temp_method_name, delimiter
                        Case 22:
                            outputline = outputline + CStr(temp_artno) + delimiter
                        End Select
                    Next K
                    outputline = outputline + Chr(13)
                    Print #FNum, outputline
                        
                End If
                If found = True Then
                    GoTo end_export_prop
                End If
            Next J
end_export_prop:
            
        Next i
            
        
endloop:
   
    export_chem = True
    Close #FNum
    Exit Function
   
BadCasError:
    Resume endloop
    
File_Open_Error:
    
    MsgBox "Error exporting data ", 48, "Error"
    Close #FNum
    export_chem = False
    Exit Function

File_Closed_Error:
    MsgBox "Can't find export file"
    export_chem = False
    Exit Function

End Function

Public Function is_f_of_t(propnum As Integer) As Boolean

    Dim answer As Boolean
    Select Case propnum
        Case 2:
            answer = True
        Case 6:
            answer = True
        Case 8:
            answer = True
        Case 9:
            answer = True
        Case 12:
            answer = True
        Case 18:
            answer = True
        Case 19:
            answer = True
        Case 20:
            answer = True
        Case 21:
            answer = True
        Case 22:
            answer = True
        Case Else
            answer = False
        End Select
        is_f_of_t = answer
        
    
End Function

Private Sub get_db_temps(cnum As Double, artno As Long, mname As String, propnum As Long)
 
    Dim DBTbl As Recordset
   
        ' initialize these in case we can't find them
    cnum = 0
    artno = 0
    mname = ""
    On Error GoTo DB_Closed_Error
    Set DBTbl = DBJetMaster.OpenRecordset("DIPPR911", dbOpenTable)
    On Error GoTo DB_Open_Error
    DBTbl.Index = "PrimaryKey1"
    DBTbl.Seek "=", Cur_Disp.CAS
    
    If DBTbl.NoMatch Then
         DBTbl.Close
         Exit Sub
    End If
    cnum = DBTbl("C#")
    Do While DBTbl("CAS #") = Cur_Disp.CAS
        If (DBTbl("Method Name") = mname) And (DBTbl("PEARLS Code") = propnum) Then
            artno = DBTbl("Article #")
            DBTbl.Close
            Exit Sub
        End If
        DBTbl.MoveNext
    Loop
    
            ' not giving any message if this routine fails, the
            ' export file just won't have an entry for c# and article#
            ' for this property.  This way if they don't own the 911 database
            ' we won't have a problem
DB_Open_Error:
    
    DBTbl.Close
    Exit Sub
DB_Closed_Error:
    
End Sub

Private Function build_fields() As String

    Dim i As Integer
    Dim fieldstring As String
    Dim quotes As String
    ' this function needs to return a string that is the header of the export
    ' file based on which fields the (global) field_list contains
    
    ' first, if the delimiter is spaces, we'll need quotes around these to distinguish them
    If delimiter = " " Then
        quotes = Chr(34)
    Else
        quotes = ""
    End If
    
    fieldstring = ""
    For i = 0 To fieldcount - 1
        Select Case fieldlist(i)
            Case 0:
                fieldstring = fieldstring + quotes + "Chemical Name" + quotes + delimiter
            Case 1:
                fieldstring = fieldstring + quotes + "C #" + quotes + delimiter
            Case 2:
                fieldstring = fieldstring + quotes + "Cas #" + quotes + delimiter
            Case 3:
                fieldstring = fieldstring + quotes + "Property" + quotes + delimiter
            Case 4:
                fieldstring = fieldstring + quotes + "Property Code" + quotes + delimiter
            Case 5:
                fieldstring = fieldstring + quotes + "Value" + quotes + delimiter
            Case 6:
                fieldstring = fieldstring + quotes + "Value Units" + quotes + delimiter
            Case 7:
                fieldstring = fieldstring + quotes + "Operating Temp" + quotes + delimiter
            Case 8:
                fieldstring = fieldstring + quotes + "Op Temp Units" + quotes + delimiter
            Case 9:
                fieldstring = fieldstring + quotes + "Operating Pressure" + quotes + delimiter
            Case 10:
                fieldstring = fieldstring + quotes + "Op Pressure Units" + quotes + delimiter
            Case 11:
                fieldstring = fieldstring + quotes + "Min Temp" + quotes + delimiter
            Case 12:
                fieldstring = fieldstring + quotes + "Max Temp" + quotes + delimiter
            Case 13:
                fieldstring = fieldstring + quotes + "Temp Range Units" + quotes + delimiter
            Case 14:
                fieldstring = fieldstring + quotes + "Coefficients" + quotes + delimiter
            Case 15:
                fieldstring = fieldstring + quotes + "Quality Checked Code" + quotes + delimiter
            Case 16:
                fieldstring = fieldstring + quotes + "Ranking Code" + quotes + delimiter
            Case 17:
                fieldstring = fieldstring + quotes + "DIPPR Code" + quotes + delimiter
            Case 18:
                fieldstring = fieldstring + quotes + "Family Groups" + quotes + delimiter
            Case 19:
                fieldstring = fieldstring + quotes + "Reference" + quotes + delimiter
            Case 20:
                fieldstring = fieldstring + quotes + "Citation" + quotes + delimiter
            Case 21:
                fieldstring = fieldstring + quotes + "Method Name" + quotes + delimiter
            Case 22:
                fieldstring = fieldstring + quotes + "Article Number" + quotes + delimiter
            Case -1:
                Exit For
        End Select
    Next i
    If Right(fieldstring, 1) = "," Then
        fieldstring = Left(fieldstring, Len(fieldstring) - 1)
    End If
    build_fields = fieldstring
End Function

Private Sub set_flags()
    Dim i As Integer
    
    ' this function just sets flags to true or false depending
    ' on whether its a field for export
    ' first initialize all to false
    For i = 0 To MAXFIELDS - 1
        fieldflags(i) = False
    Next i
    For i = 0 To fieldcount - 1
        fieldflags(fieldlist(i)) = True
    Next i
        
End Sub

Private Function get_propname(property_number As Integer) As String

    Dim PropertyName As String
    PropertyName = ""
    Select Case property_number
        Case MW:
            PropertyName = "Molecular Weight"
        Case LD25:
            PropertyName = "Liquid Density @ 25C"
        Case LD:
            PropertyName = "Liquid Density as f(T)"
        Case mp:
            PropertyName = "Melting Point"
        Case NBP:
            PropertyName = "Normal Boiling Point"
        Case VP25:
            PropertyName = "Vapor Pressure @ 298.15 K"
        Case VP:
            PropertyName = "Vapor Pressure as f(T)"
        Case hfor:
            PropertyName = "Heat of Formation"
        Case LHC:
            PropertyName = "Liquid Heat Capacity as f(T)"
        Case VHC:
            PropertyName = "Vapor Heat Capacity as f(T)"
        Case Hvap25:
            PropertyName = "Heat of Vaporization @ 25C"
        Case HvapNBP:
            PropertyName = "Heat of Vaporization @ NBP"
        Case Hvap:
            PropertyName = "Heat of Vaporization as f(T)"
        Case CT:
            PropertyName = "Critical Temperature"
        Case CP:
            PropertyName = "Critical Pressure"
        Case CV:
            PropertyName = "Critical Volume"
        Case Dwater:
            PropertyName = "Diffusivity in Water"
        Case Dair:
            PropertyName = "Diffusivity in Air"
        Case ST25:
            PropertyName = "Surface Tension @ 25C"
        Case ST:
            PropertyName = "Surface Tension as f(T)"
        Case VV:
            PropertyName = "Vapor Viscosity as f(T)"
        Case LV:
            PropertyName = "Liquid Viscosity as f(T)"
        Case LTC:
            PropertyName = "Liquid Thermal Conductivity as f(T)"
        Case VTC:
            PropertyName = "Vapor Thermal Conductivity as f(T)"
        Case UFL:
            PropertyName = "Upper Flammability Limit"
        Case LFL:
            PropertyName = "Lower Flammability Limit"
        Case FP:
            PropertyName = "Flash Point"
        Case AIT:
            PropertyName = "Autoignition Temperature"
        Case Hcomb:
            PropertyName = "Heat of Combustion"
        Case ThODcarb:
            PropertyName = "Carbonaceous ThOD"
        Case ThODcomb:
            PropertyName = "Combined ThOD"
        Case COD:
            PropertyName = "Chemical Oxygen Demand"
        Case BOD:
            PropertyName = "Biochemical Oxygen Demand"
        Case ACwater:
            PropertyName = "Infinite Dilution Activity Coefficient of Water in Chemical"
        Case HC:
            PropertyName = "Henry's Constant"
        Case ACchem:
            PropertyName = "Infinite Dilution Activity Coefficient of Chemical in Water"
        Case logKow:
            PropertyName = "log Kow"
        Case logKoc:
            PropertyName = "log Koc"
        Case BCF:
            PropertyName = "Bioconcentration Factor"
        Case Schem:
            PropertyName = "Solubility Limit of Chemical in Water"
        Case Swater:
            PropertyName = "Solubility Limit of Water in Chemical"
        Case ZC:
            PropertyName = "Critical Compressibility Factor"
        Case omega:
            PropertyName = "Acentric Factor"
        Case CPIGa:
            PropertyName = "Ideal Gas Heat Capacity"
        Case PLXANTb:
            PropertyName = "Antoine Eqn. Parameters"
        Case MV:
            PropertyName = "Molar Volume at Tb"
    End Select
    get_propname = PropertyName
End Function

Public Sub load_form_export()
    Dim i As Integer
    Dim DBTbl As Recordset
    Dim prev_name As String
    Dim chemcount As Integer
    
    ex_design_existing = False
    ex_design_modified = False
    ex_from_design = False
    chemcount = 0
    On Error GoTo DB_Closed_Error
    
    ' clear the list boxes first
    FRMExport!chemlst.Clear
    FRMExport!alllst.Clear
    FRMExport!designlst.Clear
    
    Set DBTbl = DBJetUser.OpenRecordset("User List", dbOpenTable)
    On Error GoTo DB_Open_Error
    DBTbl.MoveFirst
    
    ' add name and cas # to list, enough white space so the user doesn't see the cas #
    While DBTbl.EOF = False
            FRMExport!chemlst.AddItem Trim(DBTbl("Name")) & "                                              " & CStr(DBTbl("CAS"))
            DBTbl.MoveNext
    Wend
    DBTbl.Close
    
    On Error GoTo other_error
    FRMExport!nametbx.Text = "unnamed"
    FRMExport!spaceopt.value = False
    FRMExport!commaopt.value = True
    
    ' fill the rest of the list boxes
    FRMExport!alllst.AddItem "Chemical Name"
    FRMExport!alllst.AddItem "C#"
    FRMExport!alllst.AddItem "Cas #"
    FRMExport!alllst.AddItem "Property Name"
    FRMExport!alllst.AddItem "Property Code"
    FRMExport!alllst.AddItem "Value"
    FRMExport!alllst.AddItem "Value Units"
    FRMExport!alllst.AddItem "Operating Temperature"
    FRMExport!alllst.AddItem "Op Temperature Units"
    FRMExport!alllst.AddItem "Operating Pressure"
    FRMExport!alllst.AddItem "Op Pressure Units"
    FRMExport!alllst.AddItem "Min Temperature"
    FRMExport!alllst.AddItem "Max Temperature"
    FRMExport!alllst.AddItem "Temperature Range Units"
    FRMExport!alllst.AddItem "F(t) Coefficients"
    FRMExport!alllst.AddItem "Reference"
    FRMExport!alllst.AddItem "Citation"
    FRMExport!alllst.AddItem "Method Name"
    FRMExport!alllst.AddItem "Article Number"
    FRMExport!alllst.AddItem "Quality Checked Code"
    FRMExport!alllst.AddItem "Ranking Code"
    FRMExport!alllst.AddItem "DIPPR Code"
    FRMExport!alllst.AddItem "Family Group"
    
        ' get the defined designs from the def file
    Call parse_def_for_export
        ' add the hard-coded designs (ASPEN)
    Call add_ASPEN_export
    Call CenterForm(FRMExport)
    
    Exit Sub
other_error:
    Resume Next
DB_Open_Error:
    MsgBox ("Error reading user database")
    DBTbl.Close
    Exit Sub
DB_Closed_Error:
    MsgBox ("Can't find user database.  Use " & Chr(34) & "file preferences" & Chr(34) & " to set path")
    Exit Sub
   
    Call CenterForm(FRMExport)
End Sub

Public Sub do_aspen_export(exportfile As String, count As Integer, chemname_list() As String, export_list() As Long, exporttype As Integer)
    ' exportfile = the file to write to
    ' count = the number of chemicals
    ' chemname = the list of Chemical names
    ' export_list = the list of cas numbers
    ' exporttype = 1 (new) or 2 (append)
    
    Dim FNum As Integer
    Dim i As Integer    ' the counter for each line
    Dim J As Integer    ' the counter for each chemical
    Dim Z As Integer
    Dim antione_params(5) As Double
    Dim cpig_param(10) As Double
    'Dim cpig(11) As Double
    Dim cpig As Double
    Dim zc_value As Double
    Dim temp_cv As Double
    Dim temp_pc As Double
    Dim temp_t As Double
    
    Dim vb_value As Double
    Dim omega_value As Double
    'Dim plxant(2) As Double
    Dim plxant As Double
    Dim vb As Double
    Dim num_output_lines As Integer
    Dim components(10) As String
    Dim comp_names(10) As String
    Dim outputline(10, 28) As String    ' properties for each chemical
    num_output_lines = 28
    FNum = FreeFile
    components(0) = "COMPONENTS   "
    comp_names(0) = "COMP-NAMES   "
    
    For J = 0 To count - 1
            ' recalculate this chemical
            Cur_Info.CAS = export_list(J)
            ' get zc_value, critical compressibility factor
            temp_pc = Convert(InfoMethod(CP).value(InfoMethod(CP).CurMethod), CP, InfoMethod(CP).Unit, "kPa ", False)
            temp_cv = Convert(InfoMethod(CV).value(InfoMethod(CV).CurMethod), CV, InfoMethod(CV).Unit, "m3/kmol", False)
            temp_t = Convert(Cur_Info.OpT, OptTemp, Cur_Info.OpTUnit, "K", False)
            ' gas_constant = 8.314
            If temp_t <> 0# Then
                zc_value = (temp_pc * temp_cv) / (8.314 * temp_t)
            Else
                zc_value = 0#
            End If
            
            'get all the values for THIS chemical
            Call Recalculate
                    
            'this value requires seperate calculation due to the way its defined
            omega_value = calc_acentric_factor
                        
            ' calculate antoine stuff
            Call get_antione_params(Cur_Info.CAS, antione_params)
            plxant = antione_params(0) + (antione_params(1) / (temp_t + antione_params(2)))
 
            ' calculate cpig and it's parameters, for now use BIP 1 for coefficients
            cpig = 0#
                'For I = 1 To NumMethods
                '    If infomethod(VHC).EqNum(I) <> 0 Then
                '        cpig = 0#
                '        cpig = CalcFofT(VHC, I)
                '        If cpig <> 0# And cpig <> ERROR_FLAG Then
                '            For Z = 0 To 4
                '                cpig_param(Z) = infomethod(VHC).Coeff(I, Z)
                '            Next Z
                '            Exit For
                '        End If
                '    End If
                'Next I
                
                'USE 5/2 METHOD TO CALC
            cpig_param(0) = 1000 * 8.314 * (5 / 2)
            For Z = 1 To 7
                cpig_param(Z) = 0#
            Next Z
            cpig_param(8) = 0
            cpig_param(9) = 1000
                  
           ' next the molar volume = density at normal boiling point
            vb_value = 0#
            For i = 1 To NumMethods
                If InfoMethod(LD).EqNum(i) <> 0 Then
                    vb_value = 0#
                    vb_value = CalcFofT(LD, i)
                    If vb_value <> 0# And vb_value <> ERROR_FLAG Then
                        vb_value = Convert(vb_value, LD, Get_DefaultUnit(LD), "m3/kmol", False)
                        Exit For
                    End If
                End If
            Next i
            
            ' add names to the list of components and component names and
            ' then build the lines for this chemical
            
            components(J) = components(J) & Cur_Info.Formula & Cur_Info.Formula
            comp_names(J) = comp_names(J) & Cur_Info.Formula & Cur_Info.name
            outputline(J, 0) = "COMP-LIST   " & Cur_Info.name
            outputline(J, 1) = "    CVAL   MW      1  1   " & InfoMethod(MW).value(InfoMethod(MW).CurMethod)
            outputline(J, 2) = "    CVAL   ZC      1  1   " & zc_value
            outputline(J, 3) = "    CVAL   PC      1  1   " & Convert(InfoMethod(CP).value(InfoMethod(CP).CurMethod), CP, InfoMethod(CP).Unit, "N/m2", False)
            outputline(J, 4) = "    CVAL   VC      1  1   " & Convert(InfoMethod(CV).value(InfoMethod(CV).CurMethod), CV, InfoMethod(CV).Unit, "m3/kmol", False)
            outputline(J, 5) = "    CVAL   VB      1  1   " & vb_value  ' already converted
            outputline(J, 6) = "    CVAL   TC      1  1   " & Convert(InfoMethod(CT).value(InfoMethod(CT).CurMethod), CT, InfoMethod(CT).Unit, "K", False)
            outputline(J, 7) = "    CVAL   TB      1  1   " & Convert(InfoMethod(NBP).value(InfoMethod(NBP).CurMethod), NBP, InfoMethod(NBP).Unit, "K", False)
            outputline(J, 8) = "    CVAL   OMEGA   1  1   " & omega_value
            outputline(J, 9) = "    CVAL   DHFORM  1  1   " & InfoMethod(hfor).value(InfoMethod(hfor).CurMethod)
            outputline(J, 10) = "    CVAL   DHVLWT  1  1   " & InfoMethod(HvapNBP).value(InfoMethod(HvapNBP).CurMethod) & " /"
            outputline(J, 11) = "                      2   "
            outputline(J, 12) = "    CVAL   CPIG    1  1   " & cpig_param(0) & " /"
            outputline(J, 13) = "                      2   " & cpig_param(1) & " /"
            outputline(J, 14) = "                      3   " & cpig_param(2) & " /"
            outputline(J, 15) = "                      4   " & cpig_param(3) & " /"
            outputline(J, 16) = "                      5   " & cpig_param(4) & " /"
            outputline(J, 17) = "                      6   " & cpig_param(5) & " /"
            outputline(J, 18) = "                      7   " & cpig_param(6) & " /"
            outputline(J, 19) = "                      8   " & cpig_param(7) & " /"
            outputline(J, 20) = "                      9   " & cpig_param(8) & " /"
            outputline(J, 21) = "                      10  " & cpig_param(9)
            outputline(J, 22) = "    CVAL   PLXANT  1  1   " & antione_params(0) & " /"
            outputline(J, 23) = "                      2   " & antione_params(1) & " /"
            outputline(J, 24) = "                      3   " & antione_params(2) & " /"
            outputline(J, 25) = "                      4   " & antione_params(3) & " /"
            outputline(J, 26) = "                      5   " & antione_params(4)
    Next J
    ' print the stuff to the file
    If exporttype = 1 Then
        Open exportfile For Output As FNum
    Else
        Open exportfile For Append As FNum
    End If
    
    For i = 0 To count - 1
        Print #FNum, components(i)
        Print #FNum, comp_names(i)
        For J = 0 To num_output_lines - 1
            Print #FNum, outputline(i, J)
        Next J
    Next i
    Close FNum
    If count <> 1 Then
        MsgBox ("Successfully exported " & count & " chemicals")
    Else
        MsgBox ("Successfully exported " & count & " chemical")
    End If
End Sub

Public Sub parse_def_for_export()

    Dim FNum As Integer
    Dim i As Integer
    Dim J As Integer
    Dim K As Integer
    Dim def_line As String
    Dim def_list(20) As String
    Dim get_char As String
    Dim design_name As String
    Dim temp As String
    Dim getchar As String
    Dim delimiter As String
    Dim size_def_line As Integer
    'On Error GoTo end_of_sub
    For i = 0 To MAX_DESIGNS - 1
        For J = 0 To MAX_FIELDS - 1
            def_settings(i, J) = -1
        Next J
    Next i
    FNum = FreeFile
    Open App.path & "\" & deffile For Input As FNum
    i = 0
    While Not EOF(FNum) And i < MAX_DESIGNS
        Input #FNum, def_line
        temp = ""
        design_name = ""
        get_char = ""
        If Left(def_line, 14) = "Pearls.Export." Then
            On Error GoTo end_of_loop
                ' first get the design name
            def_line = Right(def_line, Len(def_line) - 14)
            get_char = Left(def_line, 1)
            def_line = Right(def_line, Len(def_line) - 1)
        
            While get_char <> "$"
                design_name = design_name & get_char
                get_char = Left(def_line, 1)
                def_line = Right(def_line, Len(def_line) - 1)
            Wend
                ' now we have to get the delimiter (comma or space) from the file
            FRMExport!designlst.AddItem design_name
            design_array(i, 0) = design_name

            get_char = Left(def_line, 1)
            def_line = Right(def_line, Len(def_line) - 1)
            delimiter = ""
            While get_char <> "$"
                delimiter = delimiter & get_char
                get_char = Left(def_line, 1)
                def_line = Right(def_line, Len(def_line) - 1)
            Wend
            design_array(i, 1) = delimiter
                '
            J = 0
            get_char = Left(def_line, 1)
            def_line = Right(def_line, Len(def_line) - 1)
            While get_char <> "$" And Len(def_line) > 1
                    
                'getchar = Left(def_line, 1)
                While get_char <> "," And Len(def_line) > 1
                    temp = temp + get_char
                    get_char = Left(def_line, 1)
                    def_line = Right(def_line, Len(def_line) - 1)
                    'get_char = Left(def_line, 1)
                    If get_char = "-" Then
                        GoTo finishline
                    End If
                Wend
                If Right(temp, 1) = "," Then
                    temp = Left(temp, Len(temp) - 1)
                ElseIf Left(temp, 1) = "," Then
                    temp = Right(temp, Len(temp) - 1)
                End If
                If Len(temp) > 0 Then
                    def_settings(i, J) = CInt(temp)
                    get_char = ""
                    temp = ""
                    design_name = ""
                    delimiter = ""
                End If

                J = J + 1
            Wend
finishline:
            def_settings(i, J) = -1
            i = i + 1
        End If
end_of_loop:
     Wend
     Close #FNum
     

end_of_sub:


End Sub

Public Sub write_def_for_export()

    Dim FNum As Integer
    Dim i As Integer
    Dim J As Integer
    Dim filelines(50) As String
    Dim outputline As String
    'On Error GoTo file_closed_error
    FNum = FreeFile
    Open App.path & "\" & deffile For Input As FNum
    'On Error GoTo file_open_error
        ' first read in the entire contents of the file
    While Not EOF(FNum)
        Input #FNum, filelines(i)
        i = i + 1
    Wend
    Close FNum
    Open App.path & "\" & deffile For Output As FNum

        ' now find out where the export stuff starts
        ' and write the rest back to the defs file
    For J = 0 To i - 1
        If Left(filelines(J), 14) = "Pearls.Export." Then
            Exit For
        Else
            Write #FNum, filelines(J)
        End If
    Next J
    For i = 0 To FRMExport!designlst.ListCount - 1
        If Trim(FRMExport!designlst.List(i)) <> "ASPEN" Then    ' don't write the hard-coded designs in
    
            outputline = "Pearls.Export." & FRMExport!designlst.List(i) & "$" & design_array(i, 1) & "$"
            For J = 0 To MAX_FIELDS - 1
                If def_settings(i, J) <> -1 Then
                    outputline = outputline & CStr(def_settings(i, J))
                    outputline = outputline & ","
                Else
                    Exit For
                End If
                
            Next J
            outputline = outputline & "-1$$"
            Write #FNum, outputline
        End If
    Next i
        
    
    Close FNum
    Exit Sub
File_Open_Error:
    Close FNum
    Exit Sub
File_Closed_Error:
    



End Sub

Public Sub add_ASPEN_export()

    ' this just adds ASPEN to the list, including all it's properties to the
    ' design list (def_settings)
    Dim position_num As Integer
    ' add the ASPEN format to the export designs
    position_num = FRMExport!designlst.ListCount
    FRMExport!designlst.AddItem "ASPEN"
    If position_num < 0 Then
        position_num = 0
    End If
    design_array(position_num, 0) = "ASPEN"
    design_array(position_num, 1) = "space" ' this is ignored anyway but ...
    def_settings(position_num, 0) = MW  ' Molecular Weight
    def_settings(position_num, 1) = CP  ' Critical Pressure
    def_settings(position_num, 2) = CV  ' Critical Volume
    def_settings(position_num, 3) = CT  ' Critical Temperature
    def_settings(position_num, 4) = NBP ' Normal Boiling Point
    def_settings(position_num, 5) = MV  ' Molar Volume at Tb
    def_settings(position_num, 0) = ZC ' Critical Compressibility Factor
    def_settings(position_num, 7) = omega   ' Acentric Factor
    def_settings(position_num, 8) = CPIGa   ' Ideal Gas Heat Capacity
    def_settings(position_num, 9) = hfor  ' Ideal Gas Heat of Formation
    def_settings(position_num, 10) = VHC ' Heat of Vap at Tb
    def_settings(position_num, 11) = PLXANTb    ' Antoine Eqn Parameters

End Sub

Public Sub fill_design_array()

    ' this function just modifies the array of settings
    ' the def file isn't written to until the form
    ' is deactivated
    ' it also adds the design to the design list if it's a new one
    Dim i As Integer
    Dim design_num As Integer
    Dim next_item As Integer
    Dim found As Boolean
    found = False
    
    For i = 0 To FRMExport!designlst.ListCount - 1
        If FRMExport!designlst.List(i) = FRMExport!nametbx.Text Then
            design_num = i
            design_array(i, 0) = FRMExport!nametbx.Text
            If FRMExport!spaceopt.value = True Then
                design_array(design_num, 1) = "space"
            Else
                design_array(design_num, 1) = "comma"
            End If
            found = True
            Exit For
        End If
    Next i
    If found = False Then
        next_item = FRMExport!designlst.ListCount
        design_array(next_item, 0) = FRMExport!nametbx.Text
        If FRMExport!spaceopt.value = True Then
            design_array(next_item, 1) = "space"
        Else
            design_array(next_item, 1) = "comma"
        End If
        FRMExport!designlst.AddItem FRMExport!nametbx.Text
    End If
    For i = 0 To FRMExport!selectlst.ListCount - 1
        Select Case FRMExport!selectlst.List(i)
            Case "Chemical Name"
                def_settings(next_item, i) = CHEM_NAME
            Case "C#"
                def_settings(next_item, i) = C_NUM
            Case "Cas #"
                def_settings(next_item, i) = CAS_NUM
            Case "Property Name"
                def_settings(next_item, i) = property_name
            Case "Property Code"
                def_settings(next_item, i) = PROPERTY_CODE
            Case "Value"
                def_settings(next_item, i) = value
            Case "Value Units"
                def_settings(next_item, i) = VALUE_UNITS
            Case "Operating Temperature"
                def_settings(next_item, i) = OP_TEMP
            Case "Op Temperature Units"
                def_settings(next_item, i) = OP_TEMP_UNITS
            Case "Operating Pressure"
                def_settings(next_item, i) = OP_PRESSURE
            Case "Op Pressure Units"
                def_settings(next_item, i) = OP_PRESSURE_UNITS
            Case "Min Temperature"
                def_settings(next_item, i) = MIN_TEMP
            Case "Max Temperature"
                def_settings(next_item, i) = MAX_TEMP
            Case "Temperature Range Units"
                def_settings(next_item, i) = TEMP_RANGE_UNITS
            Case "F(t) Coefficients"
                def_settings(next_item, i) = COEFFICIENTS
            Case "Reference"
                def_settings(next_item, i) = REFERENCE
            Case "Citation"
                def_settings(next_item, i) = CITATION
            Case "Method Name"
                def_settings(next_item, i) = METHOD_NAME
            Case "Quality Checked Code"
                def_settings(next_item, i) = QUALITY_CHECKED_CODE
            Case "Ranking Code"
                def_settings(next_item, i) = RANKING_CODE
            Case "DIPPR Code"
                def_settings(next_item, i) = DIPPR_CODE
            Case "Family Group"
                def_settings(next_item, i) = FAMILY_GROUPS
            Case Else
                
        End Select
   Next i
   def_settings(design_num, i) = -1
End Sub




Public Sub fill_fields(fields() As Integer)

    Dim i As Integer
    
    ' a function to fill an array of fields, called if
    ' the design being used isn't a saved one so we
    ' need to go through the selectlst to get the right
    ' fields
    For i = 0 To FRMExport!selectlst.ListCount - 1
        Select Case FRMExport!selectlst.List(i)
            Case "Chemical Name"
                fields(i) = CHEM_NAME
            Case "C#"
                fields(i) = C_NUM
            Case "Cas #"
                fields(i) = CAS_NUM
            Case "Property Name"
                fields(i) = property_name
            Case "Property Code"
                fields(i) = PROPERTY_CODE
            Case "Value"
                fields(i) = value
            Case "Value Units"
                fields(i) = VALUE_UNITS
            Case "Operating Temperature"
                fields(i) = OP_TEMP
            Case "Op Temperature Units"
                fields(i) = OP_TEMP_UNITS
            Case "Operating Pressure"
                fields(i) = OP_PRESSURE
            Case "Op Pressure Units"
                fields(i) = OP_PRESSURE_UNITS
            Case "Min Temperature"
                fields(i) = MIN_TEMP
            Case "Max Temperature"
                fields(i) = MAX_TEMP
            Case "Temperature Range Units"
                fields(i) = TEMP_RANGE_UNITS
            Case "F(t) Coefficients"
                fields(i) = COEFFICIENTS
            Case "Quality Checked Code"
                fields(i) = QUALITY_CHECKED_CODE
            Case "Reference"
                fields(i) = REFERENCE
            Case "Citation"
                fields(i) = CITATION
            Case "Method Name"
                fields(i) = METHOD_NAME
            Case "Ranking Code"
                fields(i) = RANKING_CODE
            Case "DIPPR Code"
                fields(i) = DIPPR_CODE
            Case "Family Group"
                fields(i) = FAMILY_GROUPS
            Case "Article Number"
                fields(i) = ART_NO
            Case Else
                i = FRMExport!selectlst.ListCount
        End Select
   Next i
   fields(i) = -1
End Sub



Public Function calc_acentric_factor() As Double

    ' a procedure to calculate acentric factor (added for ASPEN export feature)
    
    ' inputs:   temperature reduced (K) = .7
    '           Critical TEmperature    (K)
    '           Critical Pressure   (atm)
    '           Vapor Pressure at Temperature reduced (.7)
    '  equation:    omega = - log(VPr @ Tr = .7) - 1.0
    
    ' 6-15-98 :  removed recalculate and replaced with direct calls for just the
    '   properties needed. (jem)
    '  equation: omega = -log (Vapor pressure@ Tr =.7 * CT  / reduced pressure) -1.0
    
    Dim omega As Double
    Dim VP_reduced As Double
    Dim vp_Critical As Double
    Dim prev_temperature As Double
    
    ' set the temperature to .7 * Critical Tempreture
    prev_temperature = Cur_Info.OpT
    InfoMethod(VP).TFT = 0.7 * Convert(InfoMethod(CT).value(InfoMethod(CT).CurMethod), CT, InfoMethod(CT).Unit, "K ", False)
    
    'find vapor pressure at .7 * critical temp then reset temp
    VP_reduced = CalcFofT(VP, InfoMethod(VP).CurMethod)
    'VP_reduced = Convert(VP, infomethod(VP).Unit, "Pa ", VP_reduced)

    InfoMethod(VP).TFT = prev_temperature
    
    'get critical pressure
    vp_Critical = Convert(InfoMethod(CP).value(InfoMethod(CP).CurMethod), CP, InfoMethod(CP).Unit, "Pa ", False)

    If VP_reduced >= 0# Then
            omega = -(log10(VP_reduced / vp_Critical)) - 1#
    Else
        omega = 0#
    End If
    
    'return value
    calc_acentric_factor = omega

End Function

Public Sub update_design_description()

    Dim field_list(MAX_FIELDS) As Boolean
    Dim cur_index As Integer
    Dim i As Integer
    Dim field_position As Integer
    cur_index = FRMExport!designlst.ListIndex
    If cur_index = -1 Then Exit Sub
    ex_design_existing = False
    FRMExport!nametbx.Text = FRMExport!designlst.List(cur_index)
    ex_from_design = True
    For i = 0 To FRMExport!designlst.ListCount - 1
        If FRMExport!designlst.List(cur_index) = design_array(i, 0) Then
            If design_array(i, 1) = "space" Then
                FRMExport!spaceopt.value = True
                FRMExport!commaopt.value = False
            Else
                FRMExport!spaceopt.value = False
                FRMExport!commaopt.value = True
            End If
        End If
    Next i
    FRMExport!selectlst.Clear
    FRMExport!alllst.Clear
       ' first initialize all these to false
    For i = 0 To MAX_FIELDS - 1
        field_list(i) = False
    Next i
    ' first take care of the hard-coded designs (ASPEN)
    If FRMExport!designlst.List(cur_index) = "ASPEN" Then
        FRMExport!selectlst.AddItem "Molecular Weight"
        FRMExport!selectlst.AddItem "Critical Pressure"
        FRMExport!selectlst.AddItem "Critical Volume"
        FRMExport!selectlst.AddItem "Critical Temperature"
        FRMExport!selectlst.AddItem "Normal Boiling Point"
        FRMExport!selectlst.AddItem "Molar Volume"
        FRMExport!selectlst.AddItem "Critical Compressibility Factor"
        FRMExport!selectlst.AddItem "Acentric Factor"
        FRMExport!selectlst.AddItem "Ideal Gas Heat Capacity"
        FRMExport!selectlst.AddItem "Ideal Gas Heat of Formation"
        FRMExport!selectlst.AddItem "Heat of Vaporization at NBP"
        FRMExport!selectlst.AddItem "Antoine Equation Parameters"
    
        FRMExport!alllst.AddItem "Molecular Weight"
        FRMExport!alllst.AddItem "Critical Pressure"
        FRMExport!alllst.AddItem "Critical Volume"
        FRMExport!alllst.AddItem "Critical Temperature"
        FRMExport!alllst.AddItem "Normal Boiling Point"
        FRMExport!alllst.AddItem "Molar Volume"
        FRMExport!alllst.AddItem "Critical Compressibility Factor"
        FRMExport!alllst.AddItem "Acentric Factor"
        FRMExport!alllst.AddItem "Ideal Gas Heat Capacity"
        FRMExport!alllst.AddItem "Ideal Gas Heat of Formation"
        FRMExport!alllst.AddItem "Heat of Vaporization at NBP"
        FRMExport!alllst.AddItem "Antoine Equation Parameters"
        Exit Sub
    End If
    
        ' this looks at the array of fields for this design,
        ' the fields for any numbers there will be set to true
        ' in the field list so we know what should be in each list
        ' on the form
        ' doing the actual adding to the lists separately just
        ' to make sure we don't end up with duplicates in either
        ' list
        field_position = 0
    For i = 0 To MAX_FIELDS - 1
            Select Case def_settings(cur_index, i)
                Case CHEM_NAME:
                    If field_list(CHEM_NAME) <> True Then
                        FRMExport!selectlst.AddItem "Chemical Name", field_position
                        field_position = field_position + 1
                    End If
                    field_list(CHEM_NAME) = True
                Case C_NUM:
                    If field_list(C_NUM) <> True Then
                        FRMExport!selectlst.AddItem "C#", field_position
                        field_position = field_position + 1
                    End If
                    field_list(C_NUM) = True
                Case CAS_NUM:
                    If field_list(CAS_NUM) <> True Then
                        FRMExport!selectlst.AddItem "Cas #", field_position
                        field_position = field_position + 1
                    End If
                    field_list(CAS_NUM) = True
                Case property_name:
                    If field_list(property_name) <> True Then
                        FRMExport!selectlst.AddItem "Property Name", field_position
                        field_position = field_position + 1
                    End If
                    field_list(property_name) = True
                Case PROPERTY_CODE:
                    If field_list(PROPERTY_CODE) <> True Then
                        FRMExport!selectlst.AddItem "Property Code", field_position
                        field_position = field_position + 1
                    End If
                    field_list(PROPERTY_CODE) = True
                Case value:
                    If field_list(value) <> True Then
                        FRMExport!selectlst.AddItem "Value", field_position
                        field_position = field_position + 1
                    End If
                    field_list(value) = True
                Case VALUE_UNITS:
                    If field_list(VALUE_UNITS) <> True Then
                        FRMExport!selectlst.AddItem "Value Units", field_position
                        field_position = field_position + 1
                    End If
                    field_list(VALUE_UNITS) = True
                Case OP_TEMP:
                    If field_list(OP_TEMP) <> True Then
                        FRMExport!selectlst.AddItem "Operating Temperature", field_position
                        field_position = field_position + 1
                    End If
                    field_list(OP_TEMP) = True
                Case OP_TEMP_UNITS
                    If field_list(OP_TEMP_UNITS) <> True Then
                        FRMExport!selectlst.AddItem "Op Temperature Units", field_position
                        field_position = field_position + 1
                    End If
                    field_list(OP_TEMP_UNITS) = True
                Case OP_PRESSURE:
                    If field_list(OP_PRESSURE) <> True Then
                        FRMExport!selectlst.AddItem "Operating Pressure", field_position
                        field_position = field_position + 1
                    End If
                    field_list(OP_PRESSURE) = True
                Case OP_PRESSURE_UNITS:
                    If field_list(OP_PRESSURE_UNITS) <> True Then
                        FRMExport!selectlst.AddItem "Op Pressure Units", field_position
                        field_position = field_position + 1
                    End If
                    field_list(OP_PRESSURE_UNITS) = True
                Case MIN_TEMP:
                    If field_list(MIN_TEMP) <> True Then
                        FRMExport!selectlst.AddItem "Min Temperature", field_position
                        field_position = field_position + 1
                    End If
                    field_list(MIN_TEMP) = True
                Case MAX_TEMP:
                    If field_list(MAX_TEMP) <> True Then
                        FRMExport!selectlst.AddItem "Max Temperature", field_position
                        field_position = field_position + 1
                    End If
                    field_list(MAX_TEMP) = True
                Case TEMP_RANGE_UNITS:
                    If field_list(TEMP_RANGE_UNITS) <> True Then
                        FRMExport!selectlst.AddItem "Temperature Range Units", field_position
                        field_position = field_position + 1
                    End If
                    field_list(TEMP_RANGE_UNITS) = True
                Case COEFFICIENTS:
                    If field_list(COEFFICIENTS) <> True Then
                        FRMExport!selectlst.AddItem "F(t) Coefficients", field_position
                        field_position = field_position + 1
                    End If
                    field_list(COEFFICIENTS) = True
                Case REFERENCE:
                    If field_list(REFERENCE) <> True Then
                        FRMExport!selectlst.AddItem "Reference", field_position
                        field_position = field_position + 1
                    End If
                    field_list(REFERENCE) = True
                Case CITATION:
                    If field_list(CITATION) <> True Then
                        FRMExport!selectlst.AddItem "Citation", field_position
                        field_position = field_position + 1
                    End If
                    field_list(CITATION) = True
                Case METHOD_NAME:
                    If field_list(METHOD_NAME) <> True Then
                        FRMExport!selectlst.AddItem "Method Name", field_position
                        field_position = field_position + 1
                    End If
                    field_list(METHOD_NAME) = True
                Case QUALITY_CHECKED_CODE:
                    If field_list(QUALITY_CHECKED_CODE) <> True Then
                        FRMExport!selectlst.AddItem "Quality Checked Code", field_position
                        field_position = field_position + 1
                    End If
                    field_list(QUALITY_CHECKED_CODE) = True
                Case RANKING_CODE:
                    If field_list(RANKING_CODE) <> True Then
                        FRMExport!selectlst.AddItem "Ranking Code", field_position
                        field_position = field_position + 1
                    End If
                    field_list(RANKING_CODE) = True
                Case DIPPR_CODE:
                    If field_list(DIPPR_CODE) <> True Then
                        FRMExport!selectlst.AddItem "DIPPR Code", field_position
                        field_position = field_position + 1
                    End If
                    field_list(DIPPR_CODE) = True
                Case FAMILY_GROUPS:
                    If field_list(FAMILY_GROUPS) <> True Then
                        FRMExport!selectlst.AddItem "Family Group", field_position
                        field_position = field_position + 1
                    End If
                    field_list(FAMILY_GROUPS) = True
                Case Else
                    i = MAX_FIELDS
            End Select
    Next i
        ' now add to the alllst if it's false
    For i = 0 To MAX_FIELDS - 1
        Select Case i
                Case CHEM_NAME:
                    If field_list(i) <> True Then
                        FRMExport!alllst.AddItem "Chemical Name"
                    End If
                Case C_NUM:
                    If field_list(i) <> True Then
                        FRMExport!alllst.AddItem "C#"
                    End If
                Case CAS_NUM:
                    If field_list(i) <> True Then
                        FRMExport!alllst.AddItem "Cas #"
                    End If
                Case property_name:
                    If field_list(i) <> True Then
                        FRMExport!alllst.AddItem "Property Name"
                    End If
                Case PROPERTY_CODE:
                    If field_list(i) <> True Then
                        FRMExport!alllst.AddItem "Property Code"
                    End If
                Case value:
                    If field_list(i) <> True Then
                        FRMExport!alllst.AddItem "Value"
                    End If
                Case VALUE_UNITS:
                    If field_list(i) <> True Then
                        FRMExport!alllst.AddItem "Value Units"
                    End If
                Case OP_TEMP:
                    If field_list(i) <> True Then
                        FRMExport!alllst.AddItem "Operating Temperature"
                    End If
                Case OP_TEMP_UNITS:
                    If field_list(i) <> True Then
                        FRMExport!alllst.AddItem "Op Temperature Units"
                    End If
                Case OP_PRESSURE:
                    If field_list(i) <> True Then
                        FRMExport!alllst.AddItem "Operating Pressure"
                    End If
                Case OP_PRESSURE_UNITS:
                    If field_list(i) <> True Then
                        FRMExport!alllst.AddItem "Op Pressure Units"
                    End If
                Case MIN_TEMP:
                    If field_list(i) <> True Then
                        FRMExport!alllst.AddItem "Min Temperature"
                    End If
                Case MAX_TEMP:
                    If field_list(i) <> True Then
                        FRMExport!alllst.AddItem "Max Temperature"
                    End If
                Case TEMP_RANGE_UNITS:
                    If field_list(TEMP_RANGE_UNITS) <> True Then
                        FRMExport!alllst.AddItem "Temperature Range Units"
                    End If
                Case COEFFICIENTS:
                    If field_list(COEFFICIENTS) <> True Then
                        FRMExport!alllst.AddItem "F(t) Coefficients"
                    End If
                Case METHOD_NAME:
                    If field_list(METHOD_NAME) <> True Then
                        FRMExport!alllst.AddItem "Method Name"
                    End If
                Case REFERENCE:
                    If field_list(REFERENCE) <> True Then
                        FRMExport!alllst.AddItem "Reference"
                    End If
                Case CITATION:
                    If field_list(CITATION) <> True Then
                        FRMExport!alllst.AddItem "Citation"
                    End If
                Case QUALITY_CHECKED_CODE:
                    If field_list(QUALITY_CHECKED_CODE) <> True Then
                        FRMExport!alllst.AddItem "Quality Checked Code"
                    End If
                Case RANKING_CODE:
                    If field_list(i) <> True Then
                        FRMExport!alllst.AddItem "Ranking Code"
                    End If
                Case DIPPR_CODE:
                    If field_list(i) <> True Then
                         FRMExport!alllst.AddItem "DIPPR Code"
                    End If
                Case FAMILY_GROUPS:
                    If field_list(i) <> True Then
                       FRMExport!alllst.AddItem "Family Group"
                    End If
                Case Else
                
            End Select
        Next i
        ex_design_existing = True
        ex_design_modified = False
End Sub

Public Sub remove_design()


    Dim Position As Integer
    Dim M As Integer
    Dim N As Integer
    Dim i As Integer
    Dim found As Boolean
    Dim hold_to_remove(MAX_DESIGNS) As Boolean
    
    Position = 0    ' the position of the next element  in the updated array
        
        ' just ignore it if there are no designs in the list
    If FRMExport!designlst.ListCount < 1 Then
        Exit Sub
    End If
    
        ' prompt user to select one if none are selected
    If FRMExport!designlst.SelCount < 1 Then
        MsgBox ("Select a design to remove")
        Exit Sub
    End If
    
        ' this assumes multi-select is possible
    For i = 0 To FRMExport!designlst.ListCount - 1
        If FRMExport!designlst.Selected(i) <> True Then
            ' keep this design
            design_array(Position, 0) = design_array(i, 0)
            design_array(i, 0) = ""     ' make it null
            design_array(Position, 1) = design_array(i, 1)
            design_array(i, 1) = ""     ' make it null
            For M = 0 To MAX_FIELDS - 1
                def_settings(Position, M) = def_settings(i, M)
                def_settings(i, M) = -1 ' make it null
            Next M
            hold_to_remove(i) = False
            Position = Position + 1
        Else
            hold_to_remove(i) = True
        End If
        
    Next i
    
        ' now actually remove the items from the list
    For i = 0 To FRMExport!designlst.ListCount - 1
        If hold_to_remove(i) = True Then
            FRMExport!designlst.RemoveItem i
        End If
    Next i






End Sub

Public Sub start_new_export(exportfile As String)
Dim J As Integer
    Dim delimiter As String
    Dim i As Integer
    Dim chemcount As Integer
    ReDim chemlist(FRMExport!chemlst.SelCount) As String
    Dim fields(MAX_FIELDS) As Integer
    Dim found As Boolean
    ' make sure a chemical is selected
    
    
    ReDim caslist(FRMExport!chemlst.SelCount) As Long
        ' first get the list of chemicals to be exported
    chemcount = 0
    For i = 0 To FRMExport!chemlst.ListCount - 1
        If FRMExport!chemlst.Selected(i) = True Then
            chemlist(chemcount) = Trim(Left(FRMExport!chemlst.List(i), Len(FRMExport!chemlst.List(i)) - 9))
            caslist(chemcount) = CLng(Trim(Right(FRMExport!chemlst.List(i), 9)))
            chemcount = chemcount + 1
        End If
    Next i
        
    
    
    ' If it's s hard-coded format, do that
    If Trim(FRMExport!nametbx.Text) = "ASPEN" Then
        Call do_aspen_export(exportfile, chemcount, chemlist, caslist, 1) ' the 1 is for new
        GoTo done_export
    End If
        ' fill the array of fields we're exporting
    found = False
    For i = 0 To FRMExport!designlst.ListCount - 1
        If FRMExport!nametbx.Text = FRMExport!designlst.List(i) Then
            found = True
            Exit For
        End If
    Next i
    If found = False Then
        Call fill_fields(fields)
        If FRMExport!spaceopt.value = True Then
            delimiter = "space"
        Else
            delimiter = "comma"
        End If
    Else
        delimiter = design_array(i, 1)
        For J = 0 To MAX_FIELDS - 1
            fields(J) = def_settings(i, J)
        Next J
    End If
        ' done filling the fields array
    Call write_def_for_export
    
   
    Call do_new(exportfile, chemcount, chemlist, caslist, fields, delimiter)
done_export:

   
End Sub

Public Sub start_append_export(exportfile As String)
 
    Dim fields(MAX_FIELDS) As Integer
    
    Dim delimiter As String
    Dim i As Integer
    Dim J As Integer
    Dim chemcount As Integer
    ReDim chemlist(FRMExport!chemlst.SelCount) As String
    Dim success As Boolean
    Dim found As Boolean
    ReDim caslist(FRMExport!chemlst.SelCount) As Long
        ' get the list of chemicals to be exported
    chemcount = 0
    For i = 0 To FRMExport!chemlst.ListCount - 1
        If FRMExport!chemlst.Selected(i) = True Then
            chemlist(chemcount) = Trim(Left(FRMExport!chemlst.List(i), Len(FRMExport!chemlst.List(i)) - 9))
            caslist(chemcount) = CLng(Trim(Right(FRMExport!chemlst.List(i), 9)))
            chemcount = chemcount + 1
        End If
    Next i
    ' If it's s hard-coded format, do that
    If Trim(FRMExport!nametbx.Text) = "ASPEN" Then
        Call do_aspen_export(exportfile, chemcount, chemlist, caslist, 2) ' the 2 is for append
        GoTo done_export
    End If
    ' else fill the array of fields we're exporting
    found = False
    For i = 0 To FRMExport!designlst.ListCount - 1
        If FRMExport!nametbx.Text = FRMExport!designlst.List(i) Then
            found = True
            Exit For
        End If
    Next i
    If found = False Then
        Call fill_fields(fields)
        If FRMExport!spaceopt.value = True Then
            delimiter = "space"
        Else
            delimiter = "comma"
        End If
    Else
        delimiter = design_array(i, 1)
        For J = 0 To MAX_FIELDS - 1
            fields(J) = def_settings(i, J)
        Next J
    End If
        ' done filling the fields array

    success = False
    success = do_append(exportfile, chemcount, chemlist, caslist, fields, delimiter)
    If success = False Then
        MsgBox ("Export error:  The fields you have selected for export don't match the existing export file")
    End If
done_export:
End Sub

 


