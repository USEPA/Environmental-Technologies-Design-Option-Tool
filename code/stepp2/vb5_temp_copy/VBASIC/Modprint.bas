Attribute VB_Name = "modprint"

Option Explicit
Public Function export_print_info() As Boolean

    ' export cur_info to printtable for use by rpt
    ' this function assumes that cur_info has already been calculated for the
    ' selected chemical.
    
    Dim i As Integer
    Dim J As Integer
    Dim DBTbl As Recordset
    Dim TempAntInfo As AntoineInfoType
    
        ' below we need to check whether we're entering a new
        ' chemical to the table or editing an existing one
    On Error GoTo DB_Closed_Error
    Set DBTbl = DBJetUser.OpenRecordset("PrintTable2", dbOpenTable)
    On Error GoTo DB_Open_Error
   
    DBTbl.Index = "PrimaryKey"
    DBTbl.Seek "=", CStr(Cur_Info.CAS)
    
        If DBTbl.NoMatch Then
            
            For i = 0 To NumProperties
                For J = 1 To NumMethods
'mrt- we can't update antoine here, see note in export_custom_info
                    If J = InfoMethod(i).CurMethod Then
                        DBTbl.AddNew
                        On Error Resume Next
                        DBTbl("CAS") = CStr(Cur_Info.CAS)
                        DBTbl("Name") = Cur_Info.name
                        DBTbl("Property Name") = GetPropName(i)
                        DBTbl("Property Number") = i
                        DBTbl("Property Group") = GetPropGroup(i)
                        DBTbl("Method Name") = InfoMethod(i).MethodName(J)
                        DBTbl("Method Number") = J
                        DBTbl("Method Enabled") = InfoMethod(i).Enabled(J)
                        DBTbl("Current Method Index") = InfoMethod(i).CurMethod
                        DBTbl("Value") = FormatVal(InfoMethod(i).value(J))
                        DBTbl("Unit") = InfoMethod(i).Unit
                        DBTbl("TFT") = FormatVal(InfoMethod(i).TFT)
                        DBTbl("TFTUnit") = InfoMethod(i).TFTUnit
                        DBTbl("OpT") = FormatVal(Cur_Info.OpT)
                        DBTbl("OpP") = FormatVal(Cur_Info.OpP)
                        DBTbl("OpTUnit") = Cur_Info.OpTUnit
                        DBTbl("OpPUnit") = Cur_Info.OpPUnit
                        DBTbl("EqNum") = CStr(InfoMethod(i).EqNum(J))
                        DBTbl("Coeff A") = FormatVal(InfoMethod(i).Coeff(J, 1))
                        DBTbl("Coeff B") = FormatVal(InfoMethod(i).Coeff(J, 2))
                        DBTbl("Coeff C") = FormatVal(InfoMethod(i).Coeff(J, 3))
                        DBTbl("Coeff D") = FormatVal(InfoMethod(i).Coeff(J, 4))
                        DBTbl("Coeff E") = FormatVal(InfoMethod(i).Coeff(J, 5))
                        DBTbl("MinT") = FormatVal(InfoMethod(i).MinT(J))
                        DBTbl("MaxT") = FormatVal(InfoMethod(i).MaxT(J))
                        DBTbl.Update
                    End If
                Next J
            Next i
            
            If Antoine_Info.AntCalc Then
'mrt- special case for antoine. This stems from antoine's "special" status
                DBTbl.AddNew
                On Error Resume Next
                DBTbl("CAS") = CStr(Cur_Info.CAS)
                DBTbl("Name") = Cur_Info.name
                DBTbl("Property Name") = "Antoine"
                DBTbl("Property Number") = ANT
                DBTbl("Property Group") = 7 'It has its own group?
                DBTbl("Method Name") = Antoine_Info.MethodName
                DBTbl("Method Number") = 0
                DBTbl("Method Enabled") = True
                DBTbl("Current Method Index") = 0
                DBTbl("Value") = Antoine_Info.value
                DBTbl("Unit") = Antoine_Info.Unit
                DBTbl("TFT") = Antoine_Info.TFT
                DBTbl("TFTUnit") = Antoine_Info.TFTUnit
                DBTbl("OpT") = "0"
                DBTbl("OpP") = "0"
                DBTbl("OpTUnit") = "None"
                DBTbl("OpPUnit") = "None"
                DBTbl("EqNum") = Antoine_Info.EqNum
                DBTbl("Coeff A") = Antoine_Info.A
                DBTbl("Coeff B") = Antoine_Info.B
                DBTbl("Coeff C") = Antoine_Info.C
                DBTbl("Coeff D") = Antoine_Info.D
                DBTbl("Coeff E") = Antoine_Info.E
                DBTbl("MinT") = Antoine_Info.TMin
                DBTbl("MaxT") = Antoine_Info.TMax
                DBTbl.Update
            Else
                
                TempAntInfo.MethodName = "Antoine"
                Call run_default_ant_calc
                Call antoine_check_update_udb(TempAntInfo)
                
                DBTbl.AddNew
                On Error Resume Next
                DBTbl("CAS") = CStr(Cur_Info.CAS)
                DBTbl("Name") = Cur_Info.name
                DBTbl("Property Name") = "Antoine"
                DBTbl("Property Number") = ANT
                DBTbl("Property Group") = 7 'It has its own group?
                DBTbl("Method Name") = TempAntInfo.MethodName
                DBTbl("Method Number") = 0
                DBTbl("Method Enabled") = True
                DBTbl("Current Method Index") = 0
                DBTbl("Value") = TempAntInfo.value
                DBTbl("Unit") = TempAntInfo.Unit
                DBTbl("TFT") = TempAntInfo.TFT
                DBTbl("TFTUnit") = TempAntInfo.TFTUnit
                DBTbl("OpT") = "0"
                DBTbl("OpP") = "0"
                DBTbl("OpTUnit") = "None"
                DBTbl("OpPUnit") = "None"
                DBTbl("EqNum") = TempAntInfo.EqNum
                DBTbl("Coeff A") = TempAntInfo.A
                DBTbl("Coeff B") = TempAntInfo.B
                DBTbl("Coeff C") = TempAntInfo.C
                DBTbl("Coeff D") = TempAntInfo.D
                DBTbl("Coeff E") = TempAntInfo.E
                DBTbl("MinT") = TempAntInfo.TMin
                DBTbl("MaxT") = TempAntInfo.TMax
                DBTbl.Update
            End If
        Else
            For i = 0 To NumProperties
                For J = 1 To NumMethods
                    If J = InfoMethod(i).CurMethod Then
                        DBTbl.Edit
                        On Error Resume Next
                        DBTbl("CAS") = CStr(Cur_Info.CAS)
                        DBTbl("Name") = Cur_Info.name
                        DBTbl("Property Name") = GetPropName(i)
                        DBTbl("Property Number") = i
                        DBTbl("Property Group") = GetPropGroup(i)
                        DBTbl("Method Name") = InfoMethod(i).MethodName(J)
                        DBTbl("Method Number") = J
                        DBTbl("Method Enabled") = InfoMethod(i).Enabled(J)
                        DBTbl("Value") = FormatVal(InfoMethod(i).value(J))
                        DBTbl("Unit") = InfoMethod(i).Unit
                        DBTbl("TFT") = FormatVal(InfoMethod(i).TFT)
                        DBTbl("TFTUnit") = InfoMethod(i).TFTUnit
                        DBTbl("OpT") = FormatVal(Cur_Info.OpT)
                        DBTbl("OpP") = FormatVal(Cur_Info.OpP)
                        DBTbl("OpTUnit") = Cur_Info.OpTUnit
                        DBTbl("OpPUnit") = Cur_Info.OpPUnit
                        DBTbl("EqNum") = CStr(InfoMethod(i).EqNum(J))
                        DBTbl("Coeff A") = FormatVal(InfoMethod(i).Coeff(J, 1))
                        DBTbl("Coeff B") = FormatVal(InfoMethod(i).Coeff(J, 2))
                        DBTbl("Coeff C") = FormatVal(InfoMethod(i).Coeff(J, 3))
                        DBTbl("Coeff D") = FormatVal(InfoMethod(i).Coeff(J, 4))
                        DBTbl("Coeff E") = FormatVal(InfoMethod(i).Coeff(J, 5))
                        DBTbl("MinT") = FormatVal(InfoMethod(i).MinT(J))
                        DBTbl("MaxT") = FormatVal(InfoMethod(i).MaxT(J))
                        DBTbl.Update
                    End If
                Next J
            Next i
            
            If Antoine_Info.AntCalc Then
'mrt- special case for antoine. This stems from antoine's "special" status
                DBTbl.Edit
                On Error Resume Next
                DBTbl("CAS") = CStr(Cur_Info.CAS)
                DBTbl("Name") = Cur_Info.name
                DBTbl("Property Name") = "Antoine"
                DBTbl("Property Number") = ANT
                DBTbl("Property Group") = 7 'It has its own group?
                DBTbl("Method Name") = Antoine_Info.MethodName
                DBTbl("Method Number") = 0
                DBTbl("Method Enabled") = False
                DBTbl("Current Method Index") = 0
                DBTbl("Value") = Antoine_Info.value
                DBTbl("Unit") = Antoine_Info.Unit
                DBTbl("TFT") = Antoine_Info.TFT
                DBTbl("TFTUnit") = Antoine_Info.TFTUnit
                DBTbl("OpT") = "0"
                DBTbl("OpP") = "0"
                DBTbl("OpTUnit") = "None"
                DBTbl("OpPUnit") = "None"
                DBTbl("EqNum") = Antoine_Info.EqNum
                DBTbl("Coeff A") = Antoine_Info.A
                DBTbl("Coeff B") = Antoine_Info.B
                DBTbl("Coeff C") = Antoine_Info.C
                DBTbl("Coeff D") = Antoine_Info.D
                DBTbl("Coeff E") = Antoine_Info.E
                DBTbl("MinT") = Antoine_Info.TMin
                DBTbl("MaxT") = Antoine_Info.TMax
                DBTbl.Update
            Else
                
                TempAntInfo.MethodName = "Antoine"
                Call run_default_ant_calc
                Call antoine_check_update_udb(TempAntInfo)
                
                DBTbl.Edit
                On Error Resume Next
                DBTbl("CAS") = CStr(Cur_Info.CAS)
                DBTbl("Name") = Cur_Info.name
                DBTbl("Property Name") = "Antoine"
                DBTbl("Property Number") = ANT
                DBTbl("Property Group") = 7 'It has its own group?
                DBTbl("Method Name") = TempAntInfo.MethodName
                DBTbl("Method Number") = 0
                DBTbl("Method Enabled") = True
                DBTbl("Current Method Index") = 0
                DBTbl("Value") = TempAntInfo.value
                DBTbl("Unit") = TempAntInfo.Unit
                DBTbl("TFT") = TempAntInfo.TFT
                DBTbl("TFTUnit") = TempAntInfo.TFTUnit
                DBTbl("OpT") = "0"
                DBTbl("OpP") = "0"
                DBTbl("OpTUnit") = "None"
                DBTbl("OpPUnit") = "None"
                DBTbl("EqNum") = TempAntInfo.EqNum
                DBTbl("Coeff A") = TempAntInfo.A
                DBTbl("Coeff B") = TempAntInfo.B
                DBTbl("Coeff C") = TempAntInfo.C
                DBTbl("Coeff D") = TempAntInfo.D
                DBTbl("Coeff E") = TempAntInfo.E
                DBTbl("MinT") = TempAntInfo.TMin
                DBTbl("MaxT") = TempAntInfo.TMax
                DBTbl.Update
            End If
        End If
        DBTbl.Close
        export_print_info = True
        Exit Function
        

DB_Open_Error:
    
    MsgBox "Error saving data to user database", 48, "Error"
    DBTbl.Close
    export_print_info = False
    Exit Function
    
DB_Closed_Error:
    MsgBox "Can't find user database"
    export_print_info = False
    Exit Function

End Function

Public Function export_custom_info(caslist() As String, num_chems As Integer) As Boolean

    ' this function exports each of the chemicals to the table
    ' where the report can find them
    Dim i As Integer
    Dim J As Integer
    Dim N As Integer
    Dim DBTbl As Recordset
    Dim found As Boolean
    
    Dim TempAntInfo As AntoineInfoType
        
        ' set the error conditions below depending on whether the
        ' database is open or not
        
    On Error GoTo DB_Closed_Error
    Set DBTbl = DBJetUser.OpenRecordset("PrintTable2", dbOpenTable)
    On Error GoTo DB_Open_Error
    
    For N = 0 To num_chems - 1
        Cur_Info.CAS = CLng(caslist(N))
        If Cur_Info.CAS = 0 Then
            GoTo BadCasError
        End If
        Call Recalculate
        DBTbl.Index = "PrimaryKey"
    
    'look for existing entries
        DBTbl.Seek "=", CStr(Cur_Info.CAS)
    
    'if there are no existing entries
        If DBTbl.NoMatch Then
            For i = 0 To NumProperties
                found = False
                For J = 1 To NumMethods
'mrt- because of problems in the numbering system, this update can't be done
'       for the antoine coefficients.
                    If J = InfoMethod(i).CurMethod Then
                        found = True
                        DBTbl.AddNew
                        On Error Resume Next
                        DBTbl("CAS") = CStr(Cur_Info.CAS)
                        DBTbl("Name") = Cur_Info.name
                        DBTbl("Property Name") = GetPropName(i)
                        DBTbl("Property Number") = i
                        DBTbl("Property Group") = GetPropGroup(i)
                        DBTbl("Method Name") = InfoMethod(i).MethodName(J)
                        DBTbl("Method Number") = J
                        DBTbl("Method Enabled") = InfoMethod(i).Enabled(J)
                        DBTbl("Current Method Index") = InfoMethod(i).CurMethod
                        DBTbl("Value") = FormatVal(InfoMethod(i).value(J))
                        DBTbl("Unit") = InfoMethod(i).Unit
                        DBTbl("TFT") = FormatVal(InfoMethod(i).TFT)
                        DBTbl("TFTUnit") = InfoMethod(i).TFTUnit
                        DBTbl("OpT") = FormatVal(Cur_Info.OpT)
                        DBTbl("OpP") = FormatVal(Cur_Info.OpP)
                        DBTbl("OpTUnit") = Cur_Info.OpTUnit
                        DBTbl("OpPUnit") = Cur_Info.OpPUnit
                        DBTbl("EqNum") = CStr(InfoMethod(i).EqNum(J))
                        DBTbl("Coeff A") = FormatVal(InfoMethod(i).Coeff(J, 1))
                        DBTbl("Coeff B") = FormatVal(InfoMethod(i).Coeff(J, 2))
                        DBTbl("Coeff C") = FormatVal(InfoMethod(i).Coeff(J, 3))
                        DBTbl("Coeff D") = FormatVal(InfoMethod(i).Coeff(J, 4))
                        DBTbl("Coeff E") = FormatVal(InfoMethod(i).Coeff(J, 5))
                        DBTbl("MinT") = FormatVal(InfoMethod(i).MinT(J))
                        DBTbl("MaxT") = FormatVal(InfoMethod(i).MaxT(J))
                        DBTbl.Update
                        
                    End If
                    If found = True Then
                        Exit For
                    End If
                Next J
            Next i

'mrt- special case for antoine. This stems from antoine's "special" status
            
            TempAntInfo.MethodName = "Antoine"
            Call run_default_ant_calc
            Call antoine_check_update_udb(TempAntInfo)
            
            DBTbl.AddNew
            On Error GoTo 0
            DBTbl("CAS") = CStr(Cur_Info.CAS)
            DBTbl("Name") = Cur_Info.name
            DBTbl("Property Name") = "Antoine"
            DBTbl("Property Number") = ANT
            DBTbl("Property Group") = 7 'It has its own group?
            DBTbl("Method Name") = TempAntInfo.MethodName
            DBTbl("Method Number") = 0
            DBTbl("Method Enabled") = True
            DBTbl("Current Method Index") = 0
            DBTbl("Value") = TempAntInfo.value
            DBTbl("Unit") = TempAntInfo.Unit
            DBTbl("TFT") = TempAntInfo.TFT
            DBTbl("TFTUnit") = TempAntInfo.TFTUnit
            DBTbl("OpT") = "0"
            DBTbl("OpP") = "0"
            DBTbl("OpTUnit") = "None"
            DBTbl("OpPUnit") = "None"
            DBTbl("EqNum") = TempAntInfo.EqNum
            DBTbl("Coeff A") = TempAntInfo.A
            DBTbl("Coeff B") = TempAntInfo.B
            DBTbl("Coeff C") = TempAntInfo.C
            DBTbl("Coeff D") = TempAntInfo.D
            DBTbl("Coeff E") = TempAntInfo.E
            DBTbl("MinT") = TempAntInfo.TMin
            DBTbl("MaxT") = TempAntInfo.TMax
            DBTbl.Update
            
    'else if there are existing entries
        Else
            For i = 0 To NumProperties
                found = False
                For J = 1 To NumMethods
                    If J = InfoMethod(i).CurMethod Then
                        found = True
                        DBTbl.Edit
                        On Error Resume Next
                        DBTbl("CAS") = CStr(Cur_Info.CAS)
                        DBTbl("Name") = Cur_Info.name
                        DBTbl("Property Name") = GetPropName(i)
                        DBTbl("Property Number") = i
                        DBTbl("Property Group") = GetPropGroup(i)
                        DBTbl("Method Name") = InfoMethod(i).MethodName(J)
                        DBTbl("Method Number") = J
                        DBTbl("Method Enabled") = InfoMethod(i).Enabled(J)
                        DBTbl("Value") = FormatVal(InfoMethod(i).value(J))
                        DBTbl("Unit") = InfoMethod(i).Unit
                        DBTbl("TFT") = FormatVal(InfoMethod(i).TFT)
                        DBTbl("TFTUnit") = InfoMethod(i).TFTUnit
                        DBTbl("OpT") = FormatVal(Cur_Info.OpT)
                        DBTbl("OpP") = FormatVal(Cur_Info.OpP)
                        DBTbl("OpTUnit") = Cur_Info.OpTUnit
                        DBTbl("OpPUnit") = Cur_Info.OpPUnit
                        DBTbl("EqNum") = CStr(InfoMethod(i).EqNum(J))
                        DBTbl("Coeff A") = FormatVal(InfoMethod(i).Coeff(J, 1))
                        DBTbl("Coeff B") = FormatVal(InfoMethod(i).Coeff(J, 2))
                        DBTbl("Coeff C") = FormatVal(InfoMethod(i).Coeff(J, 3))
                        DBTbl("Coeff D") = FormatVal(InfoMethod(i).Coeff(J, 4))
                        DBTbl("Coeff E") = FormatVal(InfoMethod(i).Coeff(J, 5))
                        DBTbl("MinT") = FormatVal(InfoMethod(i).MinT(J))
                        DBTbl("MaxT") = FormatVal(InfoMethod(i).MaxT(J))
                        DBTbl.Update
                        
                    End If
                    If (found = True) Then
                        Exit For
                    End If
                 Next J
            Next i
            
            TempAntInfo.MethodName = "Antoine"
            Call run_default_ant_calc
            Call antoine_check_update_udb(TempAntInfo)
            
            DBTbl.Edit
            On Error Resume Next
            DBTbl("CAS") = CStr(Cur_Info.CAS)
            DBTbl("Name") = Cur_Info.name
            DBTbl("Property Name") = "Antoine"
            DBTbl("Property Number") = ANT
            DBTbl("Property Group") = 7 'It has its own group?
            DBTbl("Method Name") = TempAntInfo.MethodName
            DBTbl("Method Number") = 0
            DBTbl("Method Enabled") = True
            DBTbl("Current Method Index") = 0
            DBTbl("Value") = TempAntInfo.value
            DBTbl("Unit") = TempAntInfo.Unit
            DBTbl("TFT") = TempAntInfo.TFT
            DBTbl("TFTUnit") = TempAntInfo.TFTUnit
            DBTbl("OpT") = "0"
            DBTbl("OpP") = "0"
            DBTbl("OpTUnit") = "None"
            DBTbl("OpPUnit") = "None"
            DBTbl("EqNum") = TempAntInfo.EqNum
            DBTbl("Coeff A") = TempAntInfo.A
            DBTbl("Coeff B") = TempAntInfo.B
            DBTbl("Coeff C") = TempAntInfo.C
            DBTbl("Coeff D") = TempAntInfo.D
            DBTbl("Coeff E") = TempAntInfo.E
            DBTbl("MinT") = TempAntInfo.TMin
            DBTbl("MaxT") = TempAntInfo.TMax
            DBTbl.Update
        End If
endloop:
            
        Next N
        DBTbl.Close
        export_custom_info = True
        Exit Function
   

BadCasError:
    Resume endloop
    
DB_Open_Error:
    
    MsgBox "Error saving data to user database", 48, "Error"
    DBTbl.Close
    export_custom_info = False
    Exit Function

DB_Closed_Error:
    MsgBox "Can't find user database"
    export_custom_info = False
    Exit Function

End Function

Public Sub Clear_Print_Table()

    Dim DBTbl As Recordset
    
    On Error GoTo DB_Closed_Error
    Set DBTbl = DBJetUser.OpenRecordset("PrintTable2", dbOpenTable)
    On Error GoTo DB_Open_Error
    DBTbl.MoveFirst
    On Error Resume Next
    While DBTbl.EOF = False
            DBTbl.Delete
            DBTbl.MoveNext
    Wend
    DBTbl.Close
    Exit Sub
DB_Open_Error:
        DBTbl.Close
        Exit Sub
DB_Closed_Error:
        Exit Sub
End Sub
