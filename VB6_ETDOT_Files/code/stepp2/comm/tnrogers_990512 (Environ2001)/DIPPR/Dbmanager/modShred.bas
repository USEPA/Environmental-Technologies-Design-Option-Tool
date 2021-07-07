Attribute VB_Name = "modShred"
Option Explicit

Public Function Run_Mosdap(ByVal Smiles As String, Dat_File As String, Search_Type As Integer) As Byte
'        Arguments:
'            Smiles -> the smiles string we're disassembling
'            Dat_file -> the file containing the groups for input
'            Search_type -> the type of search
'                    0: Sequential, Non-Truncating
'                    1: Sequential, Truncating
'                    2: Combinatorial, Truncating
'
'        Return Value:
'            0: Unable to disassemble the given smiles string
'                    or Error occured in funtion
'            1: Successfully disassembled
'            2: Partially disassembled

    Dim SearchResult As Byte
    Dim local_file As String
    Dim i As Integer
    Dim intSF_ID(0 To 99) As Long, intSF_Quant(0 To 99) As Long
    Dim intMF_ID(0 To 20) As Long, intMF_Quant(0 To 20) As Long
        
    Screen.MousePointer = 11
    Run_Mosdap = 0
    
    local_file = AppPath() & "dat\" & Dat_File ' the file we're reading groups from (ie unifac.dat)
    
    If Not FileExists(local_file) Then
        MsgBox "Mosdap dat file '" & local_file & "' : doesn't exist"
        GoTo Exit_Function
    End If
    
    Call MOSDAP(Smiles, 0, local_file, "", Search_Type, SearchResult, intSF_ID(0), intSF_Quant(0), intMF_ID(0), intMF_Quant(0))
    
    For i = 0 To 21
        If intSF_ID(i) > 0 And SearchResult <> 0 Then
            cur_chem_groups(i) = intSF_ID(i)
            num_cur_chem_groups(i) = intSF_Quant(i)
        Else
            cur_chem_groups(i) = 0
            num_cur_chem_groups(i) = 0
        End If
    Next i
    
    Run_Mosdap = SearchResult
    
Exit_Function:
    Screen.MousePointer = 1

End Function

Public Function Calc_Ambrose(Value As Double, Units As String, ByVal Cur_Property As String) As Boolean
' Paul ... done ... Check on 4/10/99
Dim sql As String
Dim DB As Database
Dim RS As Recordset
Dim Tc As Double
Dim Vc As Double
Dim Pc As Double
Dim NBP As Double
Dim MW As Double
Dim i As Integer

    Calc_Ambrose = False
    Set DB = OpenDatabase(AppPath & "Methods.mdb", False, False)
    For i = 0 To 21
        If num_cur_chem_groups(i) > 0 Then
            sql = "select * from [Joback] where [Mosdap ID] = " & cur_chem_groups(i)
            Set RS = DB.OpenRecordset(sql, dbOpenSnapshot)
            If Not IsNull(RS("Tc")) And Cur_Property = "Critical Temperature" Then
                Tc = Tc + num_cur_chem_groups(i) * RS("Tc")
            ElseIf Not IsNull(RS("Pc")) And Cur_Property = "Critical Pressure" Then
                Pc = Pc + num_cur_chem_groups(i) * RS("Pc")
            ElseIf Not IsNull(RS("Vc")) And Cur_Property = "Critical Volume" Then
                Vc = Vc + num_cur_chem_groups(i) * RS("Vc")
            Else
                RS.Close
                DB.Close
                Exit Function
            End If
            RS.Close
        Else
            Exit For
        End If
    Next
    DB.Close
    
    Select Case Cur_Property
        Case Cur_Property = "Critical Temperature"
'            T = Tb(1 + (1.242 + sumTc) ^ -1)
            If GetMasterPropInfo(NBP, Units, "NBP", "3e", selected_name) = False Then
                Exit Function
            End If
            Value = NBP * (1 + (1.242 + Tc) ^ 1)
            Units = "K"
        Case Cur_Property = "Critical Pressure"
'            P = M(0.339 + sumPc) ^ -2
            If GetMasterPropInfo(MW, Units, "MW", "3a", selected_name) = False Then
                Exit Function
            End If
            Value = MW * (0.339 + Pc) ^ -2
            Units = "bar"
        Case Cur_Property = "Critical Volume"
'            V = 40 + sumVc
            Value = 40 + Vc
            Units = "cm3/mol"
        Case Else
            Exit Function
    End Select
    Calc_Ambrose = True
End Function

Public Function Calc_Joback(Value As Double, Units As String, Cur_Property As String) As Boolean
' Paul ... done ... checked on 3/30/99
Dim sql As String
Dim DB As Database
Dim RS As Recordset
Dim Tc As Double
Dim Vc As Double
Dim Pc As Double
Dim Tb As Double
Dim Tf As Double
Dim Na As Integer
Dim i As Integer

    Calc_Joback = False
    Set DB = OpenDatabase(AppPath & "Methods.mdb", False, False)
    For i = 0 To 21
        If num_cur_chem_groups(i) > 0 Then
            sql = "select * from [Joback] where [Mosdap ID] = " & cur_chem_groups(i)
            Set RS = DB.OpenRecordset(sql, dbOpenSnapshot)
            If Not IsNull(RS("Tc")) And Not IsNull(RS("Tb")) And Cur_Property = "Critical Temperature" Then
                Tc = Tc + num_cur_chem_groups(i) * RS("Tc")
                Tb = Tb + num_cur_chem_groups(i) * RS("Tb")
            ElseIf Not IsNull(RS("Pc")) And Not IsNull(RS("Na")) And Cur_Property = "Critical Pressure" Then
                Pc = Pc + num_cur_chem_groups(i) * RS("Pc")
                Na = Na + num_cur_chem_groups(i) * RS("Na")
            ElseIf Not IsNull(RS("Vc")) And Cur_Property = "Critical Volume" Then
                Vc = Vc + num_cur_chem_groups(i) * RS("Vc")
            ElseIf Cur_Property = "Freezing/Melting Point" Then
            
            ElseIf Not IsNull(RS("Tb")) And Cur_Property = "Normal Boiling Point" Then
                Tb = Tb + num_cur_chem_groups(i) * RS("Tb")
            ElseIf Not IsNull(RS("Tf")) And Cur_Property = "Normal Freezing Point" Then
                Tf = Tf + num_cur_chem_groups(i) * RS("Tf")
            Else
                RS.Close
                DB.Close
                Exit Function
            End If
            RS.Close
        Else
            Exit For
        End If
    Next
    DB.Close
    
    Units = "K"
    Select Case Cur_Property
        Case "Critical Temperature"
            Tb = 198 + Tb
            Value = Tb * (0.584 + (0.965 * Tc) - Tc ^ 2) ^ -1
        Case "Critical Pressure"
            Value = (0.113 + (0.0032 * Na) - Pc) ^ -2
            Units = "bar"
        Case "Critical Volume"
            Value = 17.5 + Vc
            Units = "cm3/mol"
        Case "Freezing/Melting Point"
        
        Case "Normal Boiling Point"
            Value = 198 + Tb
        Case "Normal Freezing Point"
            Value = 122 + Tf
        Case Else
            Exit Function
    End Select
    
    Calc_Joback = True
End Function

Public Function Calc_Fedors(Value As Double, Units As String, ByVal Cur_Property As String) As Boolean
' Paul ... done ... Check on 3/30/99
Dim sql As String
Dim DB As Database
Dim RS As Recordset
Dim Tc As Double
Dim i As Integer
    
    Calc_Fedors = False
    Set DB = OpenDatabase(AppPath & "Methods.mdb", False, False)
    For i = 0 To 21
        If num_cur_chem_groups(i) > 0 Then
            sql = "select * from [Fedors] where [Mosdap ID] = " & cur_chem_groups(i)
            Set RS = DB.OpenRecordset(sql, dbOpenSnapshot)
            If Not IsNull(RS("Tc")) And Cur_Property = "Critical Temperature" Then
                Tc = Tc + num_cur_chem_groups(i) * RS("Tc")
            Else
                RS.Close
                DB.Close
                Exit Function
            End If
            RS.Close
        Else
            Exit For
        End If
    Next
    DB.Close

    Units = "K"
    Select Case Cur_Property
        Case "Critical Temperature"
'            T = 535 * Log(sumTc)
            Value = 535 * Log(Tc)
        Case Else
            Exit Function
    End Select
    
    Calc_Fedors = True
End Function

Public Function Calc_Hine_and_Mookerjee(Value As Double, Units As String, ByVal Cur_Property As String) As Boolean
' Paul ... done ... Check on 3/30/99
Dim sql As String
Dim DB As Database
Dim RS As Recordset
Dim HC As Double
Dim i As Integer

    Calc_Hine_and_Mookerjee = False
    Set DB = OpenDatabase(AppPath & "Methods.mdb", False, False)
    For i = 0 To 21
        If num_cur_chem_groups(i) > 0 Then
            sql = "select * from [Hine & Mookerjee] where [Mosdap ID] = " & cur_chem_groups(i)
            Set RS = DB.OpenRecordset(sql, dbOpenSnapshot)
            If Not IsNull(RS("Hine & Mookerjee")) And Cur_Property = "Henry's Constant" Then
                HC = HC + num_cur_chem_groups(i) * RS("Hine & Mookerjee")
            Else
                RS.Close
                DB.Close
                Exit Function
            End If
            RS.Close
        Else
            Exit For
        End If
    Next
    DB.Close
    
    Units = "unitless"
    Select Case Cur_Property
        Case "Henry's Constant"
            Value = 10 ^ (-HC)
        Case Else
            Exit Function
    End Select
    
    Calc_Hine_and_Mookerjee = True
End Function

Public Function Calc_Lebas(Value As Double, Units As String, ByVal Cur_Property As String) As Boolean
' Paul ... done ... Check on 3/30/99
Dim sql As String
Dim DB As Database
Dim RS As Recordset
Dim MV As Double
Dim i As Integer

    Calc_Lebas = False
    Set DB = OpenDatabase(AppPath & "Methods.mdb", False, False)
    For i = 0 To 21
        If num_cur_chem_groups(i) > 0 Then
            sql = "select * from [LeBas] where [Mosdap ID] = " & cur_chem_groups(i)
            Set RS = DB.OpenRecordset(sql, dbOpenSnapshot)
            If Not IsNull(RS("LeBas")) And Cur_Property = "Molar Volume" Then
                MV = MV + num_cur_chem_groups(i) * RS("LeBas")
            Else
                RS.Close
                DB.Close
                Exit Function
            End If
            RS.Close
        Else
            Exit For
        End If
    Next
    DB.Close
    
    Units = "cm3/mol"
    Select Case Cur_Property
        Case "Molar Volume"
            Value = MV
        Case Else
            Exit Function
    End Select
    
    Calc_Lebas = True
End Function

Public Function Calc_Lyderson(Value As Double, Units As String, ByVal Cur_Property As String) As Boolean
'9999 not finished
Dim sql As String
Dim DB As Database
Dim RS As Recordset
Dim Tc As Double
Dim i As Integer

    Calc_Lyderson = False
    Set DB = OpenDatabase(AppPath & "Methods.mdb", False, False)
    For i = 0 To 21
        If num_cur_chem_groups(i) > 0 Then
            sql = "select * from [Lyderson] where [Mosdap ID] = " & cur_chem_groups(i)
            Set RS = DB.OpenRecordset(sql, dbOpenSnapshot)
            If Not IsNull(RS("Tc")) And Cur_Property = "Critical Temperature" Then
                Tc = Tc + num_cur_chem_groups(i) * RS("Tc")
            Else
                RS.Close
                DB.Close
                Exit Function
            End If
            RS.Close
        Else
            Exit For
        End If
    Next
    DB.Close
    
'    pos = 1
'    While pos <> 0
'        pos = InStr(pos, LCase(selected_smiles), "c")
'        If pos <> 0 Then
'            Carbon_Count = Carbon_Count + 1
'        End If
'    Loop
'
'    a =
'
'    Tc = (8 * A) / (27 * B * 82.05)
''    Tc = (8 * a) / (27 * b * 82.05)
'    Units = "K"
'
'    Calc_Lyderson = True
End Function

Public Function Calc_Pintar(Value As Double, Units As String, ByVal Cur_Property As String, ByVal Cur_Method As String) As Boolean
' Paul ... done ... Check on 3/30/99
Dim sql As String
Dim DB As Database
Dim RS As Recordset
Dim Hi As Double
Dim Ai As Double
Dim Bi As Double
Dim i As Integer

    Calc_Pintar = False
    Set DB = OpenDatabase(AppPath & "Methods.mdb", False, False)
    For i = 0 To 21
        If num_cur_chem_groups(i) > 0 Then
            sql = "select * from [Pintar] where [Mosdap ID] = " & cur_chem_groups(i)
            Set RS = DB.OpenRecordset(sql, dbOpenSnapshot)
            If Not IsNull(RS("Hi")) And Cur_Method = "MTU Logarithmic Groups (Pintar)" And Cur_Property = "Auto-Ignition T (AIT)" Then
                Hi = Hi + num_cur_chem_groups(i) * RS("Hi")
            ElseIf Not IsNull(RS("Ai")) And Cur_Method = "MTU Logarithmic Groups (Pintar)" And Cur_Property = "Auto-Ignition T (AIT)" Then
                Ai = Ai + num_cur_chem_groups(i) * RS("Ai")
            ElseIf Not IsNull(RS("Bi")) And Cur_Method = "MTU Linear Groups (Pintar)" And Cur_Property = "Auto-Ignition T (AIT)" Then
                Bi = Bi + num_cur_chem_groups(i) * RS("Bi")
            Else
                RS.Close
                DB.Close
                Exit Function
            End If
            RS.Close
        Else
            Exit For
        End If
    Next
    DB.Close
    
    Units = "K"
    Select Case Cur_Method
        Case "MTU Logarithmic Groups (Pintar)"
            Value = 1500 * ((1 + Ai) / (ln(Hi)))
        Case "MTU Linear Groups (Pintar)"
            Value = Bi
        Case Else
            Exit Function
    End Select
    
    Calc_Pintar = True
End Function

Public Function Calc_Reichenberg(Value As Double, Units As String, ByVal Cur_Property As String) As Boolean
' Paul ... done ... Check on 4/10/99
Dim sql As String
Dim DB As Database
Dim RS As Recordset
Dim Ci As Double
Dim MW As Double
Dim T As Double
Dim Tr As Double
Dim Tc As Double
Dim A As Double
Dim i As Integer

    Calc_Reichenberg = False
    Set DB = OpenDatabase(AppPath & "Methods.mdb", False, False)
    For i = 0 To 21
        If num_cur_chem_groups(i) > 0 Then
            sql = "select * from [Pintar] where [Mosdap ID] = " & cur_chem_groups(i)
            Set RS = DB.OpenRecordset(sql, dbOpenSnapshot)
            If Not IsNull(RS("Ci")) Then
                Ci = Ci + num_cur_chem_groups(1) * RS("Ci") * 1000000000#
            Else
                RS.Close
                DB.Close
                Exit Function
            End If
            RS.Close
        Else
            Exit For
        End If
    Next
    DB.Close
    
    If GetMasterPropInfo(Tc, Units, "TC", "3q", selected_name) = False And GetMasterPropInfo(MW, Units, "MW", "3a", selected_name) = False Then
        Exit Function
    End If
    
    T = 25
    A = (MW ^ (1 / 2) * Tc) / Ci
    Tr = T / Tc
    
    Value = (A * Tr) / ((1 + (0.36 * Tr * (Tr - 1))) ^ (1 / 6))
    Units = "(Pa-s)"
    
    Calc_Reichenberg = True
End Function

Public Function Calc_Unifac(Value As Double, Units As String, ByVal OperatingTemp As Double, ByVal Cur_Property As String, Optional MW As Double) As Boolean
' Paul ... done ... Check on 3/30/99
Dim Unifac_Short As Long
Dim Unifac_Long As Long
Dim Unifac_Error As Long
Dim Unifac_Temp As Double
Dim FGRPError As Long
Dim MaxGroups As Long
Dim MS(1 To 10, 1 To 10, 1 To 2) As Long
Dim BIP_DataBase As Long

Dim ActivityCoefficient As Double
Dim XMW(1 To 2) As Double
Dim Vp As Double
'Dim OperatingTemp As Double
'    Dim MGSG(1 To 116) As Long
'    Dim Ai(1 To 58, 1 To 58) As Double
'    Dim RI(1 To 116) As Double
'    Dim QI(1 To 116) As Double
'    Dim MWS(1 To 116) As Double
'    Dim MVS(1 To 116) As Double
Dim i As Integer
Dim j As Integer
    On Error GoTo Handler
    
    Calc_Unifac = False
    BIP_DataBase = 0
Set_BIP:
    '*********************************************************
    '*        1 = Original UNIFAC VLE (AVLE.DAT)             *
    '*        2 = UNIFAC LLE (ALLE.DAT)                      *
    '*        3 = Environmental VLE (AENV.DAT)               *
    '*********************************************************
    Select Case Cur_Property
        Case "Henry's Constant"
            ActivityCoefficient = 0     'Returned Value
            Vp = 0
            If Calc_Unifac(ActivityCoefficient, Units, OperatingTemp, "Activity coefficient of chemical in water") = False Then
                Exit Function
            ElseIf GetMasterPropInfo(Vp, Units, "", "3f", selected_name) = False Then
                Exit Function
            End If
            Call HC1CALL(Value, Unifac_Short, Unifac_Long, Unifac_Error, Unifac_Temp, _
                        OperatingTemp, ActivityCoefficient, Vp)
            'If Unifac_Error < 0 Then GoTo Set_BIP
            Units = "unitless"
            Calc_Unifac = True
            Exit Function
        Case "Activity coefficient of chemical in water"
            If BIP_DataBase = 0 Then
                BIP_DataBase = 3
            ElseIf BIP_DataBase = 1 Then
                BIP_DataBase = 2
            ElseIf BIP_DataBase = 3 Then
                BIP_DataBase = 1
            Else
                Exit Function
            End If
        Case "Aqueous Solubility"
            If BIP_DataBase = 0 Then
                BIP_DataBase = 2
            ElseIf BIP_DataBase = 1 Then
                BIP_DataBase = 3
            ElseIf BIP_DataBase = 2 Then
                BIP_DataBase = 1
            Else
                Exit Function
            End If
        Case "Log10 Kow"
            If BIP_DataBase = 0 Then
                BIP_DataBase = 2
            ElseIf BIP_DataBase = 2 Then
                BIP_DataBase = 3
            ElseIf BIP_DataBase = 3 Then
                BIP_DataBase = 1
            Else
                Exit Function
            End If
        Case Else
            Exit Function
    End Select
    
    For i = 1 To 10
        For j = 1 To 10
            MS(i, j, 1) = 0
            MS(i, j, 2) = 0
        Next j
    Next i

    For i = 1 To 10
        If num_cur_chem_groups(i - 1) > 0 Then
            MS(2, i, 1) = cur_chem_groups(i - 1)
            MS(2, i, 2) = num_cur_chem_groups(i - 1)
        Else
            Exit For
        End If
    Next i
    MaxGroups = i
    
    Unifac_Short = 0   'Not Important
    Unifac_Long = 0   'Not Important
    Unifac_Error = 0  'Not Important
    Unifac_Temp = 0 'Not Important
    FGRPError = 0 'Not Important

    Select Case Cur_Property
        Case "Activity coefficient of chemical in water"
            Call ACCALL(Value, Unifac_Short, Unifac_Long, Unifac_Error, Unifac_Temp, _
                        OperatingTemp, FGRPError, MaxGroups, MS(1, 1, 1), BIP_DataBase)
            If Unifac_Error < 0 Then GoTo Set_BIP
            Units = "unitless"
        Case "Aqueous Solubility"
            If XMW(1) = 0 Then
                If GetMasterPropInfo(XMW(1), Units, "MW", "3a", "Water") = False Then
                    Exit Function
                End If
            End If
            If XMW(2) = 0 Then
                If IsMissing(MW) Then
                    If GetMasterPropInfo(XMW(2), Units, "MW", "3a", selected_name) = False Then
                        Exit Function
                    End If
                Else
                    XMW(2) = MW
                End If
            End If
            Call AQSCALL(Value, Unifac_Short, Unifac_Long, Unifac_Error, Unifac_Temp, _
                        OperatingTemp, MaxGroups, MS(1, 1, 1), XMW(1), BIP_DataBase)
            If Unifac_Error < 0 Then GoTo Set_BIP
            Units = "ppm(wt)"
        Case "Log10 Kow"
            Call KOWCALL(Value, Unifac_Short, Unifac_Long, Unifac_Error, Unifac_Temp, _
                        OperatingTemp, FGRPError, MaxGroups, MS(1, 1, 1), BIP_DataBase)
            If Unifac_Error < 0 Then GoTo Set_BIP
            Units = "unitless"
        Case Else
            Exit Function
    End Select
    
    Calc_Unifac = True
Exit Function

Handler:
    MsgBox Err & " " & Error & Chr(13) & "In form frmShredDisp : UNIFAC Calc"
'    Resume Next
End Function

Function GetMasterPropInfo(Value As Double, Units As String, PropName_801 As String, PropCode_911 As String, Chem_Name) As Boolean
'        Arguments:
'            Value -> is the value obtained from the database
'            Units -> the units obtained from the database
'            PropName_801 -> a string representing the column name in the 801 database
'            Propcode_911 -> a string representing a property for the ChemName in the 911 database
'            Chem_Name -> a string representing the Chemical Name for pulling info from the database
'
'        Return Value:
'            True/False -> if the function was succesful on pulling data or not

Dim sql As String
Dim DB As Database
Dim RS As Recordset
Dim CASnum As Double
    On Error GoTo Handler

    GetMasterPropInfo = False
    If Not FileExists(AppPath & "Master.mdb") Then
        Exit Function
    End If
Try801:
    If PropName_801 <> "" Then
        Set DB = OpenDatabase(AppPath & "Master.mdb", False, False)
        sql = "select * from [DIPPR801] where [Name] = '" & Trim(Chem_Name) & "'"
        Set RS = DB.OpenRecordset(sql, dbOpenSnapshot)
        If RS.recordcount > 0 Then
            If Not IsNull(RS(PropName_801)) Then
                Value = RS(PropName_801)
            End If
        Else
            RS.Close
            DB.Close
            GoTo Try911
        End If
        RS.Close
        DB.Close
        GetMasterPropInfo = True
        Exit Function
    End If
Try911:
    If PropCode_911 <> "" Then
        Set DB = OpenDatabase(AppPath & "Master.mdb", False, False)
        sql = "select * from [Chemical Name] where [Name] = '" & Trim(Chem_Name) & "'"
        Set RS = DB.OpenRecordset(sql, dbOpenSnapshot)
        If RS.recordcount > 0 Then
            CASnum = RS("CAS")
            RS.Close
            sql = "select [Property Code],[Cas #],[Value],[Units] from [DIPPR911] where [Property Code] = '" & PropCode_911 & "' and [Cas #] = " & CASnum
            Set RS = DB.OpenRecordset(sql, dbOpenSnapshot)
            If RS.recordcount > 0 Then
                If Not IsNull(RS("Value")) Then
                    Value = RS("Value")
                End If
                If Not IsNull(RS("Units")) Then
                    Units = RS("Units")
                End If
            Else
                RS.Close
                DB.Close
                Exit Function
            End If
        Else
            RS.Close
            DB.Close
            Exit Function
        End If
        RS.Close
        DB.Close
        GetMasterPropInfo = True
    End If
Exit Function

Handler:
    MsgBox Err & " " & Error & Chr(13) & "In form frmShredDisp : GetMasterPropInfo"
End Function

'-----------------------------------------------------------
' FUNCTION: FileExists
' Determines whether the specified file exists
'
' IN: [strPathName] - file to check for
'
' Returns: True if file exists, False otherwise
'-----------------------------------------------------------
'
Function FileExists(ByVal strPathName As String) As Integer
    Dim intFileNum As Integer

    On Error Resume Next

    '
    'Remove any trailing directory separator character
    '
    If Right$(strPathName, 1) = "\" Then
        strPathName = Left$(strPathName, Len(strPathName) - 1)
    End If

    '
    'Attempt to open the file, return value of this function is False
    'if an error occurs on open, True otherwise
    '
    intFileNum = FreeFile
    Open strPathName For Input As intFileNum

    FileExists = IIf(Err = 0, True, False)

    Close intFileNum

    Err = 0
End Function


Function AppPath() As String
'        Arguments:
'            None
'
'        Return Value:
'            The Application path with only one "/" as the trailing character

    AppPath = App.path
    If Right(App.path, 1) <> "/" Then
        AppPath = AppPath & "/"
    End If
End Function
