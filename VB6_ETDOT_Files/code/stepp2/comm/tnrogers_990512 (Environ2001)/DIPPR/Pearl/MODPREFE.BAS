Attribute VB_Name = "modpreferences"
Sub LoadPreferences()
    
    Dim i As Integer
        
    SetDefaultUnit = True
    
    FRMPreferences!GRDDefaultUnits.ColWidth(0) = 3000
    FRMPreferences!GRDDefaultUnits.ColWidth(1) = 1160
    FRMPreferences!GRDDefaultUnits.Row = 0
    FRMPreferences!GRDDefaultUnits.Col = 0
    FRMPreferences!GRDDefaultUnits.Text = "Property Name"
    FRMPreferences!GRDDefaultUnits.Col = 1
    FRMPreferences!GRDDefaultUnits.Text = "Default Unit"
   
    FRMPreferences!TXTGT1000.Text = FormatGT1000
    FRMPreferences!TXTLT001.Text = FormatLT001
    FRMPreferences!TXTGeneral.Text = FormatGeneral
    
    FRMPreferences!CMBProperty.AddItem "Activity Coefficient"
    FRMPreferences!CMBProperty.AddItem "log Kow"
    FRMPreferences!CMBProperty.AddItem "Solubility"
    
    FRMPreferences!CMB1stChoice.AddItem "AGLB"
    FRMPreferences!CMB1stChoice.AddItem "AVLE"
    FRMPreferences!CMB1stChoice.AddItem "AENV"
    FRMPreferences!CMB1stChoice.AddItem "ALLE"
    
    FRMPreferences!CMB2ndChoice.AddItem "AGLB"
    FRMPreferences!CMB2ndChoice.AddItem "AVLE"
    FRMPreferences!CMB2ndChoice.AddItem "AENV"
    FRMPreferences!CMB2ndChoice.AddItem "ALLE"
    
    FRMPreferences!CMB3rdChoice.AddItem "AGLB"
    FRMPreferences!CMB3rdChoice.AddItem "AVLE"
    FRMPreferences!CMB3rdChoice.AddItem "AENV"
    FRMPreferences!CMB3rdChoice.AddItem "ALLE"
    
    FRMPreferences!CMBProperty.ListIndex = 0
        
    For i = 0 To NumProperties
        FRMPreferences!GRDDefaultUnits.Row = i + 1
        FRMPreferences!GRDDefaultUnits.Col = 0
        FRMPreferences!GRDDefaultUnits.Text = GetPropName(i)
        FRMPreferences!GRDDefaultUnits.Col = 1
        FRMPreferences!GRDDefaultUnits.Text = DefaultUnit(i)
    Next i
        
    FRMPreferences!GRDDefaultUnits.Row = i + 1
    FRMPreferences!GRDDefaultUnits.Col = 0
    FRMPreferences!GRDDefaultUnits.Text = "f(T) Temperatures"
    FRMPreferences!GRDDefaultUnits.Col = 1
    FRMPreferences!GRDDefaultUnits.Text = DefaultTFTUnit
    
    ' the block 5 preferences
    FRMPreferences!uflcmbx.Clear
    FRMPreferences!lflcmbx.Clear
    FRMPreferences!fpcmbx.Clear
    FRMPreferences!aitcmbx.Clear
    For i = 0 To NumMethods
        prefstring = get_B5_method_name(1, i)
        If Len(prefstring) > 3 Then
            FRMPreferences!uflcmbx.AddItem prefstring
        End If
        prefstring = get_B5_method_name(2, i)
        If Len(prefstring) > 3 Then
            FRMPreferences!lflcmbx.AddItem prefstring
        End If
        prefstring = get_B5_method_name(3, i)
        If Len(prefstring) > 3 Then
            FRMPreferences!fpcmbx.AddItem prefstring
        End If
        prefstring = get_B5_method_name(4, i)
        If Len(prefstring) > 3 Then
            FRMPreferences!aitcmbx.AddItem prefstring
        End If
    Next i
    On Error Resume Next
    FRMPreferences!uflcmbx.ListIndex = 0
    FRMPreferences!lflcmbx.ListIndex = 0
    FRMPreferences!fpcmbx.ListIndex = 0
    FRMPreferences!aitcmbx.ListIndex = 0
    PrefStartup = False
    
End Sub



Public Function get_B5_method_name(property As Integer, method As Integer) As String

    ' this function takes two indexes for the B5Preference array and
    ' returns the name of the corresponding method
    
    Dim answer As String
    Select Case property
        Case 1      ' UFL
            Select Case B5Preference(0, method)
                Case 1
                    answer = "801 Database"
                Case 2
                    answer = "911 Database"
                'Case 3
                '    answer = "Master Database"
                Case 4
                    answer = "MTU Fire & Explosion Data"
                Case 5
                    answer = "MTU Group Contribution"
                Case 6
                    answer = "MTU Combustion Reaction"
                Case 7
                    answer = "Penn State Group Contribution"
                Case Else
                    answer = " "
            End Select
        Case 2  ' LFL
            Select Case B5Preference(1, method)
                Case 1
                    answer = "801 Database"
                Case 2
                    answer = "911 Database"
                'Case 3
                '    answer = "Master Database"
                Case 4
                    answer = "MTU Fire & Explosion Data"
                Case 5
                    answer = "MTU Group Contribution"
                Case 6
                    answer = "Penn State Group Contribution"
                Case 7
                    answer = "MTU Combustion Reaction"
                Case 8
                    answer = "MTU FlashPoint Method"
                Case Else
                    answer = " "
                End Select
                
        Case 3
            Select Case B5Preference(2, method)
                Case 1
                    answer = "801 Database"
                Case 2
                    answer = "911 Database"
                'Case 3
                '    answer = "Master Database"
                Case 4
                    answer = "MTU Fire & Explosion Data"
                Case 5
                    answer = "LFL Data"
                Case 6
                    answer = "MTU LFL Group Contribution"
                Case 7
                    answer = "Penn State Group Contribution"
                Case 8
                    answer = "MTU LFL Combustion Reaction"
                Case Else
                    answer = " "
                End Select
                
        Case 4
            Select Case B5Preference(3, method)
                Case 1
                    answer = "801 Database"
                Case 2
                    answer = "911 Database"
                'Case 3
                '    answer = "Master Database"
                Case 4
                    answer = "MTU Fire & Explosion Data"
                Case 5
                    answer = "MTU Logarithmic Method"
                Case 6
                    answer = "MTU Linear Method"
                Case Else
                    answer = " "
                End Select
                
        End Select
        get_B5_method_name = answer
                    
                    
End Function



Public Sub update_B5_preferences()

' this function does the reverse, it goes through
' the list on the preference form and updates the
' B5 preference array

    Dim i As Integer
    Dim array_position As Integer
    Dim method_index As Integer
    
    ' first find the first one
    method_index = FRMPreferences!uflcmbx.ListIndex
    Select Case Trim(FRMPreferences!uflcmbx.List(method_index))
            Case "801 Database"
                B5Preference(0, 0) = 1
            Case "911 Database"
                B5Preference(0, 0) = 2
            Case "Master Database"
                B5Preference(0, 0) = 3
            Case "MTU Fire & Explosion Data"
                B5Preference(0, 0) = 4
            Case "MTU Group Contribution"
                B5Preference(0, 0) = 5
            Case "MTU Combustion Reaction"
                B5Preference(0, 0) = 6
            Case "Penn State Group Contribution"
                B5Preference(0, 0) = 7
            Case Else
                B5Preference(0, 0) = 0  ' the null value
        End Select
    ' now fill in the rest of the preferences
    array_position = 1
    For i = 0 To FRMPreferences!uflcmbx.ListCount - 1
        Select Case Trim(FRMPreferences!uflcmbx.List(i))
            Case Trim(FRMPreferences!uflcmbx.List(method_index))
                'continue
            Case "801 Database"
                B5Preference(0, array_position) = 1
                array_position = array_position + 1
            Case "911 Database"
                B5Preference(0, array_position) = 2
                array_position = array_position + 1
            Case "Master Database"
                B5Preference(0, array_position) = 3
                array_position = array_position + 1
            Case "MTU Fire & Explosion Data"
                B5Preference(0, array_position) = 4
                array_position = array_position + 1
            Case "MTU Group Contribution"
                B5Preference(0, array_position) = 5
                array_position = array_position + 1
            Case "MTU Combustion Reaction"
                B5Preference(0, array_position) = 6
                array_position = array_position + 1
            Case "Penn State Group Contribution"
                B5Preference(0, array_position) = 7
                array_position = array_position + 1
            Case Else
                B5Preference(0, array_position) = 0
                array_position = array_position + 1
        End Select
    Next i
    ' the lower flam limit preferences
    method_index = FRMPreferences!lflcmbx.ListIndex
    Select Case Trim(FRMPreferences!lflcmbx.List(method_index))
            Case "801 Database"
                B5Preference(1, 0) = 1
            Case "911 Database"
                B5Preference(1, 0) = 2
            Case "Master Database"
                B5Preference(1, 0) = 3
            Case "MTU Fire & Explosion Data"
                B5Preference(1, 0) = 4
            Case "MTU Group Contribution"
                B5Preference(1, 0) = 5
            Case "Penn State Group Contribution"
                B5Preference(1, 0) = 6
            Case "MTU Combustion Reaction"
                B5Preference(1, 0) = 7
            Case "MTU FlashPoint Method"
                B5Preference(1, 0) = 8
            Case Else
                B5Preference(1, 0) = 0
        End Select
    array_position = 1
    For i = 0 To FRMPreferences!lflcmbx.ListCount - 1
        Select Case Trim(FRMPreferences!lflcmbx.List(i))
            Case Trim(FRMPreferences!lflcmbx.List(method_index))
                'continue
            Case "801 Database"
                B5Preference(1, array_position) = 1
                array_position = array_position + 1
            Case "911 Database"
                B5Preference(1, array_position) = 2
                array_position = array_position + 1
            Case "Master Database"
                B5Preference(1, array_position) = 3
                array_position = array_position + 1
            Case "MTU Fire & Explosion Data"
                B5Preference(1, array_position) = 4
                array_position = array_position + 1
            Case "MTU Group Contribution"
                B5Preference(1, array_position) = 5
                array_position = array_position + 1
            Case "Penn State Group Contribution"
                B5Preference(1, array_position) = 6
                array_position = array_position + 1
            Case "MTU Combustion Reaction"
                B5Preference(1, array_position) = 7
                array_position = array_position + 1
            Case "MTU FlashPoint Method"
                B5Preference(1, array_position) = 8
                array_position = array_position + 1
            Case Else
                B5Preference(1, array_position) = 0
                array_position = array_position + 1
        End Select
    Next i
    ' now flashpoint preferences
    method_index = FRMPreferences!fpcmbx.ListIndex
    Select Case Trim(FRMPreferences!fpcmbx.List(method_index))
            Case "801 Database"
                B5Preference(2, 0) = 1
            Case "911 Database"
                B5Preference(2, 0) = 2
            Case "Master Database"
                B5Preference(2, 0) = 3
            Case "MTU Fire & Explosion Data"
                B5Preference(2, 0) = 4
            Case "LFL Data"
                B5Preference(2, 0) = 5
            Case "MTU LFL Group Contribution"
                B5Preference(2, 0) = 6
            Case "Penn State Group Contribution"
                B5Preference(2, 0) = 7
            Case "MTU LFL Combustion Reaction"
                B5Preference(2, 0) = 8
            Case Else
                B5Preference(2, 0) = 0
        End Select
    array_position = 1
    For i = 0 To FRMPreferences!fpcmbx.ListCount - 1
        Select Case Trim(FRMPreferences!fpcmbx.List(i))
            Case Trim(FRMPreferences!fpcmbx.List(method_index))
                'continue
            Case "801 Database"
                B5Preference(2, array_position) = 1
                array_position = array_position + 1
            Case "911 Database"
                B5Preference(2, array_position) = 2
                array_position = array_position + 1
            Case "Master Database"
                B5Preference(2, array_position) = 3
                array_position = array_position + 1
            Case "MTU Fire & Explosion Data"
                B5Preference(2, array_position) = 4
                array_position = array_position + 1
            Case "LFL Data"
                B5Preference(2, array_position) = 5
                array_position = array_position + 1
            Case "MTU LFL Group Contribution"
                B5Preference(2, array_position) = 6
                array_position = array_position + 1
            Case "Penn State Group Contribution"
                B5Preference(2, array_position) = 7
                array_position = array_position + 1
            Case "MTU LFL Combustion Reaction"
                B5Preference(2, array_position) = 8
                array_position = array_position + 1
            Case Else
                B5Preference(2, array_position) = 0
                array_position = array_position + 1
        End Select
    Next i
    
    ' now ait preferences
    method_index = FRMPreferences!aitcmbx.ListIndex
    Select Case Trim(FRMPreferences!aitcmbx.List(method_index))
            Case "801 Database"
                B5Preference(3, 0) = 1
            Case "911 Database"
                B5Preference(3, 0) = 2
            Case "Master Database"
                B5Preference(3, 0) = 3
            Case "MTU Fire & Explosion Data"
                B5Preference(3, 0) = 4
            Case "MTU Logarithmic Method"
                B5Preference(3, 0) = 5
            Case "MTU Linear Method"
                B5Preference(3, 0) = 6
            Case Else
                B5Preference(3, 0) = 0
        End Select
    array_position = 1
    For i = 0 To FRMPreferences!aitcmbx.ListCount - 1
        Select Case Trim(FRMPreferences!aitcmbx.List(i))
            Case Trim(FRMPreferences!aitcmbx.List(method_index))
                'continue
            Case "801 Database"
                B5Preference(3, array_position) = 1
                array_position = array_position + 1
            Case "911 Database"
                B5Preference(3, array_position) = 2
                array_position = array_position + 1
            Case "Master Database"
                B5Preference(3, array_position) = 3
                array_position = array_position + 1
            Case "MTU Logarithmic Method"
                B5Preference(3, array_position) = 4
                array_position = array_position + 1
            Case "MTU Linear Method"
                B5Preference(3, array_position) = 5
                array_position = array_position + 1
            Case Else
                B5Preference(3, array_position) = 0
                array_position = array_position + 1
        End Select
    Next i
        
        
            
                
End Sub


