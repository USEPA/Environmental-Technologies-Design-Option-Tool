Attribute VB_Name = "AddChemMod"

Sub addtolist(casnum As Integer)
    Dim i As Integer, J As Integer
    Dim msg$, Response As Integer

' RETURNS FALSE IF THE CHEMICAL CAN CONTINUE ON
' ALWAYS RETURNS FALSE IF NOT IN DEMOMODE CHECK DEMOMODE.BAS
    ''''If (demo_check_chemicals(contam_prop_form.contam_combo)) Then Exit Sub

    If NumSelectedChemicals = MAXSELECTEDCHEMICALS Then
       msg$ = "The maximum number of contaminants that can be selected at a time in the StEPP program is " & Str$(MAXSELECTEDCHEMICALS) & ".  Therefore, you may not select this chemical unless you Unselect a contaminant you selected previously or begin the program again."
       MsgBox msg$, MB_ICONSTOP, "Too Many Contaminants Selected"
       Exit Sub
    End If

    Screen.MousePointer = 11   'Hourglass

    Update_Fields (casnum)

    If NumSelectedChemicals = 0 Then
       contam_prop_form!mnuFile(4).Enabled = True
       contam_prop_form!mnuFile(5).Enabled = True
       contam_prop_form!mnuFile(7).Enabled = True
       contam_prop_form!cmdUnselectContaminant.Enabled = True
       
'*** Modification v1019 by David R. Hokanson (16may2000)
'       Call frmmain.frmMain_Reset_DemoVersionDisablings
       Call contam_prop_form.frmMain_Reset_DemoVersionDisablings
'*** End Modification v1019 by David R. Hokanson (16may2000)

    End If

    For i = 0 To contam_prop_form.cboSelectContaminant.ListCount - 1
        If Trim$(contam_prop_form.cboSelectContaminant.List(i)) = Trim$(contam_prop_form.contam_combo.List(contam_prop_form.contam_combo.ListIndex)) Then
           msg$ = "There is already a contaminant named "
           msg$ = msg$ + contam_prop_form.contam_combo.List(contam_prop_form.contam_combo.ListIndex) + " selected. "
           msg$ = msg$ + Chr$(13) + Chr$(13)
           msg$ = msg$ + "Do you wish to reinitialize it to default properties by selecting it now?"
           Response = MsgBox(msg$, MB_ICONQUESTION + MB_YESNO, "Contaminant Already Selected")
           If Response = IDYES Then
              If Trim$(contam_prop_form.contam_combo.List(contam_prop_form.contam_combo.ListIndex)) <> Trim$(contam_prop_form.cboSelectContaminant.Text) Then  'If contaminant currently selected is not the one being replaced then update its values before performing calculations
                 For J = 1 To NUMBER_OF_PROPERTIES
                     phprop.HaveProperty(J) = HaveProperty(J)
                 Next J
                 For J = 1 To NUMBER_OF_PROPERTIES_AVAILABLE
                     phprop.PROPAVAILABLE(J) = PROPAVAILABLE(J)
                 Next J
                 PropContaminant(PreviouslySelectedIndex) = phprop
              End If

              contam_prop_form.cboSelectContaminant.RemoveItem i
              For J = i + 2 To NumSelectedChemicals
                  PropContaminant(J - 1) = PropContaminant(J)
              Next J
              NumSelectedChemicals = NumSelectedChemicals - 1
              Exit For
           Else
              Screen.MousePointer = 0   'Arrow
              Exit Sub
           End If
        End If
    Next i

    contam_prop_form.cboSelectContaminant.AddItem contam_prop_form.contam_combo.List(contam_prop_form.contam_combo.ListIndex)
    contam_prop_form.lblSelectedContaminant.Caption = Trim$(dbinput.Name)

    'Update the contaminant selected prior to the new one if necessary
    If PreviouslySelectedIndex >= 0 Then
       For J = 1 To NUMBER_OF_PROPERTIES
           phprop.HaveProperty(J) = HaveProperty(J)
       Next J
       For J = 1 To NUMBER_OF_PROPERTIES_AVAILABLE
           phprop.PROPAVAILABLE(J) = PROPAVAILABLE(J)
       Next J
       PropContaminant(PreviouslySelectedIndex) = phprop
    End If

'* initialize binary interaction parameter database choices
    phprop.ActivityCoefficient.BinaryInteractionParameterDatabase = BIP_dbHierarchy.ActivityCoefficient(1)
    phprop.AqueousSolubility.BinaryInteractionParameterDatabase = BIP_dbHierarchy.AqueousSolubility(1)
    phprop.OctWaterPartCoeff.BinaryInteractionParameterDatabase = BIP_dbHierarchy.OctWaterPartCoeff(1)
    For i = 1 To 3
        phprop.ActivityCoefficient.BinaryInteractionParameterDBAvailable(i) = True
        phprop.AqueousSolubility.BinaryInteractionParameterDBAvailable(i) = True
        If i <> 3 Then phprop.OctWaterPartCoeff.BinaryInteractionParameterDBAvailable(i) = True
    Next i
    UserSelectedTheUnifacBIPDBActCoeff = False
    UserSelectedTheUnifacBIPDBAqSol = False
    UserSelectedTheUnifacBIPDBKow = False

'* Set Current Selections to None
    Call InitializeCurrentSelections

    NumSelectedChemicals = NumSelectedChemicals + 1
    Call InitializeHilights
    Call InitializePROPandHAVEAVAILABLEArrays
    Call InitializeUserInputs

    Call BlankAllTextBoxes
    frmWaitForCalculations.Show
    frmWaitForCalculations.Refresh

' THIS IS HERE TO MAKE SURE THAT THE CURRENT DIRECTORY IS WITH THE FORTRAN DLL
' FILES.   THE DIFFERENT *.DAT FILES THAT ARE USED BY THE FORTRAN DLLS MUST BE THERE
'    msg$ = CurDir$
'    ChDrive app.path
'    ChDir app.path + "\dlls"

    Call DoCalculationForThisContaminant

' RETURNING TO WHERE WE WERE BEFORE
'    ChDrive msg$
'    ChDir msg$

    phprop.AqueousSolubility.PreviousBinaryInteractionParameterDB = phprop.AqueousSolubility.BinaryInteractionParameterDatabase
    
    If NumSelectedChemicals > 0 Then contam_prop_form.cboSelectContaminant.Enabled = True

    frmWaitForCalculations.Hide
    contam_prop_form.cboSelectContaminant.ListIndex = contam_prop_form.cboSelectContaminant.ListCount - 1
    contam_prop_form.cboSelectContaminant.SetFocus
  
    Screen.MousePointer = 0   'Arrow
End Sub

Sub BlankAllTextBoxes()
    Dim i As Integer

    contam_prop_form.lblSelectedContaminant.Caption = ""

    For i = 0 To 12
      contam_prop_form.lblContaminantProperties(i).Caption = ""
    Next i

    For i = 0 To 4
       contam_prop_form.lblAirWaterProperties(i).Caption = ""
    Next i

End Sub

Function get_source(s As String) As Long
    If (s = "DIPPR801") Then
        get_source = 4
    ElseIf (s = "YAWS") Then
        get_source = 1
    ElseIf (s = "SUPERFUND") Then
        get_source = 2
    ElseIf (s = "RTI") Then
        get_source = 3
    Else
        get_source = -1
    End If
End Function

Function nullcheck(s) As String
    If VarType(s) <> 8 Then
        nullcheck = ""
    Else
        nullcheck = s
    End If
End Function

Function number(num) As String
    If VarType(num) = 2 Or VarType(num) = 5 Or VarType(num) = 3 Then
        If num = -1 Then
            number = ""
        Else
            number = Str$(num)
        End If
    Else
        number = ""
    End If
End Function

Sub Update_Fields(RecordNo As Long)
    
    Dim i As Long
    Dim J As Long
    Dim K As Long
    Dim TempD As Double
    Dim HC_Count As Long
    Dim hc_string As String
    Dim last_hc_string As String
    Dim HC_DB_Source As String
    Dim HC_DB_Value As String * 36
    Dim HC_DB_Temp As String
    
    dbinput.CasNumber = db_index(RecordNo + 1)
    
    '
    ' OPEN RECORDSET.
    '
    Set RS_Main = DB_Main.OpenRecordset( _
        "SELECT * FROM [Names (Master)] WHERE [Names (Master)].CAS = " & _
        Format$(dbinput.CasNumber, "0"))
    If (RS_Main.EOF = False) Then
      RS_Main.MoveFirst
      RS_Main.MoveLast
      RS_Main.MoveFirst
    End If
    Set Selection = RS_Main
    'If (DemoMode) Then
    '    contam_prop_form.Data1.DatabaseName = Database_Path + "\demo_db.mdb"
    'Else
    '    contam_prop_form.Data1.DatabaseName = Database_Path + "\stepp_db.mdb"
    'End If
    'contam_prop_form.Data1.RecordSource = "SELECT * FROM [Names (Master)] WHERE [Names (Master)].CAS = " & Format$(dbinput.CASNumber, "0")
    'contam_prop_form.Data1.Refresh
    'Set Selection = contam_prop_form.Data1.Recordset

    dbinput.Name = Selection(2)

    'Look into the Properties Table ----------------------------------

    '
    ' OPEN RECORDSET.
    '
    Set RS_Main = DB_Main.OpenRecordset( _
        "SELECT * FROM DIPPR801 WHERE DIPPR801.CAS = " & _
        Format$(dbinput.CasNumber, "0"))
    If (RS_Main.EOF = False) Then
      RS_Main.MoveFirst
      RS_Main.MoveLast
      RS_Main.MoveFirst
    End If
    Set Selection = RS_Main
    'contam_prop_form.Data1.RecordSource = "SELECT * FROM DIPPR801 WHERE DIPPR801.CAS = " & Format$(dbinput.CASNumber, "0")
    'contam_prop_form.Data1.Refresh
    'Set Selection = contam_prop_form.Data1.Recordset

    If Selection.EOF = False Then
        
        dbinput.formula = nullcheck(Selection("FORM"))
        dbinput.MolecularWeight = Selection("MW")
        dbinput.BoilingPoint = Selection("NBP")
        dbinput.BoilingPointSource = get_source(nullcheck("DIPPR801"))
        dbinput.RefractiveIndex = Selection("RI")
        dbinput.VaporPressureDatabaseEquation = Selection("VPEQN")
        dbinput.VaporPressureNumberCoefficients = Selection("VPNUM")
        dbinput.VaporPressureAntoineA = Selection("VPA")
        dbinput.VaporPressureAntoineB = Selection("VPB")
        dbinput.VaporPressureAntoineC = Selection("VPC")
        dbinput.VaporPressureAntoineD = Selection("VPD")
        dbinput.VaporPressureAntoineE = Selection("VPE")
        dbinput.VaporPressureMinimumT = Selection("VPTMIN")
        dbinput.VaporPressureMaximumT = Selection("VPTMAX")
        dbinput.VaporPressureSource = get_source(nullcheck("DIPPR801"))
        dbinput.LiquidDensityEquation = Selection("LDNEQN")
        dbinput.LiquidDensityNumberCoefficients = Selection("LDNNUM")
        dbinput.LiquidDensityCoefficientA = Selection("LDNA")
        dbinput.LiquidDensityCoefficientB = Selection("LDNB")
        dbinput.LiquidDensityCoefficientC = Selection("LDNC")
        dbinput.LiquidDensityCoefficientD = Selection("LDND")
        dbinput.LiquidDensityMinimumT = Selection("LDNTMIN")
        dbinput.LiquidDensityMaximumT = Selection("LDNTMAX")
        dbinput.LiquidDensitySource = get_source(nullcheck("DIPPR801"))
    
    Else
        
        dbinput.MolecularWeight = -1
        dbinput.BoilingPointSource = -1
        dbinput.RefractiveIndex = -1
        dbinput.VaporPressureAntoineA = -1
        dbinput.VaporPressureDatabaseEquation = -1
        dbinput.LiquidDensityEquation = -1

    End If

    If dbinput.MolecularWeight = 0 Then
        dbinput.MolecularWeight = -1
    End If
    
    If dbinput.BoilingPoint = 0 Then
        dbinput.BoilingPointSource = -1
    End If
    
    If dbinput.RefractiveIndex = 0 Then
        dbinput.RefractiveIndex = -1
    End If
    
    If dbinput.VaporPressureAntoineA = 0 Then
        dbinput.VaporPressureAntoineA = -1
        dbinput.VaporPressureDatabaseEquation = -1
    End If
    
    If dbinput.LiquidDensityEquation = 0 Then
        dbinput.LiquidDensityEquation = -1
    End If
    
    If dbinput.VaporPressureAntoineA = -1 Then

        '
        ' OPEN RECORDSET.
        '
        Set RS_Main = DB_Main.OpenRecordset( _
            "SELECT * FROM [VP Yaws] WHERE [VP Yaws].CAS = " & _
            Format$(dbinput.CasNumber, "0"))
        If (RS_Main.EOF = False) Then
          RS_Main.MoveFirst
          RS_Main.MoveLast
          RS_Main.MoveFirst
        End If
        Set Selection = RS_Main
        'contam_prop_form.Data1.RecordSource = "SELECT * FROM [VP Yaws] WHERE [VP Yaws].CAS = " & Format$(dbinput.CASNumber, "0")
        'contam_prop_form.Data1.Refresh
        'Set Selection = contam_prop_form.Data1.Recordset
        
        If Selection.EOF = False Then
            
            dbinput.VaporPressureNumberCoefficients = 3
            dbinput.VaporPressureAntoineA = Selection("ANTA")
            dbinput.VaporPressureAntoineB = Selection("ANTB")
            dbinput.VaporPressureAntoineC = Selection("ANTC")
            dbinput.VaporPressureMinimumT = Selection("MINT")
            dbinput.VaporPressureMaximumT = Selection("MAXT")
            dbinput.VaporPressureSource = get_source(nullcheck("YAWS"))
        
        Else
        
            dbinput.VaporPressureAntoineA = -1
            dbinput.VaporPressureDatabaseEquation = -1

        End If

    End If

    If dbinput.VaporPressureAntoineA = 0 Then
        dbinput.VaporPressureAntoineA = -1
        dbinput.VaporPressureDatabaseEquation = -1
    End If
    
    If dbinput.VaporPressureAntoineA = -1 Then
    
        '
        ' OPEN RECORDSET.
        '
        Set RS_Main = DB_Main.OpenRecordset( _
            "SELECT * FROM [VP@25 Superfund] WHERE [VP@25 Superfund].CAS = " & _
            Format$(dbinput.CasNumber, "0"))
        If (RS_Main.EOF = False) Then
          RS_Main.MoveFirst
          RS_Main.MoveLast
          RS_Main.MoveFirst
        End If
        Set Selection = RS_Main
        'contam_prop_form.Data1.RecordSource = "SELECT * FROM [VP@25 Superfund] WHERE [VP@25 Superfund].CAS = " & Format$(dbinput.CASNumber, "0")
        'contam_prop_form.Data1.Refresh
        'Set Selection = contam_prop_form.Data1.Recordset
    
        If Selection.EOF = False Then
            dbinput.VaporPressureSuperfund = Selection("VP")
            dbinput.VaporPressureSuperfundTemperature = 25
            dbinput.VaporPressureSource = get_source(nullcheck("SUPERFUND"))
        Else
            dbinput.VaporPressureSuperfund = -1
        End If

    End If

    If dbinput.VaporPressureSuperfund = 0 Then
        dbinput.VaporPressureSuperfund = -1
    End If
    
    '
    ' OPEN RECORDSET.
    '
    Set RS_Main = DB_Main.OpenRecordset( _
        "SELECT * FROM [SB@25 Yaws] WHERE [SB@25 Yaws].CAS = " & _
        Format$(dbinput.CasNumber, "0"))
    If (RS_Main.EOF = False) Then
      RS_Main.MoveFirst
      RS_Main.MoveLast
      RS_Main.MoveFirst
    End If
    Set Selection = RS_Main
    'contam_prop_form.Data1.RecordSource = "SELECT * FROM [SB@25 Yaws] WHERE [SB@25 Yaws].CAS = " & Format$(dbinput.CASNumber, "0")
    'contam_prop_form.Data1.Refresh
    'Set Selection = contam_prop_form.Data1.Recordset
    
    If Selection.EOF = False Then
        dbinput.AqueousSolubility = Selection("Sol")
        dbinput.AqueousSolubilityTemperature = 25
        dbinput.AqueousSolubilitySource = get_source(nullcheck("YAWS"))
    Else
        dbinput.AqueousSolubility = -1
    End If

    If dbinput.AqueousSolubility = 0 Then

        '
        ' OPEN RECORDSET.
        '
        Set RS_Main = DB_Main.OpenRecordset( _
            "SELECT * FROM [SB@25 Superfund] WHERE [SB@25 Superfund].CAS = " & _
            Format$(dbinput.CasNumber, "0"))
        If (RS_Main.EOF = False) Then
          RS_Main.MoveFirst
          RS_Main.MoveLast
          RS_Main.MoveFirst
        End If
        Set Selection = RS_Main
        'contam_prop_form.Data1.RecordSource = "SELECT * FROM [SB@25 Superfund] WHERE [SB@25 Superfund].CAS = " & Format$(dbinput.CASNumber, "0")
        'contam_prop_form.Data1.Refresh
        'Set Selection = contam_prop_form.Data1.Recordset
    
        If Selection.EOF = False Then
            dbinput.AqueousSolubility = Selection("Sol")
            dbinput.AqueousSolubilityTemperature = 25
            dbinput.AqueousSolubilitySource = get_source(nullcheck("SUPERFUND"))
        
        Else
        
            dbinput.AqueousSolubility = -1
        
        End If
        
    End If

    If dbinput.AqueousSolubility = 0 Then
        dbinput.AqueousSolubility = -1
    End If
    
    '
    ' OPEN RECORDSET.
    '
    Set RS_Main = DB_Main.OpenRecordset( _
        "SELECT * FROM [Kow@25 Superfund] WHERE [Kow@25 Superfund].CAS = " & _
        Format$(dbinput.CasNumber, "0"))
    If (RS_Main.EOF = False) Then
      RS_Main.MoveFirst
      RS_Main.MoveLast
      RS_Main.MoveFirst
    End If
    Set Selection = RS_Main
    'contam_prop_form.Data1.RecordSource = "SELECT * FROM [Kow@25 Superfund] WHERE [Kow@25 Superfund].CAS = " & Format$(dbinput.CASNumber, "0")
    'contam_prop_form.Data1.Refresh
    'Set Selection = contam_prop_form.Data1.Recordset
    
    If Selection.EOF = False Then
        dbinput.OctWaterPartCoeff = Selection("log Kow")
        dbinput.OctWaterPartCoeffTemperature = 25
        dbinput.OctWaterPartCoeffSource = get_source(nullcheck("SUPERFUND"))
    Else
        dbinput.OctWaterPartCoeff = -1
    End If
    
    If dbinput.OctWaterPartCoeff = 0 Then
        dbinput.OctWaterPartCoeff = -1
    End If
    
    '
    ' OPEN RECORDSET.
    '
    Set RS_Main = DB_Main.OpenRecordset( _
        "SELECT * FROM [Rogers/Miller] WHERE [Rogers/Miller].CAS = " & _
        Format$(dbinput.CasNumber, "0"))
    If (RS_Main.EOF = False) Then
      RS_Main.MoveFirst
      RS_Main.MoveLast
      RS_Main.MoveFirst
    End If
    Set Selection = RS_Main
    'contam_prop_form.Data1.RecordSource = "SELECT * FROM [Rogers/Miller] WHERE [Rogers/Miller].CAS = " & Format$(dbinput.CASNumber, "0")
    'contam_prop_form.Data1.Refresh
    'Set Selection = contam_prop_form.Data1.Recordset
    
    If Selection.EOF = False Then
        
        If Selection("MX") <= 0 Then dbinput.MaximumUnifacGroups = 0
    
        For i = 1 To NC
            For J = 1 To 10
                For K = 1 To 2
                    dbinput.MS(i, J, K) = 0
                Next K
            Next J
        Next i

        dbinput.NumberofRingsinCompound = Selection("RG")
        dbinput.MaximumUnifacGroups = Selection("MX")
        
        For i = 1 To dbinput.MaximumUnifacGroups
            dbinput.MS(NC, i, 1) = Selection("G" + Trim$(Str$(i)))
            dbinput.MS(NC, i, 2) = Selection("N" + Trim$(Str$(i)))
        Next i
    
    Else
       
       dbinput.NumberofRingsinCompound = -1
       dbinput.MaximumUnifacGroups = -1
    
    End If

    If dbinput.formula = "" Then
        If Selection.EOF = False Then
            dbinput.formula = Selection("Formula")
        End If
    End If

    HC_Count = 0
    hc_string = ""

    '
    ' OPEN RECORDSET.
    '
    Set RS_Main = DB_Main.OpenRecordset( _
        "SELECT * FROM [HC RTI] WHERE [HC RTI].CAS = " & _
        Format$(dbinput.CasNumber, "0"))
    If (RS_Main.EOF = False) Then
      RS_Main.MoveFirst
      RS_Main.MoveLast
      RS_Main.MoveFirst
    End If
    Set Selection = RS_Main
    'contam_prop_form.Data1.RecordSource = "SELECT * FROM [HC RTI] WHERE [HC RTI].CAS = " & Format$(dbinput.CASNumber, "0")
    'contam_prop_form.Data1.Refresh
    'Set Selection = contam_prop_form.Data1.Recordset

    Do While Not Selection.EOF

        If number(Selection(1)) <> "" Then
            
            last_hc_string = hc_string
            HC_DB_Source = "RTI"
            hc_form!lblDatabase = HC_DB_Source
            LSet HC_DB_Value = Format$(number(Selection(1)), GetTheFormat(CDbl(number(Selection(1)))))
            HC_DB_Temp = Format$(number(Selection(2)), GetTheFormat(CDbl(number(Selection(2)))))
            hc_string = HC_DB_Value + HC_DB_Temp
            
            If hc_string <> last_hc_string Then
                HC_Count = HC_Count + 1
                dbinput.HenrysConstantSource = get_source(nullcheck("RTI"))
                dbinput.HenrysConstant(HC_Count) = Selection(1)
                dbinput.HenrysConstantTemperature(HC_Count) = Selection(2)
            End If
     
        End If
     
        Selection.MoveNext
     
    Loop
    
    dbinput.NumberOfDatabaseHenrysConstants = HC_Count
     
    If dbinput.NumberOfDatabaseHenrysConstants = 0 Then

        '
        ' OPEN RECORDSET.
        '
        Set RS_Main = DB_Main.OpenRecordset( _
            "SELECT * FROM [HC Superfund] WHERE [HC Superfund].CAS = " & _
            Format$(dbinput.CasNumber, "0"))
        If (RS_Main.EOF = False) Then
          RS_Main.MoveFirst
          RS_Main.MoveLast
          RS_Main.MoveFirst
        End If
        Set Selection = RS_Main
        'contam_prop_form.Data1.RecordSource = "SELECT * FROM [HC Superfund] WHERE [HC Superfund].CAS = " & Format$(dbinput.CASNumber, "0")
        'contam_prop_form.Data1.Refresh
        'Set Selection = contam_prop_form.Data1.Recordset

        Do While Not Selection.EOF

            If number(Selection(1)) <> "" Then
                
                last_hc_string = hc_string
                HC_DB_Source = "SUPERFUND"
                hc_form!lblDatabase = HC_DB_Source
                LSet HC_DB_Value = Format$(number(Selection(1)), GetTheFormat(CDbl(number(Selection(1)))))
                HC_DB_Temp = Format$(number(Selection(2)), GetTheFormat(CDbl(number(Selection(2)))))
                hc_string = HC_DB_Value + HC_DB_Temp
                
                If hc_string <> last_hc_string Then
                    HC_Count = HC_Count + 1
                    dbinput.HenrysConstantSource = get_source(nullcheck("SUPERFUND"))
                    dbinput.HenrysConstant(HC_Count) = Selection(1)
                    dbinput.HenrysConstantTemperature(HC_Count) = Selection(2)
                End If
         
            End If
         
            Selection.MoveNext
         
        Loop
         
        dbinput.NumberOfDatabaseHenrysConstants = HC_Count
     
    End If

    If dbinput.NumberOfDatabaseHenrysConstants = 0 Then

        '
        ' OPEN RECORDSET.
        '
        Set RS_Main = DB_Main.OpenRecordset( _
            "SELECT * FROM [HC Yaws] WHERE [HC Yaws].CAS = " & _
            Format$(dbinput.CasNumber, "0"))
        If (RS_Main.EOF = False) Then
          RS_Main.MoveFirst
          RS_Main.MoveLast
          RS_Main.MoveFirst
        End If
        Set Selection = RS_Main
        'contam_prop_form.Data1.RecordSource = "SELECT * FROM [HC Yaws] WHERE [HC Yaws].CAS = " & Format$(dbinput.CASNumber, "0")
        'contam_prop_form.Data1.Refresh
        'Set Selection = contam_prop_form.Data1.Recordset

        Do While Not Selection.EOF

            If number(Selection(1)) <> "" Then
                last_hc_string = hc_string
                HC_DB_Source = "YAWS"
                hc_form!lblDatabase = HC_DB_Source
                LSet HC_DB_Value = Format$(number(Selection(1)), GetTheFormat(CDbl(number(Selection(1)))))
                HC_DB_Temp = Format$(number(Selection(2)), GetTheFormat(CDbl(number(Selection(2)))))
                hc_string = HC_DB_Value + HC_DB_Temp
                If hc_string <> last_hc_string Then
                    HC_Count = HC_Count + 1
                    dbinput.HenrysConstantSource = get_source(nullcheck("YAWS"))
                    dbinput.HenrysConstant(HC_Count) = Selection(1)
                    dbinput.HenrysConstantTemperature(HC_Count) = Selection(2)
                End If
            
            End If
         
        Selection.MoveNext
         
        Loop
        
        dbinput.NumberOfDatabaseHenrysConstants = HC_Count
     
     End If

     'Convert database Henry's constants to dimensionless units
     
     If dbinput.NumberOfDatabaseHenrysConstants > 0 Then
         Call HCDBCONV(dbinput.HenrysConstant(1), dbinput.HenrysConstantTemperature(1), dbinput.NumberOfDatabaseHenrysConstants, dbinput.HenrysConstantSource)
     End If

End Sub

