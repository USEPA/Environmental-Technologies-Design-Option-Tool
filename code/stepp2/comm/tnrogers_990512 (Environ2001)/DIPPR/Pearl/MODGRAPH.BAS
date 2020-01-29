Attribute VB_Name = "modgraph"
Option Explicit

Sub Graph()

    Dim i As Integer
    Dim J As Integer
    Dim CNT As Integer
    Dim Increment As Double
    Dim CurTemp As Double
    Dim CurPt As Integer
    Dim CurSet As Integer
    Dim XData(10, 100) As Double
    Dim YData(10, 100) As Double
    Dim YDataFlag(10) As Double
    Dim FDupPointX(10) As Double
    Dim FDupPointY(10) As Double
    Dim LDupPointX(10) As Double
    Dim LDupPointY(10) As Double
    Dim FirstHalf As Boolean
    Dim DefaultUnit As String
    Dim TempTFT As Double
    Dim Prop As Integer
    Dim NumSets As Integer
    Dim LegendText(10) As String
    Dim YGraphType, XGraphType As Boolean
    Dim NumPoints As Integer
    Dim X() As Integer
    Screen.MousePointer = 11
        
    On Error Resume Next
'msh
   For i = 0 To 9
        For J = 0 To 99
            XData(i, J) = 0
            YData(i, J) = 0
        Next J
    Next i
'msh

    'Set property selected
    Prop = CurProp
            
    FRMGraph!GRHChem.chartType = VtChChartType2dXY  'set chart type to xy scatter-plot
    
    'set chart titles
    FRMGraph!GRHChem.Plot.Axis(0, 1).AxisTitle.Text = FRMGraphSet!TXTXAxis.Text
    FRMGraph!GRHChem.Plot.Axis(1, 1).AxisTitle.Text = FRMGraphSet!TXTYAxis.Text
    FRMGraph!GRHChem.TitleText = FRMGraphSet!TXTGraphTitle.Text
 
    'Set user selection for Y-axis type
    If FRMGraphSet!OPBYAxis(2).value = True Or FRMGraphSet!OPBYAxis(3).value = True Then
        YGraphType = True  '2 or 3 use log y's
    Else
        YGraphType = False '0 or 1 dont use log y's
    End If
    
    'Set user selection for X-axis type
    If FRMGraphSet!OPBYAxis(2).value = True Or FRMGraphSet!OPBYAxis(3).value = True Then
        XGraphType = True  '2 or 3 use  log x's
    Else
        XGraphType = False '0 or 1 dont use log's
    End If
 
    'Set the number of points to graph from user data
'msh     NumPoints = FRMGraphSet!TXTNumPts
    NumPoints = FRMGraphSet!TXTNumPts + 1
     
    'Set how often to label points
'    If FRMGraph!GRHChem.NumPoints > 5 Then
'       FRMGraph!GRHChem.LabelEvery = 2
'    Else
'       FRMGraph!GRHChem.LabelEvery = 1
'    End If
                
    'Determine temperature increment based on user temperature range
'msh     Increment = (Val(FRMGraphSet!TXTMaxT.Text) - Val(FRMGraphSet!TXTMinT.Text)) / (Val(FRMGraphSet!TXTNumPts.Text) - 1)
    Increment = (Val(FRMGraphSet!TXTMaxT.Text) - Val(FRMGraphSet!TXTMinT.Text)) / (Val(FRMGraphSet!TXTNumPts.Text))
    
    'Determine selected property method
    Select Case Trim(FRMGraphSet!CMBPropMethod.Text)
        Case "801 Database"
            CNT = 1
        Case "911 Database"
            CNT = 2
        Case "User Input"
            CNT = 10
    End Select
    
    'Reinitialize X and Y data matrices
    For CurPt = 1 To 10
        For CurSet = 1 To 100
            XData(CurSet, CurPt) = 0
            YData(CurSet, CurPt) = 0
        Next CurSet
    Next CurPt
    
    NumSets = 0
    
    'Calculate X and Y data to graph
    For i = 0 To FRMGraphSet!LSTUserList.ListCount - 1
        If FRMGraphSet!LSTUserList.Selected(i) = True Then
            NumSets = NumSets + 1
            FRMMain!Data1.Recordset.FindFirst "Name = '" & Trim(FRMGraphSet!LSTUserList.List(i)) & "'"

            'Load in data for current chemical
            If FRMMain!Data1.Recordset("CAS") <> Cur_Info.CAS Then
                Cur_Info.CAS = FRMMain!Data1.Recordset("CAS")
                Call Recalculate
                CurProp = Prop
            End If
            
            'Set legend text for this set to current chemical name
            LegendText(NumSets) = Trim(FRMMain!Data1.Recordset("Name"))
            
            'Store current TFT
            TempTFT = InfoMethod(Prop).TFT
            
            'Find default Unit
            DefaultUnit = Get_DefaultUnit(Prop)
             
            'Set the X and Y axis data
            CurTemp = FRMGraphSet!TXTMinT.Text
            For CurPt = 1 To NumPoints
                 If Trim(FRMGraphSet!LBLTFTUnit(0).caption) = "K" Then
                     InfoMethod(Prop).TFT = CurTemp
                 Else
                     InfoMethod(Prop).TFT = Convert(CurTemp, CLng(Prop), Trim(FRMGraphSet!LBLTFTUnit(0).caption), "K", False)
                 End If
                 YData(NumSets, CurPt) = CalcFofT(Prop, CNT)
                 XData(NumSets, CurPt) = CurTemp
                'Convert units to user selection
                 If YData(NumSets, CurPt) <> 0 Then
                     If Trim(FRMGraphSet!CMBPropUnits.Text) <> DefaultUnit Then
                         YData(NumSets, CurPt) = Convert(YData(NumSets, CurPt), OptPress, Trim(DefaultUnit), Trim(FRMGraphSet!CMBPropUnits.Text), False) 'FIX : I don't know what exactly is going on so -2 is a dummy code
                     End If
                     If FRMGraphSet!OPBYAxis(1) = True Then  '1/y
                         YData(NumSets, CurPt) = 1 / YData(NumSets, CurPt)
                     ElseIf FRMGraphSet!OPBYAxis(2) = True Then 'log y
                         YData(NumSets, CurPt) = YData(NumSets, CurPt)
                     ElseIf FRMGraphSet!OPBYAxis(3) = True Then 'log 1/y
                         YData(NumSets, CurPt) = 1 / YData(NumSets, CurPt)
                     End If
                 End If
                 If XData(NumSets, CurPt) <> 0 Then
                     If FRMGraphSet!OPBXAxis(1) = True Then  '1/y
                         XData(NumSets, CurPt) = 1 / CurTemp
                     ElseIf FRMGraphSet!OPBXAxis(2) = True Then 'log y
                         XData(NumSets, CurPt) = log10(CurTemp)
                     ElseIf FRMGraphSet!OPBXAxis(3) = True Then 'log 1/y
                         XData(NumSets, CurPt) = log10(1 / CurTemp)
                     End If
                 End If
                 CurTemp = CurTemp + Increment
             Next CurPt
             'Restore TFT to original value
             InfoMethod(Prop).TFT = TempTFT
        End If
    Next i
         
    'Determine last duplicate point for each set
     For CurSet = 1 To NumSets
         For CurPt = 1 To NumPoints
             If YData(CurSet, CurPt) <> 0 Then
                YDataFlag(CurSet) = True
                 LDupPointY(CurSet) = YData(CurSet, CurPt)
                 LDupPointX(CurSet) = XData(CurSet, CurPt)
             End If
         Next CurPt
     Next CurSet
    
    'Determine first duplicate point for each set
     For CurSet = 1 To NumSets
         CurPt = 1
         FDupPointY(CurSet) = 0
         Do Until FDupPointY(CurSet) <> 0 Or CurPt > NumPoints
             If YData(CurSet, CurPt) <> 0 Then
                 FDupPointY(CurSet) = YData(CurSet, CurPt)
                 FDupPointX(CurSet) = XData(CurSet, CurPt)
             End If
             CurPt = CurPt + 1
         Loop
     Next CurSet
        
    'Determine the number of sets
     i = 0
     For CurSet = 1 To NumSets
         If YDataFlag(CurSet) = True Then i = i + 1
     Next CurSet
    

    'Rebuild x and y data
     For CurSet = 1 To i
         FirstHalf = False
         For CurPt = 1 To NumPoints
             If YDataFlag(CurSet) = True Then
                 If YData(CurSet, CurPt) = 0 Then
                     If FirstHalf = False Then
                         YData(CurSet, CurPt) = FDupPointY(CurSet)
                         XData(CurSet, CurPt) = FDupPointX(CurSet)
                     Else
                         YData(CurSet, CurPt) = LDupPointY(CurSet)
                         XData(CurSet, CurPt) = LDupPointX(CurSet)
                     End If
                 Else
                     FirstHalf = True
                 End If
             End If
         Next CurPt
     Next CurSet
     
    FRMGraph!GRHChem.ShowLegend = True
    FRMGraph!GRHChem.Legend.Location.LocationType = VtChLocationTypeRight

'********************************************
' MSH 10/19/98 - Commented out old graphing function.
'                Added correct functionality written by Paul.
'
'    'resize array
'    ReDim X(NumSets, NumPoints)
'     FRMGraph!GRHChem.ColumnCount = NumSets
'
'    'get the graph data
'     For CurSet = 1 To NumSets
'         For CurPt = 1 To NumPoints
'            FRMGraph!GRHChem.ThisPoint = CurPt
'            If YDataFlag(CurSet) = True Then
''                FRMGraph!GRHChem.Legend.TextLayout = LegendText(CurSet)
'                X(CurPt, CurSet) = YData(CurSet, CurPt)
'                X(CurPt, 0) = XData(CurSet, CurPt)
'           End If
'         Next CurPt
'     Next
'
'  FRMGraph!GRHChem.ChartData = X
' ' FRMGraph!GRHChem.Repaint = True
'**************************************************
'Dim q1
'q1 = 2 * NumSets
'FRMGraph!GRHChem.Plot.SeriesCollection
ReDim X(NumSets, NumPoints)
FRMGraph!GRHChem.ColumnCount = NumSets * 2
FRMGraph!GRHChem.RowCount = NumPoints
FRMGraph!GRHChem.Plot.UniformAxis = False

' get the graph data
'***************************************************
'msh 2/9/99
'
'the new syntax below calculates the legent text at runtime...oppose to before
'when the chemical names were truncated beyond view of the user.
'***************************************************
    For CurSet = 1 To NumSets
'msh 2/9/99 New Line:
        FRMGraph!GRHChem.Plot.SeriesCollection(2 * CurSet - 1).LegendText = LegendText(CurSet)
        For CurPt = 1 To NumPoints
            FRMGraph!GRHChem.Row = CurPt
            FRMGraph!GRHChem.Column = CurSet * 2 - 1
'msh 2/9/99 FRMGraph!GRHChem.ColumnLabel = LegendText(CurSet)
            FRMGraph!GRHChem.Data = XData(CurSet, CurPt)
            FRMGraph!GRHChem.Column = CurSet * 2
            FRMGraph!GRHChem.Data = YData(CurSet, CurPt)
        Next CurPt
    Next CurSet
      
'Set graph to show data as lines, symbols, or both
If FRMGraphSet!OPBLineSym(0).value = True Then 'lines only
     For i = 0 To FRMGraph!GRHChem.Plot.SeriesCollection.count
        FRMGraph!GRHChem.Plot.SeriesCollection(i).ShowLine = True
        FRMGraph!GRHChem.Plot.SeriesCollection(i).SeriesMarker.Show = False
    Next
ElseIf FRMGraphSet!OPBLineSym(1).value = True Then 'symbols only
     For i = 0 To FRMGraph!GRHChem.Plot.SeriesCollection.count
        FRMGraph!GRHChem.Plot.SeriesCollection(i).SeriesMarker.Show = True
        FRMGraph!GRHChem.Plot.SeriesCollection(i).ShowLine = False
    Next
Else    'show both
     For i = 0 To FRMGraph!GRHChem.Plot.SeriesCollection.count
        FRMGraph!GRHChem.Plot.SeriesCollection(i).SeriesMarker.Show = True
        FRMGraph!GRHChem.Plot.SeriesCollection(i).ShowLine = True
    Next
End If

Screen.MousePointer = 1

End Sub



Public Sub fill_range_form(range_prop As Integer)

Dim unitto As String
Dim unitfrom As String
Dim i As Integer
Dim J As Integer
Dim range_method As Integer
Dim mintemp As Double
Dim maxtemp As Double
Dim converted_temp As Double
Dim this_cas As Long

unitfrom = "K"  ' double check this is always K DENISE check
unitto = Trim(FRMGraphSet!LBLTFTUnit(1).caption)

' first clear the grid
'On Error Resume Next
'For I = 1 To FRMInfo!rangegrd.Rows - 1
 '   FRMInfo!rangegrd.Col = 1
  '  If Len(FRMInfo!rangegrd.Text) > 2 Then
   '     FRMInfo!rangegrd.RemoveItem (I)
  '  End If
'Next I

'FRMInfo.rangegrd.Rows = FRMGraphSet!LSTUserList.ListCount + 1
FRMInfo!rangegrd.Row = 0
FRMInfo!rangegrd.Col = 0
FRMInfo!rangegrd.Text = " "
FRMInfo!rangegrd.Col = 1
FRMInfo!rangegrd.Text = "CAS"
FRMInfo!rangegrd.Col = 2
FRMInfo!rangegrd.Text = "Name"
FRMInfo!rangegrd.Col = 3
FRMInfo!rangegrd.Text = "Minimum"
FRMInfo!rangegrd.Col = 4
FRMInfo!rangegrd.Text = "Maximum"
FRMInfo!rangegrd.ColWidth(0) = 200
FRMInfo!rangegrd.ColWidth(1) = 1000
FRMInfo!rangegrd.ColWidth(2) = 3500
FRMInfo!rangegrd.ColWidth(3) = 1000
FRMInfo!rangegrd.ColWidth(4) = 1000
i = 1
FRMMain!Data2.Recordset.MoveFirst
Do While Not FRMMain!Data2.Recordset.EOF
    FRMInfo!rangegrd.Row = i
    FRMInfo!rangegrd.Col = 0
    FRMInfo!rangegrd.Text = "X"
    FRMInfo!rangegrd.Col = 1
    FRMInfo!rangegrd.Text = FRMMain!Data2.Recordset("CAS")
    this_cas = CLng(Trim(FRMInfo!rangegrd.Text))
    FRMInfo!rangegrd.Col = 2
    FRMInfo!rangegrd.Text = FRMMain!Data2.Recordset("Name")
    
    ' need to recalculate?  hope not
    ' get the min and max temp for this chemical
    Call get_min_max(this_cas, mintemp, maxtemp)
    FRMInfo!rangegrd.Col = 3
    If mintemp <> -99999# Then
        If unitto <> unitfrom Then
            converted_temp = Convert(mintemp, OptTemp, unitfrom, unitto, False)
            FRMInfo!rangegrd.Text = CStr(converted_temp)
        End If
        FRMInfo!rangegrd.Text = CStr(mintemp)
    Else
        FRMInfo!rangegrd.Text = "na"
    End If
    
    FRMInfo!rangegrd.Col = 4
    If maxtemp <> 99999# Then
        If unitto <> unitfrom Then
            converted_temp = Convert(maxtemp, OptTemp, unitfrom, unitto, False)
            FRMInfo!rangegrd.Text = CStr(converted_temp)
        End If
        FRMInfo!rangegrd.Text = CStr(maxtemp)
    Else
        FRMInfo!rangegrd.Text = "na"
    End If
    
    FRMMain!Data2.Recordset.MoveNext
    i = i + 1
Loop
FRMInfo!rangegrd.Rows = FRMInfo!rangegrd.Rows
FRMInfo!Option1(0) = True
FRMInfo.mintbx.Text = FRMGraphSet!TXTMinT.Text
FRMInfo.maxtbx.Text = FRMGraphSet!TXTMaxT.Text
FRMInfo.unitlbl.caption = FRMGraphSet!LBLTFTUnit(1).caption
FRMInfo.rangepnl.caption = FRMGraphSet!CMBPropFunction.Text
FRMInfo!rangegrd.Refresh
End Sub

Public Sub get_min_max(casno As Long, min As Double, max As Double)

    ' a function which finds the temperature range for
    ' a chemical for a particular f(t) property
    ' returns the min and max as arguments
    ' first initialize these in case we don't find anything
    Dim DBTbl As Recordset
    min = -99999#
    max = 99999#
    
    Select Case Trim(FRMGraphSet!CMBPropMethod.Text)
        Case "911 Database"
            On Error Resume Next
            If Len(Path911) < 4 Then
                min = -99999#
                max = 99999#
                Exit Sub
            End If
            Set DBTbl = DBJetMaster.OpenRecordset("DIPPR911", dbOpenTable)
    
            DBTbl.Index = "PrimaryKey1"
            DBTbl.Seek "=", casno
    
            If DBTbl.NoMatch Then
                DBTbl.Close
                Exit Sub
            End If
            While DBTbl("PEARLS Code") <> CurProp
                DBTbl.MoveNext
            Wend
            If Not DBTbl.EOF Then
                min = DBTbl("Value")
                max = DBTbl("Temperature")
            End If
            DBTbl.Close
            Exit Sub
                
        Case "801 Database"
        
            If Len(Path801) < 4 Then    ' null path
                Exit Sub
            End If
            Set DBTbl = DBJetMaster.OpenRecordset("DIPPR801", dbOpenTable)
            DBTbl.Index = "PrimaryKey"
            DBTbl.Seek "=", casno
    
            If DBTbl.NoMatch Then
                DBTbl.Close
                Exit Sub
            End If
            Select Case CurProp
                Case LD
                   min = DBTbl("LDNTMIN")
                   max = DBTbl("LDNTMAX")
                Case VP
                    min = DBTbl("VPTMIN")
                    max = DBTbl("VPTMAX")
                Case Hvap
                    min = DBTbl("HVPTMIN")
                    max = DBTbl("HVPTMAX")
                Case LHC
                    min = DBTbl("LCPTMIN")
                    max = DBTbl("LCPTMAX")
                Case VHC
                    min = DBTbl("ICPTMIN")
                    max = DBTbl("ICPTMAX")
                Case LV
                    min = DBTbl("LVSTMIN")
                    max = DBTbl("LVSTMAX")
                Case VV
                    min = DBTbl("VVSTMIN")
                    max = DBTbl("VVSTMAX")
                Case LTC
                    min = DBTbl("LTCTMIN")
                    max = DBTbl("LTCTMAX")
                Case VTC
                    min = DBTbl("VTCTMIN")
                    max = DBTbl("VTCTMAX")
                Case ST
                    min = DBTbl("STTMIN")
                    max = DBTbl("STTMAX")
            End Select
            DBTbl.Close
        Case "User Input"
            
            Set DBTbl = DBJetUser.OpenRecordset("SaveTable1", dbOpenTable)
            DBTbl.Index = "PrimaryKey"
            DBTbl.Seek "=", casno
    
            If DBTbl.NoMatch Then
                DBTbl.Close
                Exit Sub
            End If
            DBTbl.Close
            Set DBTbl = DBJetUser.OpenRecordset("SaveTable2", dbOpenTable)
            DBTbl.Index = "PrimaryKey"
            DBTbl.Seek "=", casno
            While DBTbl("Current Method Index") <> CurProp
                DBTbl.MoveNext
            Wend
            If DBTbl.EOF Then
                DBTbl.Close
                Exit Sub
            End If
            min = DBTbl("User MinT")
            max = DBTbl("User MaxT")
            DBTbl.Close
        End Select
End Sub
