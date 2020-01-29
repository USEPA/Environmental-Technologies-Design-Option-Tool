Attribute VB_Name = "Refresh"
Option Explicit

Global ThisGraph As Object



Const Refresh_declarations_end = True


Sub refresh_frmMain()
Dim new_tag As Integer
Dim i As Integer

''  refresh_is_occurring = True
  If nowproj.Predicted_Available Then
    frmMain.picGraph.visible = True
    frmMain.ssframe_GraphHolder.visible = True
    frmMain.lblPlotAxes.visible = True
    frmMain.lblPlotSymbols.visible = True
    Call PlotResults(nowproj)
    Call ThisGraph.Refresh_Graph
  Else
    frmMain.picGraph.visible = False
    frmMain.ssframe_GraphHolder.visible = False
    frmMain.lblPlotAxes.visible = False
    frmMain.lblPlotSymbols.visible = False
  End If
      
  Call AssignTextAndTag(frmMain.txtData(0), nowproj.dyestudy_calcdate)
  'TRACK DOWN PLOT TYPE VALUE.
  new_tag = 0
  For i = 0 To frmMain.cboPlotType.ListCount - 1
    If (frmMain.cboPlotType.ItemData(i) = nowproj.plottype) Then
      new_tag = i
      Exit For
    End If
  Next i
  Call AssignTag_Scrollbox(frmMain.cboPlotType, new_tag)
  frmMain.cboPlotType.ListIndex = new_tag
''  refresh_is_occurring = False
  
  
End Sub
Private Sub PlotResults(proj As Project_Type)

Dim data_x() As Double
Dim data_y() As Double
Dim num_rows As Integer
Dim i As Integer
  
  '
  ' REMOVE ALL EXISTING GRAPH DATA.
  '
  Call ThisGraph.DeleteAllSeries
  '
  ' ADD THE FIRST SERIES.
  '
  Select Case proj.plottype
  
  Case 0
      num_rows = proj.Predicted_count
      ReDim data_x(1 To num_rows)
      ReDim data_y(1 To num_rows)
      
      For i = 1 To num_rows
        data_x(i) = proj.Predicted(i).Predicted_Theta: _
          data_y(i) = proj.Predicted(i).Predicted_E
      Next i
      
      Call ThisGraph.AddSeriesData( _
          "Series Whatever", CLng(num_rows), data_x, data_y, _
          0, 1#, QBColor(9))
  
  Case 1
      num_rows = proj.PredictedDispClosed_count
      ReDim data_x(1 To num_rows)
      ReDim data_y(1 To num_rows)

      For i = 1 To num_rows
        data_x(i) = proj.DispClosed(i).PredictedDispClosed_Theta: _
          data_y(i) = proj.DispClosed(i).PredictedDispClosed_E
      Next i
      
      Call ThisGraph.AddSeriesData( _
          "Series Whatever", CLng(num_rows), data_x, data_y, _
          0, 1#, QBColor(9))
  
  Case 2
      num_rows = proj.PredictedDispOpen_count
      ReDim data_x(1 To num_rows)
      ReDim data_y(1 To num_rows)

      For i = 1 To num_rows
        data_x(i) = proj.DispOpen(i).PredictedDispOpen_Theta: _
          data_y(i) = proj.DispOpen(i).PredictedDispOpen_E
      Next i
      
      Call ThisGraph.AddSeriesData( _
          "Series Whatever", CLng(num_rows), data_x, data_y, _
          0, 1#, QBColor(9))
  
  End Select
  
   num_rows = proj.Experimental_count
    ReDim data_x(1 To num_rows)
    ReDim data_y(1 To num_rows)
  
    For i = 1 To num_rows
      data_x(i) = proj.Experimental(i).Experimental_Theta: _
        data_y(i) = proj.Experimental(i).Experimental_E
    Next i
    
    Call ThisGraph.AddSeriesData( _
    "Series Whatever", CLng(num_rows), data_x, data_y, _
    1, 1#, QBColor(12))
  

End Sub



Sub refresh_frmDyeStudy(proj As Project_Type)
Dim i As Integer
  
  
  frmDyeStudy.f1book_dyestudy.visible = False
  frmDyeStudy.f1book_dyestudy.MaxRow = proj.dyestudy_count
  If proj.dyestudy_count <> 1600 Then
    For i = 1 To proj.dyestudy_count
        frmDyeStudy.f1book_dyestudy.EntryRC(i, 1) = _
            Trim$(proj.DyeStudy(i).time)
        frmDyeStudy.f1book_dyestudy.EntryRC(i, 2) = _
            Trim$(proj.DyeStudy(i).concentration)
    Next i
  End If
  frmDyeStudy.f1book_dyestudy.visible = True
  Call AssignTextAndTag(frmDyeStudy.txtData(0), proj.dyestudy_calcdate)

End Sub


