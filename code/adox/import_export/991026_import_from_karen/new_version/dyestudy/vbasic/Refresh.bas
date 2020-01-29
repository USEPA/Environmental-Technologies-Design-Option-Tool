Attribute VB_Name = "Refresh"
Option Explicit

Global ThisGraph As Object



Const Refresh_declarations_end = True


Sub refresh_frmMain()
  
  If nowproj.Predicted_Available Then
    frmMain.picGraph.visible = True
    frmMain.ssframe_GraphHolder.visible = True
    frmMain.lblPlotAxes.visible = True
    Call PlotResults(nowproj)
    Call ThisGraph.Refresh_Graph
  Else
    frmMain.picGraph.visible = False
    frmMain.ssframe_GraphHolder.visible = False
    frmMain.lblPlotAxes.visible = False
  End If
      
  Call AssignTextAndTag(frmMain.txtData(0), nowproj.dyestudy_calcdate)

End Sub
Private Sub PlotResults(proj As Project_Type)

Dim data_x() As Double
Dim data_y() As Double
Dim num_rows As Integer

  
  '
  ' REMOVE ALL EXISTING GRAPH DATA.
  '
  Call ThisGraph.DeleteAllSeries
  '
  ' ADD THE FIRST SERIES.
  '
  num_rows = proj.Predicted_count
  ReDim data_x(1 To num_rows)
  ReDim data_y(1 To num_rows)
  Dim i As Integer
  For i = 1 To num_rows
    data_x(i) = proj.Predicted(i).Predicted_Theta: _
      data_y(i) = proj.Predicted(i).Predicted_E
  Next i
  
  Call ThisGraph.AddSeriesData( _
      "Series Whatever", CLng(num_rows), data_x, data_y, _
      0, 1#, QBColor(9))
  '
  ' ADD THE SECOND SERIES.
  '
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
  If proj.dyestudy_count <> 400 Then
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


