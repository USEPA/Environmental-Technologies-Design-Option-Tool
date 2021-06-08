VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GraphControl"
Attribute VB_Creatable = False
Attribute VB_Exposed = False


Private ctlPicture As Control

'Private title_main As String
'Private title_x As String
'Private title_y As String

Private completely_new_graph As Integer     'not really used effectively...

Private xguide1 As Double
Private xguide2 As Double
Private xguide3 As Double
Private yguide1 As Double
Private yguide2 As Double
Private yguide3 As Double

'Private title_main_width As Double
'Private title_main_height As Double
'Private title_x_width As Double
'Private title_x_height As Double
'Private title_y_width As Double
'Private title_y_height As Double
'Private title_main_left As Double
'Private title_main_top As Double
  
Private xmin_specified As Integer
Private xmax_specified As Integer
Private ymin_specified As Integer
Private ymax_specified As Integer
Private xmin_user As Double
Private xmax_user As Double
Private ymin_user As Double
Private ymax_user As Double
Private xmin_current As Double
Private xmax_current As Double
Private ymin_current As Double
Private ymax_current As Double
  
Dim xaxis_zone_height As Double
Dim yaxis_zone_width As Double
Dim series_zone_width As Double

Private xticks_specified As Integer
Private yticks_specified As Integer
Private xticks_user As Double
Private yticks_user As Double
Private xticks_current As Double
Private yticks_current As Double
  
Private Type tick_numbers_type
  font As String
  font_size As Integer
  font_bold As Integer
  visible As Integer
End Type
Private xticks_numbers As tick_numbers_type
Private yticks_numbers As tick_numbers_type

Private Type graph_title_type
  title As String
  top As Double
  left As Double
  width As Double
  height As Double
  font As String
  font_size As Integer
  font_bold As Integer
  visible As Integer
  number_format As String       'A BLANK STRING INDICATES NON-FORMATTED.
End Type
Private title_main As graph_title_type
Private title_x As graph_title_type
Private title_y As graph_title_type

Private Type data_series_type
  title_series As String
  num_points As Long
  symbol_style As Integer
  symbol_size_ratio As Double
  color As Long
  data_x() As Double
  data_y() As Double
End Type

Private number_of_series As Integer
Private data_series(1 To 10) As data_series_type

Const MARGIN_TITLES_X = 150
Const MARGIN_TITLES_Y = 100
Const YAXIS_TICK_WIDTH = 100
Const XAXIS_TICK_HEIGHT = 100

Private Type Rect_Type
  x1 As Double
  y1 As Double
  x2 As Double
  y2 As Double
End Type






Const GraphControl_declarations_end = True


'RETURNS:
'- Index of the series that was added
Public Function AddSeriesData( _
    new_title_series As String, _
    new_num_points As Long, _
    new_data_x() As Double, _
    new_data_y() As Double, _
    new_symbol_style As Integer, _
    new_symbol_size_ratio As Double, _
    new_color As Long) As Integer
Dim this_series_index As Integer

  this_series_index = number_of_series + 1
  number_of_series = number_of_series + 1
  
  data_series(this_series_index).title_series = new_title_series
  data_series(this_series_index).num_points = new_num_points
  data_series(this_series_index).symbol_style = new_symbol_style
  data_series(this_series_index).color = new_color
  data_series(this_series_index).symbol_size_ratio = new_symbol_size_ratio

  ReDim data_series(this_series_index).data_x(1 To new_num_points)
  ReDim data_series(this_series_index).data_y(1 To new_num_points)
  For i = 1 To new_num_points
    data_series(this_series_index).data_x(i) = new_data_x(i)
    data_series(this_series_index).data_y(i) = new_data_y(i)
  Next i
  
  AddSeriesData = this_series_index

End Function
Public Sub DeleteAllSeries()
  number_of_series = 0
End Sub


Private Sub calculate_automatic_xy_ranges_and_ticks( _
    Use_Widest_X_Text As Double)
Dim extreme_lowest_x_rounded As Double
Dim extreme_highest_x_rounded As Double
Dim extreme_lowest_y_rounded As Double
Dim extreme_highest_y_rounded As Double
Dim extreme_lowest_x As Double
Dim extreme_highest_x As Double
Dim extreme_lowest_y As Double
Dim extreme_highest_y As Double
Dim any_x_less_than_zero As Integer
Dim any_x_greater_than_zero As Integer
Dim any_y_less_than_zero As Integer
Dim any_y_greater_than_zero As Integer

Dim i As Integer
Dim j As Integer

Dim xticks_automated As Double
Dim yticks_automated As Double
Dim xticks_max_available As Integer
Dim yticks_max_available As Integer
Dim dx_guess As Double
Dim dy_guess As Double
Dim this_guess As Double

  'DETERMINE LOWEST AND HIGHEST VALUES OF X AND Y:
  extreme_lowest_x = 1E+200
  extreme_highest_x = -1E+200
  extreme_lowest_y = 1E+200
  extreme_highest_y = -1E+200
  any_x_less_than_zero = False
  any_x_greater_than_zero = False
  any_y_less_than_zero = False
  any_y_greater_than_zero = False
  For i = 1 To number_of_series
    For j = 1 To data_series(i).num_points
      If (data_series(i).data_x(j) < extreme_lowest_x) Then extreme_lowest_x = data_series(i).data_x(j)
      If (data_series(i).data_x(j) > extreme_highest_x) Then extreme_highest_x = data_series(i).data_x(j)
      If (data_series(i).data_y(j) < extreme_lowest_y) Then extreme_lowest_y = data_series(i).data_y(j)
      If (data_series(i).data_y(j) > extreme_highest_y) Then extreme_highest_y = data_series(i).data_y(j)
      If (data_series(i).data_x(j) < 0#) Then any_x_less_than_zero = True
      If (data_series(i).data_x(j) > 0#) Then any_x_greater_than_zero = True
      If (data_series(i).data_y(j) < 0#) Then any_y_less_than_zero = True
      If (data_series(i).data_y(j) > 0#) Then any_y_greater_than_zero = True
    Next j
  Next i
  
  If (number_of_series = 0) Then
    xticks_numbers.visible = False
    yticks_numbers.visible = False
    xticks_automated = 0#
    yticks_automated = 0#
    xticks_current = xticks_automated
    yticks_current = yticks_automated
    xmax_current = 0#
    xmin_current = 0#
    ymax_current = 0#
    xmin_current = 0#
    Exit Sub
  End If
  
  'FOR A GIVEN AXIS, IF THE DATA DOES NOT CROSS THAT AXIS, SET EITHER THE
  'MINIMUM OR MAXIMUM VALUE EQUAL TO 0; THIS WILL ALLOW THE AXIS TO
  'BE DISPLAYED ON THE PLOT.
  If (any_x_less_than_zero = False) Then extreme_lowest_x = 0#
  If (any_x_greater_than_zero = False) Then extreme_highest_x = 0#
  If (any_y_less_than_zero = False) Then extreme_lowest_y = 0#
  If (any_y_greater_than_zero = False) Then extreme_highest_y = 0#
  
  'FOR A GIVEN AXIS, IF EITHER OF THE COORDINATES IS EQUAL FOR ALL
  'POINTS IN THE SET OF SERIES, ADD A 1.0 TO ITS MAXIMUM.
  'THIS ENSURES THAT VARIOUS CALCULATION ERRORS DO NOT OCCUR!
  If (extreme_lowest_x = extreme_highest_x) Then
    extreme_highest_x = extreme_highest_x + 1#
  End If
  If (extreme_lowest_y = extreme_highest_y) Then
    extreme_highest_y = extreme_highest_y + 1#
  End If
  
  'CALCULATE MAX AVAILABLE TICKS ON X AND Y AXIS and AUTOMATED TICK CONFIGURATION
  yticks_numbers.visible = True
  ctlPicture.font = yticks_numbers.font
  ctlPicture.font.Size = yticks_numbers.font_size
  ctlPicture.font.Bold = yticks_numbers.font_bold
  yticks_max_available = (yguide2 - yguide1) / (2# * CDbl(ctlPicture.TextHeight("12345.67890")))
  If (yticks_max_available <= 0) Then
    yticks_max_available = 1#
    'yticks_numbers.visible = False
  End If
  dy_guess = (extreme_highest_y - extreme_lowest_y) / CDbl(yticks_max_available)
  yticks_automated = calculate_roundup_to_125(dy_guess)
  
  xticks_numbers.visible = True
  ctlPicture.font = xticks_numbers.font
  ctlPicture.font.Size = xticks_numbers.font_size
  ctlPicture.font.Bold = xticks_numbers.font_bold
  'xticks_max_available = _
      (xguide3 - xguide2) / _
      (1.2 * CDbl(ctlPicture.TextWidth("0.0010")))
  xticks_max_available = _
      (xguide3 - xguide2) / _
      (1.2 * Use_Widest_X_Text)
  If (xticks_max_available <= 0) Then
    xticks_max_available = 1#
    'xticks_numbers.visible = False
  End If
  dx_guess = (extreme_highest_x - extreme_lowest_x) / CDbl(xticks_max_available)
  xticks_automated = calculate_roundup_to_125(dx_guess)
  
  'SET X/Y TICK INTERVAL TO EITHER (1) USER OR (2) AUTOMATED VALUES.
  If (xticks_specified) Then
    'USER SPECIFIED XTICKS:
    xticks_current = xticks_user
  Else
    'AUTOMATED XTICKS:
    xticks_current = xticks_automated
  End If
  If (yticks_specified) Then
    'USER SPECIFIED YTICKS:
    yticks_current = yticks_user
  Else
    'AUTOMATED YTICKS:
    yticks_current = yticks_automated
  End If

  'CALCULATE AUTOMATED X/Y MAX AND MIN.
  '---- X MAXIMUM:
  i = 0
  Do While (1 = 1)
    this_guess = CDbl(i) * xticks_automated
    If (this_guess >= extreme_highest_x) Then
      extreme_highest_x_rounded = this_guess
      Exit Do
    End If
    i = i + 1
  Loop
  '---- X MINIMUM:
  i = 0
  Do While (1 = 1)
    this_guess = -CDbl(i) * xticks_automated
    If (this_guess <= extreme_lowest_x) Then
      extreme_lowest_x_rounded = this_guess
      Exit Do
    End If
    i = i + 1
  Loop
  '---- Y MAXIMUM:
  i = 0
  Do While (1 = 1)
    this_guess = CDbl(i) * yticks_automated
    If (this_guess >= extreme_highest_y) Then
      extreme_highest_y_rounded = this_guess
      Exit Do
    End If
    i = i + 1
  Loop
  '---- Y MINIMUM:
  ''''i = 0     'KEEP i AS PREVIOUS VALUE
  Do While (1 = 1)
    ''''this_guess = -CDbl(i) * yticks_automated
    this_guess = CDbl(i) * yticks_automated
    If (this_guess <= extreme_lowest_y) Then
      extreme_lowest_y_rounded = this_guess
      Exit Do
    End If
    ''''i = i + 1
    i = i - 1
  Loop
  
  ''LATER: IMPROVE THIS TO DO ACTUAL ROUNDING!
  'extreme_lowest_x_rounded = extreme_lowest_x
  'extreme_highest_x_rounded = extreme_highest_x
  'extreme_lowest_y_rounded = extreme_lowest_y
  'extreme_highest_y_rounded = extreme_highest_y
  
  'SET X/Y MAX AND MIN TO EITHER (1) USER OR (2) AUTOMATED VALUES.
  If (xmin_specified) Then
    'USER SPECIFIED XMIN:
    xmin_current = xmin_user
  Else
    'AUTOMATED XMIN:
    'If (any_x_less_than_zero) Then
      xmin_current = extreme_lowest_x_rounded
    'Else
    '  xmin_current = 0#
    'End If
  End If
  If (xmax_specified) Then
    'USER SPECIFIED XMAX:
    xmax_current = xmax_user
  Else
    'AUTOMATED XMAX:
    'If (any_x_greater_than_zero) Then
      xmax_current = extreme_highest_x_rounded
    'Else
    '  xmax_current = 0#
    'End If
  End If
  If (ymin_specified) Then
    'USER SPECIFIED YMIN:
    ymin_current = ymin_user
  Else
    'AUTOMATED YMIN:
    'If (any_y_less_than_zero) Then
      ymin_current = extreme_lowest_y_rounded
    'Else
    '  ymin_current = 0#
    'End If
  End If
  If (ymax_specified) Then
    'USER SPECIFIED YMAX:
    ymax_current = ymax_user
  Else
    'AUTOMATED YMAX:
    'If (any_y_greater_than_zero) Then
      ymax_current = extreme_highest_y_rounded
    'Else
    '  ymax_current = 0#
    'End If
  End If

End Sub


Private Sub calculate_guides()
Dim Have_Adjusted_Use_Widest_X_Text As Boolean
Dim Have_Adjusted_Use_Widest_Y_Text As Boolean
Dim Use_Widest_X_Text As Double
Dim Use_Widest_Y_Text As Double
Dim Any_Collision As Boolean
Dim Widest_X_Text As Double
Dim Widest_Y_Text As Double
  
  Use_Widest_X_Text = ctlPicture.TextWidth("0.0010")
  Use_Widest_Y_Text = ctlPicture.TextWidth("0.0010")
  Have_Adjusted_Use_Widest_X_Text = False
  Have_Adjusted_Use_Widest_Y_Text = False
  
  'CALCULATE TITLE SIZES AND POSITIONS.
  ctlPicture.font = title_main.font
  ctlPicture.font.Size = title_main.font_size
  ctlPicture.font.Bold = title_main.font_bold
  title_main.width = ctlPicture.TextWidth(title_main.title)
  title_main.height = ctlPicture.TextHeight(title_main.title)
  If (Len(title_main.title) = 0) Then
    title_main.visible = False
  Else
    title_main.visible = True
  End If
  
  ctlPicture.font = title_x.font
  ctlPicture.font.Size = title_x.font_size
  ctlPicture.font.Bold = title_x.font_bold
  title_x.width = ctlPicture.TextWidth(title_x.title)
  title_x.height = ctlPicture.TextHeight(title_x.title)
  If (Len(title_x.title) = 0) Then
    title_x.visible = False
  Else
    title_x.visible = True
  End If
  
  ctlPicture.font = title_y.font
  ctlPicture.font.Size = title_y.font_size
  ctlPicture.font.Bold = title_y.font_bold
  title_y.width = ctlPicture.TextWidth(title_y.title)
  title_y.height = ctlPicture.TextHeight(title_y.title)
  If (Len(title_y.title) = 0) Then
    title_y.visible = False
  Else
    title_y.visible = True
  End If
  
recalc_after_titles:
  'TEMPORARY (???) -- IS THIS THE BEST POSSIBLE METHOD?
  ctlPicture.font = xticks_numbers.font
  ctlPicture.font.Size = xticks_numbers.font_size
  ctlPicture.font.Bold = xticks_numbers.font_bold
  'xaxis_zone_height = MARGIN_TITLES_Y + ctlPicture.TextHeight("12345.67890")
  xaxis_zone_height = MARGIN_TITLES_Y * 3 + ctlPicture.TextHeight("12345.67890")
  
  ctlPicture.font = yticks_numbers.font
  ctlPicture.font.Size = yticks_numbers.font_size
  ctlPicture.font.Bold = yticks_numbers.font_bold
  yaxis_zone_width = MARGIN_TITLES_X + Use_Widest_Y_Text
  
  'series_zone_width = 1000
  series_zone_width = 0
  
  'CALCULATE GUIDES
  xguide1 = (MARGIN_TITLES_X + title_y.width + MARGIN_TITLES_X) * (-1#) * (title_y.visible)
  xguide2 = xguide1 + yaxis_zone_width
  'xguide3 = ctlPicture.width - series_zone_width
  xguide3 = ctlPicture.width - MARGIN_TITLES_X * 3
  yguide1 = MARGIN_TITLES_Y + (title_main.height + MARGIN_TITLES_Y) * (-1#) * (title_main.visible)
  yguide3 = ctlPicture.height - (MARGIN_TITLES_Y + title_x.height + MARGIN_TITLES_Y) * (-1#) * (title_x.visible)
  yguide2 = yguide3 - xaxis_zone_height

  'CALCULATE X AND Y AXIS RANGES AND TICKS
  Call calculate_automatic_xy_ranges_and_ticks(Use_Widest_X_Text)
  
  'DETERMINE TITLE POSITIONS
  title_main.left = xguide1 + (xguide3 - xguide1) / 2# - title_main.width / 2#
  title_main.top = yguide1 / 2# - title_main.height / 2#
  
  title_x.left = xguide1 + (xguide3 - xguide1) / 2# - title_x.width / 2#
  title_x.top = yguide3 + (ctlPicture.height - yguide3) / 2# - title_x.height / 2#
  
  title_y.left = (xguide1) / 2# - title_y.width / 2#
  title_y.top = yguide1 + (yguide3 - yguide1) / 2# - title_y.height / 2#
  
  'RECALCULATE GUIDES IF X TEXT IS OVERLY SQUEEZED.
  If (Have_Adjusted_Use_Widest_X_Text = False) Then
    Call draw_axis_x(Any_Collision, False, Widest_X_Text)
    If (Any_Collision) Then
      Use_Widest_X_Text = Widest_X_Text * 1.2
      Have_Adjusted_Use_Widest_X_Text = True
      GoTo recalc_after_titles
    End If
  End If
  
  'DETERMINE WIDEST Y TEXT.
  If (Have_Adjusted_Use_Widest_Y_Text = False) Then
    Call draw_axis_y(Any_Collision, False, Widest_Y_Text)
    If (Widest_Y_Text > Use_Widest_Y_Text) Then
      Have_Adjusted_Use_Widest_Y_Text = True
      Use_Widest_Y_Text = Widest_Y_Text * 1.2
      GoTo recalc_after_titles
    End If
  End If
  
  'RECALCULATE GUIDES IF OVERLY SQUEEZED
  If (xguide3 - xguide2 < 1000#) And (title_y.visible) Then
    title_y.visible = False
    GoTo recalc_after_titles
  End If
  If (yguide2 - yguide1 < 1000#) And (title_x.visible) Then
    title_x.visible = False
    GoTo recalc_after_titles
  End If
  If (yguide2 - yguide1 < 1000#) And (title_main.visible) Then
    title_main.visible = False
    GoTo recalc_after_titles
  End If

  'If xguide3-xguide2 < 1000 then remove Y axis title and recalc.
  'If yguide2-yguide1 < 1000 then remove X axis title and recalc.
  


End Sub


Private Function calculate_roundup_to_125(x As Double) As Double
Dim alpha As Double
Dim beta As Double
Dim gamma As Double

  If (x <= 0#) Then
    calculate_roundup_to_125 = 0#
    Exit Function
  End If
  alpha = Int(Log(x) / Log(10))
  beta = x / (10# ^ alpha)
  If (beta <= 1#) Then gamma = 1#
  If ((beta > 1#) And (beta <= 2#)) Then gamma = 2#
  If ((beta > 2#) And (beta <= 5#)) Then gamma = 5#
  If ((beta > 5#) And (beta <= 10#)) Then gamma = 10#
  calculate_roundup_to_125 = gamma * (10# ^ alpha)

End Function

Public Sub CreateGraph(new_title_main As String, new_title_x As String, new_title_y As String)
Dim i As Integer

  If (Not completely_new_graph) Then
    Call errors_show("CreateGraph", "Must call this routine after a set handle_ctlPicture = {something} call only.")
    Exit Sub
  End If
  completely_new_graph = False
  
  'SET MANY VARIABLES TO THEIR INITIAL SETTINGS.
  title_main.title = ""
  title_x.title = ""
  title_y.title = ""
  title_x.number_format = ""      'A BLANK STRING INDICATES NON-FORMATTED.
  title_y.number_format = ""      'A BLANK STRING INDICATES NON-FORMATTED.
  number_of_series = 0
  xmin_specified = False
  xmax_specified = False
  ymin_specified = False
  ymax_specified = False
  xmin_user = 0#
  xmax_user = 0#
  ymin_user = 0#
  ymax_user = 0#
  xmin_current = 0#
  xmax_current = 0#
  ymin_current = 0#
  ymax_current = 0#

  xticks_specified = False
  yticks_specified = False
  xticks_user = 0#
  yticks_user = 0#
  xticks_current = 0#
  yticks_current = 0#
  
  title_main.font = "arial"
  title_main.font_size = 18
  title_main.font_bold = False
  
  title_x.font = "arial"
  title_x.font_size = 18
  title_x.font_bold = False
  
  title_y.font = "arial"
  title_y.font_size = 18
  title_y.font_bold = False
  
  xticks_numbers.font = "arial"
  xticks_numbers.font_size = 8
  xticks_numbers.font_bold = False
  xticks_numbers.visible = True
  
  yticks_numbers.font = "arial"
  yticks_numbers.font_size = 8
  yticks_numbers.font_bold = False
  yticks_numbers.visible = True
  
  Call change_title_main(new_title_main)
  Call change_title_x(new_title_x)
  Call change_title_y(new_title_y)
  
  'MAKE THE PICTUREBOX VISIBLE.
  ctlPicture.visible = True
  
  'ACTUALLY DRAW THE GRAPH.
  Call Refresh_Graph
  
End Sub


Private Sub draw_xtick( _
    x As Double, _
    Do_Draw As Boolean, _
    ThisRect As Rect_Type, _
    This_X_Text_Width As Double)
Dim x1 As Double
Dim y1 As Double
Dim x2 As Double
Dim y2 As Double
Dim outline As String
Dim num_x As Double
Dim num_y As Double
  'DRAW THE TICK LINE.
  x1 = translate_x_value_to_screen(x)
  x2 = x1
  y1 = translate_y_value_to_screen(0#)
  y2 = y1 + XAXIS_TICK_HEIGHT
  If (Do_Draw) Then
    ctlPicture.Line (x1, y1)-(x2, y2), QBColor(0)
  End If
  'DRAW THE TICK TEXT.
  ctlPicture.font = xticks_numbers.font
  ctlPicture.font.Size = xticks_numbers.font_size
  ctlPicture.font.Bold = xticks_numbers.font_bold
  outline = format_xtick(x)
  num_x = x1 - ctlPicture.TextWidth(outline) / 2#
  num_y = y2 + 0.5 * XAXIS_TICK_HEIGHT
  ctlPicture.CurrentX = num_x
  ctlPicture.CurrentY = num_y
  If (Do_Draw) Then
    ctlPicture.Print outline
  End If
  ThisRect.x1 = num_x
  ThisRect.y1 = num_y
  ThisRect.x2 = ThisRect.x1 + ctlPicture.TextWidth(outline)
  ThisRect.y2 = ThisRect.y1 + ctlPicture.TextHeight(outline)
  This_X_Text_Width = ctlPicture.TextWidth(outline)
End Sub
Private Sub draw_axis_x( _
    Any_Collision As Boolean, _
    Do_Draw As Boolean, _
    Widest_X_Text As Double)
Dim i As Integer
Dim j As Integer
Dim x As Double
Dim ThisRect As Rect_Type
Dim This_X_Text_Width As Double
Dim AllTextRects() As Rect_Type
Dim Size_AllTextRects As Integer
  Size_AllTextRects = 0
  Widest_X_Text = 0#
  If (xticks_numbers.visible = True) Then
    Call draw_xtick(0#, Do_Draw, ThisRect, This_X_Text_Width)
    If (This_X_Text_Width > Widest_X_Text) Then Widest_X_Text = This_X_Text_Width
    Size_AllTextRects = Size_AllTextRects + 1
    ReDim Preserve AllTextRects(1 To Size_AllTextRects)
    AllTextRects(Size_AllTextRects) = ThisRect
    i = 0
    Do While (1 = 1)
      i = i + 1
      x = xticks_current * CDbl(i)
      If (x > xmax_current) Then Exit Do
      Call draw_xtick(x, Do_Draw, ThisRect, This_X_Text_Width)
      If (This_X_Text_Width > Widest_X_Text) Then Widest_X_Text = This_X_Text_Width
      Size_AllTextRects = Size_AllTextRects + 1
      ReDim Preserve AllTextRects(1 To Size_AllTextRects)
      AllTextRects(Size_AllTextRects) = ThisRect
    Loop
    i = 0
    Do While (1 = 1)
      i = i + 1
      x = -xticks_current * CDbl(i)
      If (x < xmin_current) Then Exit Do
      Call draw_xtick(x, Do_Draw, ThisRect, This_X_Text_Width)
      If (This_X_Text_Width > Widest_X_Text) Then Widest_X_Text = This_X_Text_Width
      Size_AllTextRects = Size_AllTextRects + 1
      ReDim Preserve AllTextRects(1 To Size_AllTextRects)
      AllTextRects(Size_AllTextRects) = ThisRect
    Loop
    'TEST FOR _HORIZONTAL_ COLLISION.
    Any_Collision = False
    For i = 1 To Size_AllTextRects
      For j = i To Size_AllTextRects
        If (AllTextRects(i).x1 >= AllTextRects(j).x1) And (AllTextRects(i).x1 <= AllTextRects(j).x2) Then
          Any_Collision = True
        End If
        If (AllTextRects(i).x2 >= AllTextRects(j).x1) And (AllTextRects(i).x2 <= AllTextRects(j).x2) Then
          Any_Collision = True
        End If
      Next j
    Next i
  Else
    Any_Collision = False
  End If
End Sub


Private Sub draw_ytick( _
    y As Double, _
    Do_Draw As Boolean, _
    ThisRect As Rect_Type, _
    This_Y_Text_Width As Double)
Dim x1 As Double
Dim y1 As Double
Dim x2 As Double
Dim y2 As Double
Dim outline As String
Dim num_x As Double
Dim num_y As Double
  'DRAW THE TICK LINE.
  x1 = translate_x_value_to_screen(0#)
  x2 = x1 - YAXIS_TICK_WIDTH
  y1 = translate_y_value_to_screen(y)
  y2 = y1
  If (Do_Draw) Then
    ctlPicture.Line (x1, y1)-(x2, y2), QBColor(0)
  End If
  'DRAW THE TICK TEXT.
  ctlPicture.font = yticks_numbers.font
  ctlPicture.font.Size = yticks_numbers.font_size
  ctlPicture.font.Bold = yticks_numbers.font_bold
  outline = format_ytick(y)
  num_x = x2 - ctlPicture.TextWidth(outline) - 0.5 * YAXIS_TICK_WIDTH
  num_y = y1 - ctlPicture.TextHeight(outline) / 2#
  ctlPicture.CurrentX = num_x
  ctlPicture.CurrentY = num_y
  If (Do_Draw) Then
    ctlPicture.Print outline
  End If
  ThisRect.x1 = num_x
  ThisRect.y1 = num_y
  ThisRect.x2 = ThisRect.x1 + ctlPicture.TextWidth(outline)
  ThisRect.y2 = ThisRect.y1 + ctlPicture.TextHeight(outline)
  This_Y_Text_Width = ctlPicture.TextWidth(outline)
End Sub
Private Sub draw_axis_y( _
    Any_Collision As Boolean, _
    Do_Draw As Boolean, _
    Widest_Y_Text As Double)
Dim i As Integer
Dim y As Double
Dim This_Y_Text_Width As Double
Dim ThisRect As Rect_Type
  'NOTE: THE Any_Collision VARIABLE IS NOT CALCULATED !!
  Widest_Y_Text = 0#
  If (yticks_numbers.visible = True) Then
    Call draw_ytick(0#, Do_Draw, ThisRect, This_Y_Text_Width)   'IGNORE ThisRect.
    If (This_Y_Text_Width > Widest_Y_Text) Then Widest_Y_Text = This_Y_Text_Width
    i = 0
    Do While (1 = 1)
      i = i + 1
      y = yticks_current * CDbl(i)
      If (y > ymax_current) Then Exit Do
      'Call draw_ytick(y, Do_Draw, ThisRect, This_Y_Text_Width)   'IGNORE ThisRect.
      'If (This_Y_Text_Width > Widest_Y_Text) Then Widest_Y_Text = This_Y_Text_Width
    Loop
    'i = 0
    Do While (1 = 1)
      'i = i + 1
      i = i - 1
      y = yticks_current * CDbl(i)
      If (y < ymin_current) Then Exit Do
      Call draw_ytick(y, Do_Draw, ThisRect, This_Y_Text_Width)   'IGNORE ThisRect.
      If (This_Y_Text_Width > Widest_Y_Text) Then Widest_Y_Text = This_Y_Text_Width
    Loop
  End If
End Sub

Private Sub draw_axes()
'Dim i As Integer
Dim x As Double
Dim y As Double
Dim Any_Collision As Boolean
Dim Widest_X_Text As Double
Dim Widest_Y_Text As Double

  If (number_of_series = 0) Then
    Exit Sub
  End If
  
  'DRAW X AND Y AXIS LINES
  x = translate_x_value_to_screen(0#)
  y = translate_y_value_to_screen(0#)
  ctlPicture.Line (x, yguide1)-(x, yguide2), QBColor(0)
  ctlPicture.Line (xguide2, y)-(xguide3, y), QBColor(0)

  'DRAW X AXIS TICKS.
  Call draw_axis_x(Any_Collision, True, Widest_X_Text)
      'IGNORE Any_Collision and Widest_X_Text.
  
  'DRAW Y AXIS TICKS.
  Call draw_axis_y(Any_Collision, True, Widest_Y_Text)
      'IGNORE Any_Collision and Widest_Y_Text.

End Sub

Private Sub draw_graph()
Dim i As Integer
Dim j As Integer
Dim xx As Double
Dim yy As Double
Dim last_xx As Double
Dim last_yy As Double
Dim x As Double
Dim y As Double

  For i = 1 To number_of_series
    last_xx = 1E+200
    last_yy = 1E+200
    For j = 1 To data_series(i).num_points
      x = data_series(i).data_x(j)
      y = data_series(i).data_y(j)
      xx = translate_x_value_to_screen(x)
      yy = translate_y_value_to_screen(y)
      'xx = (x - xmin_current) / (xmax_current - xmin_current) * (xguide3 - xguide2) + xguide2
      'yy = (1# - (y - ymin_current) / (ymax_current - ymin_current)) * (yguide2 - yguide1) + yguide1
      If (data_series(i).symbol_style = 0) Then
        If (last_xx = 1E+200) Or (last_yy = 1E+200) Then
          ctlPicture.PSet (xx, yy), data_series(i).color
        Else
          ctlPicture.Line (last_xx, last_yy)-(xx, yy), data_series(i).color
        End If
      Else
        Call draw_graph_symbol(xx, yy, data_series(i).symbol_style, data_series(i).color)
      End If
      last_xx = xx
      last_yy = yy
    Next j
  Next i
  
'    data_series(this_series_index).data_x(i) = new_data_x(i)
'    data_series(this_series_index).data_y(i) = new_data_y(i)
'  Next i
'  AddSeriesData = this_series_index

End Sub

Private Sub draw_graph_symbol(xx As Double, yy As Double, symbol_style As Integer, color As Long)
Dim x_symbol_size As Double
Dim y_symbol_size As Double

  'symbol_size_ratio
  x_symbol_size = (xguide3 - xguide2) / 200#
  'y_symbol_size = (yguide2 - xguide1) / 100#
  y_symbol_size = x_symbol_size
  
  Select Case symbol_style
    Case 1:     'BOX
      ctlPicture.Line (xx - x_symbol_size, yy - y_symbol_size)-(xx + x_symbol_size, yy - y_symbol_size), color
      ctlPicture.Line (xx + x_symbol_size, yy - y_symbol_size)-(xx + x_symbol_size, yy + y_symbol_size), color
      ctlPicture.Line (xx + x_symbol_size, yy + y_symbol_size)-(xx - x_symbol_size, yy + y_symbol_size), color
      ctlPicture.Line (xx - x_symbol_size, yy + y_symbol_size)-(xx - x_symbol_size, yy - y_symbol_size), color
      
  End Select
  
   'ctlPicture.Line (last_xx, last_yy)-(xx, yy), data_series(i).color

End Sub


Private Sub errors_show(subname As String, msg As String)

  MsgBox "An error occurred in the routine '" & subname & "': " & msg, 64, "GraphControl"

End Sub

Private Function format_xtick(x As Double) As String
  If (title_x.number_format = "") Then
    format_xtick = Trim$(Str$(x))
  Else
    format_xtick = Format$(x, title_x.number_format)
  End If
End Function
Private Function format_ytick(y As Double) As String
  If (title_y.number_format = "") Then
    format_ytick = Trim$(Str$(y))
  Else
    format_ytick = Format$(y, title_y.number_format)
  End If
End Function


Property Set handle_ctlPicture(new_ctlPicture As Control)
  Set ctlPicture = new_ctlPicture
  completely_new_graph = True
End Property


Public Sub change_title_main(new_title_main As String)
  title_main.title = new_title_main
End Sub
Public Sub change_title_x(new_title_x As String)
  title_x.title = new_title_x
End Sub
Public Sub change_title_y(new_title_y As String)
  title_y.title = new_title_y
End Sub


Public Sub Change_X_Number_Format(NewFormat As String)
  title_x.number_format = NewFormat
End Sub
Public Sub Change_Y_Number_Format(NewFormat As String)
  title_y.number_format = NewFormat
End Sub


Public Sub Refresh_Graph()

  Call calculate_guides
  
  ctlPicture.Cls
  
  Call draw_axes
  Call draw_graph
  
  If (title_main.visible) Then
    ctlPicture.ForeColor = QBColor(0)
    ctlPicture.CurrentX = title_main.left
    ctlPicture.CurrentY = title_main.top
    ctlPicture.font = title_main.font
    ctlPicture.font.Size = title_main.font_size
    ctlPicture.font.Bold = title_main.font_bold
    ctlPicture.Print title_main.title
  End If
  
  If (title_x.visible) Then
    ctlPicture.ForeColor = QBColor(0)
    ctlPicture.CurrentX = title_x.left
    ctlPicture.CurrentY = title_x.top
    ctlPicture.font = title_x.font
    ctlPicture.font.Size = title_x.font_size
    ctlPicture.font.Bold = title_x.font_bold
    ctlPicture.Print title_x.title
  End If
  
  If (title_y.visible) Then
    ctlPicture.ForeColor = QBColor(0)
    ctlPicture.CurrentX = title_y.left
    ctlPicture.CurrentY = title_y.top
    ctlPicture.font = title_y.font
    ctlPicture.font.Size = title_y.font_size
    ctlPicture.font.Bold = title_y.font_bold
    ctlPicture.Print title_y.title
  End If

End Sub


Private Sub temp_test_output()
Dim i As Integer
  
  Open "c:\test.txt" For Output As #1
  For i = 1 To 1000
    Print #1, data_series(1).data_x(i), data_series(1).data_y(i)
  Next i
  Close #1
  
End Sub


Private Function translate_x_value_to_screen(x As Double) As Double
  translate_x_value_to_screen = (x - xmin_current) / _
      (xmax_current - xmin_current) * _
      (xguide3 - xguide2) + xguide2
End Function
Private Function translate_y_value_to_screen(y As Double) As Double
  translate_y_value_to_screen = (1# - (y - ymin_current) / _
      (ymax_current - ymin_current)) * _
      (yguide2 - yguide1) + yguide1
End Function

