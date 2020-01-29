VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmPlot 
   Caption         =   "Plotted Results"
   ClientHeight    =   4740
   ClientLeft      =   1410
   ClientTop       =   2775
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4740
   ScaleWidth      =   9270
   Begin Threed.SSPanel ssframe_GraphHolder 
      Height          =   3135
      Left            =   3090
      TabIndex        =   0
      Top             =   180
      Width           =   4275
      _Version        =   65536
      _ExtentX        =   7541
      _ExtentY        =   5530
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.PictureBox picGraph 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2805
         Left            =   90
         ScaleHeight     =   2805
         ScaleWidth      =   3405
         TabIndex        =   1
         Top             =   90
         Width           =   3405
      End
   End
   Begin Threed.SSFrame ssframe_series 
      Height          =   1425
      Left            =   30
      TabIndex        =   2
      Top             =   990
      Width           =   2985
      _Version        =   65536
      _ExtentX        =   5265
      _ExtentY        =   2514
      _StockProps     =   14
      Caption         =   "Dataset (Y) Axis Units:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CheckBox chkPlot 
         Caption         =   "Plot Item 1"
         Height          =   495
         Index           =   0
         Left            =   1350
         TabIndex        =   4
         Top             =   330
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ComboBox cboUnits 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   390
         Visible         =   0   'False
         Width           =   1275
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   885
      Left            =   30
      TabIndex        =   5
      Top             =   30
      Width           =   2985
      _Version        =   65536
      _ExtentX        =   5265
      _ExtentY        =   1561
      _StockProps     =   14
      Caption         =   "Time (X) Axis Units:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ComboBox cboTimeUnits 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   390
         Width           =   2000
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Refresh Plot"
         Index           =   100
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   198
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Close"
         Index           =   199
         Shortcut        =   ^C
      End
   End
End
Attribute VB_Name = "frmPlot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim activated_yet As Boolean

Dim Num_Series As Integer
Const SERIESLIST_SPACE_BETWEEN = 405
Const SERIESLIST_CBOUNITS_TOP_FIRST = 390
Const SERIESLIST_CHKPLOT_TOP_FIRST = 315
Const SERIESLIST_BOTTOM_MARGIN = 200
Const SERIESLIST_TAG_H2O2 = "H2O2"
Const SERIESLIST_TAG_NOM = "NOM"
Const SERIESLIST_TAG_PH = "PH"
'NOTE: THE TAG FOR TARGET COMPOUND #1 IS "1", ET CETERA.

Dim HALT_CBOUNITS As Boolean
Dim HALT_CHKPLOT As Boolean
Dim HALT_CBOTIMEUNITS As Boolean

Const CONCUNITS_G_L = 1
Const CONCUNITS_MG_L = 2
Const CONCUNITS_UG_L = 3
Const CONCUNITS_M = 4
Const CONCUNITS_MM = 5
Const CONCUNITS_UM = 6
Const CONCUNITS_C_Co = 7

Const TIMEUNITS_SECONDS = 1
Const TIMEUNITS_MINUTES = 2
Const TIMEUNITS_HOURS = 3
Const TIMEUNITS_DAYS = 4

Dim ThisGraph As Object
Dim READY_TO_PLOT As Boolean



Const frmPlot_declarations_end = True


Sub Redraw_Graph()
ReDim data_x(10) As Double
ReDim data_y(10) As Double
Dim i As Integer
Dim j As Integer
Dim num_chemicals As Integer
Dim num_tanks As Integer
Dim num_rows As Integer
Dim series_link() As Integer
Dim Conc_In_M As Double
Dim Component_MW As Double
Dim Influent_Conc_In_M As Double
Dim Time_In_Minutes As Double
Dim ThisColor As Integer

  'ABORT IF NOT READY TO PLOT.
  If (READY_TO_PLOT = False) Then
    Exit Sub
  End If
  Screen.MousePointer = 11
  'REMOVE ALL EXISTING GRAPH DATA.
  Call ThisGraph.DeleteAllSeries
  'DETERMINE INDEXES OF SERIES TO PLOT.
  num_chemicals = UBound(TankConcs, 1)
  num_tanks = UBound(TankConcs, 2)
  num_rows = UBound(TankConcs, 3)
  ReDim series_link(1 To Num_Series)
  For i = 1 To Num_Series
    series_link(i) = 0
    For j = 1 To num_chemicals
      If (Trim$(UCase$(TankConcLabels(j).Label1)) = _
          Trim$(UCase$(chkPlot(i).Tag))) Then
        series_link(i) = j
        Exit For
      End If
    Next j
  Next i
  'PLOT THE DATA.
  ThisColor = 8
  For i = 1 To Num_Series
    If (chkPlot(i).Value = 0) Then
      'DO NOT PLOT THIS SERIES.
      chkPlot(i).ForeColor = QBColor(0)
    Else
      'PLOT THIS SERIES.
      'PLOT THIS SERIES STEP #1: TRANSFER DATA TO LOCAL X-Y ARRAYS.
      ReDim data_x(1 To num_rows)
      ReDim data_y(1 To num_rows)
      For j = 1 To num_rows
        'CALCULATE VALUE OF X IN APPROPRIATE UNITS.
        Time_In_Minutes = Tank_Times(j)
        Select Case cboTimeUnits.ItemData(cboTimeUnits.ListIndex)
          Case TIMEUNITS_SECONDS: data_x(j) = Time_In_Minutes * 60#
          Case TIMEUNITS_MINUTES: data_x(j) = Time_In_Minutes * 1#
          Case TIMEUNITS_HOURS: data_x(j) = Time_In_Minutes / 60#
          Case TIMEUNITS_DAYS: data_x(j) = Time_In_Minutes / 60# / 24#
        End Select
        'CALCULATE VALUE OF Y IN APPROPRIATE UNITS.
        Conc_In_M = TankConcs(series_link(i), num_tanks, j)
        If (i <= NowProj.TargetCompounds_Count) Then
          Component_MW = NowProj.TargetCompounds(i).mw
          If (Trim$(UCase$(chkPlot(i).Tag)) = Trim$(UCase$(SERIESLIST_TAG_NOM))) Then
            'STORAGE UNITS FOR NOM ARE IN mg/L.
            Conc_In_M = Conc_In_M * 0.001 / Component_MW
            Influent_Conc_In_M = NowProj.TargetCompounds(i).concini * 0.001 / Component_MW
          Else
            'STORAGE UNITS FOR ALL OTHER COMPONENTS IN gmol/L (M).
            Conc_In_M = Conc_In_M * 1#
            Influent_Conc_In_M = NowProj.TargetCompounds(i).concini
          End If
        Else
          Select Case Trim$(UCase$(chkPlot(i).Tag))
            Case Trim$(UCase$(SERIESLIST_TAG_H2O2)):
              Component_MW = 34#
              Influent_Conc_In_M = NowProj.inf_h2o2
          End Select
        End If
        If (cboUnits(i).visible = False) Then
          data_y(j) = Conc_In_M
        Else
        Select Case cboUnits(i).ItemData(cboUnits(i).ListIndex)
          Case CONCUNITS_G_L: data_y(j) = Conc_In_M * Component_MW
          Case CONCUNITS_MG_L: data_y(j) = Conc_In_M * 1000# * Component_MW
          Case CONCUNITS_UG_L: data_y(j) = Conc_In_M * 1000# * 1000# * Component_MW
          Case CONCUNITS_M: data_y(j) = Conc_In_M * 1#
          Case CONCUNITS_MM: data_y(j) = Conc_In_M * 1000#
          Case CONCUNITS_UM: data_y(j) = Conc_In_M * 1000# * 1000#
          Case CONCUNITS_C_Co:
            If (Influent_Conc_In_M = 0#) Then
              data_y(j) = 0#
            Else
              data_y(j) = Conc_In_M / Influent_Conc_In_M
            End If
        End Select
        End If
        'data_y(j) = TankConcs(series_link(i), num_tanks, j)
      Next j
      'PLOT THIS SERIES STEP #2: TRANSFER DATA TO GRAPH CONTROL.
      ThisColor = ThisColor + 1
      If (ThisColor > 14) Then ThisColor = 1
      Call ThisGraph.AddSeriesData( _
          "Series Whatever", CLng(num_rows), data_x, data_y, _
          0, 1#, QBColor(ThisColor))
      'COLOR THIS CHECKBOX.
      chkPlot(i).ForeColor = QBColor(ThisColor)
    End If
  Next i
  
  'MISCELLANEOUS FORMATTING STUFF.
  'Call graph.Change_X_Number_Format("0.00e+00")
  'Call graph.Change_Y_Number_Format("0.00e+00")
  Call ThisGraph.Refresh_Graph
  Screen.MousePointer = 0
  
End Sub


Sub populate_cboTimeUnits()
Dim ctl As Control
Set ctl = cboTimeUnits
  ctl.Clear
  ctl.AddItem "seconds": ctl.ItemData(ctl.NewIndex) = TIMEUNITS_SECONDS
  ctl.AddItem "minutes": ctl.ItemData(ctl.NewIndex) = TIMEUNITS_MINUTES
  ctl.AddItem "hours": ctl.ItemData(ctl.NewIndex) = TIMEUNITS_HOURS
  ctl.AddItem "days": ctl.ItemData(ctl.NewIndex) = TIMEUNITS_DAYS
  ctl.ListIndex = 1       'minutes
End Sub
Sub populate_Concentration_Units(ctl As Control)
  ctl.Clear
  ctl.AddItem "g/L": ctl.ItemData(ctl.NewIndex) = CONCUNITS_G_L
  ctl.AddItem "mg/L": ctl.ItemData(ctl.NewIndex) = CONCUNITS_MG_L
  ctl.AddItem "µg/L": ctl.ItemData(ctl.NewIndex) = CONCUNITS_UG_L
  ctl.AddItem "M": ctl.ItemData(ctl.NewIndex) = CONCUNITS_M
  ctl.AddItem "mM": ctl.ItemData(ctl.NewIndex) = CONCUNITS_MM
  ctl.AddItem "µM": ctl.ItemData(ctl.NewIndex) = CONCUNITS_UM
  ctl.AddItem "C/Co": ctl.ItemData(ctl.NewIndex) = CONCUNITS_C_Co
  ctl.ListIndex = 2       'µg/L
End Sub
Sub populate_cboUnits_and_chkPlot()
Dim i As Integer
Dim TotalHeight As Long
  HALT_CBOUNITS = True
  HALT_CHKPLOT = True
  HALT_CBOTIMEUNITS = True
  Num_Series = 2 + NowProj.TargetCompounds_Count
  For i = 1 To Num_Series
    Load cboUnits(i)
    Load chkPlot(i)
    cboUnits(i).top = SERIESLIST_CBOUNITS_TOP_FIRST + (i - 1) * SERIESLIST_SPACE_BETWEEN
    chkPlot(i).top = SERIESLIST_CHKPLOT_TOP_FIRST + (i - 1) * SERIESLIST_SPACE_BETWEEN
    cboUnits(i).visible = True
    chkPlot(i).visible = True
    chkPlot(i).Value = False
  Next i
  'SET LABELS.
  For i = 1 To NowProj.TargetCompounds_Count
    chkPlot(i).Caption = left$(NowProj.TargetCompounds(i).comname, 12)
    'chkPlot(i).Tag = Trim$(Str$(i))
    chkPlot(i).Tag = left$(Trim$(UCase$(NowProj.TargetCompounds(i).comname)), 12)
    If (i = 2) Then chkPlot(i).Value = 1  'True
    Call populate_Concentration_Units(cboUnits(i))
  Next i
  i = NowProj.TargetCompounds_Count + 1
  chkPlot(i).Caption = "H2O2"
  chkPlot(i).Tag = SERIESLIST_TAG_H2O2
  Call populate_Concentration_Units(cboUnits(i))
  i = i + 1
  chkPlot(i).Caption = "pH"
  chkPlot(i).Tag = SERIESLIST_TAG_PH
  cboUnits(i).visible = False
  ''''i = i + 1
  ''''chkPlot(i).Caption = "NOM"
  ''''chkPlot(i).Tag = SERIESLIST_TAG_NOM
  'SIZE THE FRAME.
  TotalHeight = cboUnits(Num_Series).top + cboUnits(Num_Series).height
  TotalHeight = TotalHeight + SERIESLIST_BOTTOM_MARGIN
  ssframe_series.height = TotalHeight
  HALT_CBOUNITS = False
  HALT_CHKPLOT = False
  HALT_CBOTIMEUNITS = False
End Sub


Private Sub cboTimeUnits_Click()
  If (HALT_CBOTIMEUNITS = True) Then
    Exit Sub
  End If
  Call Redraw_Graph
End Sub
Private Sub cboUnits_Click(Index As Integer)
  If (HALT_CBOUNITS = True) Then
    Exit Sub
  End If
  Call Redraw_Graph
End Sub
Private Sub chkPlot_Click(Index As Integer)
  If (HALT_CHKPLOT = True) Then
    Exit Sub
  End If
  Call Redraw_Graph
End Sub


Private Sub Form_Activate()
Dim Any_Error As Boolean
  If (Not activated_yet) Then
    activated_yet = True
    'PREPARE OUTPUT CONCENTRATIONS.
    Call PrintPrepare_OutputConcs(NowProj, Any_Error)
    If (Any_Error) Then
      Unload Me
      Exit Sub
    End If
    'SET UP UNIT SCROLLBOXES AND PLOT CHECKBOXES.
    Call populate_cboUnits_and_chkPlot
    'SET UP TIME UNITS.
    Call populate_cboTimeUnits
    'CREATE GRAPH OBJECT.
    Set ThisGraph = New GraphControl
    Set ThisGraph.handle_ctlPicture = picGraph
    'Call ThisGraph.CreateGraph("Title", "X", "Y")
    Call ThisGraph.CreateGraph("", "", "")
    'DRAW GRAPH FOR FIRST TIME.
    Call CenterOnForm(Me, frmMain)
    Me.WindowState = 2  'maximized
    READY_TO_PLOT = True
    Call Redraw_Graph
  End If
End Sub
Private Sub Form_Load()
' Call FortranLink_SetFilenames
'  If (FileExists(FortranLink_fn_MainOutput)) Then
'    'DO NOTHING--CODE IS BELOW.
'  Else
'    Call Show_Error("There are no results to view.  Run the simulation first.")
'    Exit Sub
'  End If
  activated_yet = False
  READY_TO_PLOT = False
  Me.width = 9000
  Me.height = 7000
  Me.top = -30000
  Me.left = -30000
End Sub
Private Sub Form_Resize()
Dim XXX As Long
Dim USE_MARGIN As Long
  USE_MARGIN = 100
  'If (frmPrint_loading_now) Then Exit Sub
  If (Me.WindowState = 1) Then
    'CANNOT RESIZE WHEN MINIMIZED; EXIT OUTTA HERE.
    Exit Sub
  End If
  If (Me.width < 6000) Then Me.width = 6000
  If (Me.height < 3000) Then Me.height = 3000
  'RESIZE ssframe_GraphHolder.
  XXX = Me.width - (Me.width - Me.ScaleWidth) - _
      ssframe_GraphHolder.left - USE_MARGIN
  If (XXX < 1000) Then
    XXX = 1000
  End If
  ssframe_GraphHolder.width = XXX
  
'  XXX = Me.height - (Me.height - Me.ScaleHeight) - _
'      StatusBar1.height - ssframe_GraphHolder.top - USE_MARGIN
  XXX = Me.height - (Me.height - Me.ScaleHeight) - _
      ssframe_GraphHolder.top - USE_MARGIN

  If (XXX < 1000) Then
    XXX = 1000
  End If
  ssframe_GraphHolder.height = XXX
  'RESIZE picGraph.
  XXX = ssframe_GraphHolder.width - picGraph.left * 2
  If (XXX < 1000) Then XXX = 1000
  picGraph.width = XXX
  XXX = ssframe_GraphHolder.height - picGraph.top * 2
  If (XXX < 1000) Then XXX = 1000
  picGraph.height = XXX
  Call Redraw_Graph
End Sub


Private Sub mnuFileItem_Click(Index As Integer)
  Select Case Index
    Case 100:       'refresh plot
      Call Redraw_Graph
    Case 199:       'close
      Unload Me
      Exit Sub
  End Select
End Sub

