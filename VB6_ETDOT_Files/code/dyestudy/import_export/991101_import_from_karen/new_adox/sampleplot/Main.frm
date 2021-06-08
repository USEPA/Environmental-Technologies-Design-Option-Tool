VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmPlot 
   Caption         =   "Plot Example"
   ClientHeight    =   6735
   ClientLeft      =   1530
   ClientTop       =   1320
   ClientWidth     =   9150
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6735
   ScaleWidth      =   9150
   Begin VB.CommandButton Command1 
      Caption         =   "Plot !"
      Height          =   405
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2085
   End
   Begin Threed.SSPanel ssframe_GraphHolder 
      Height          =   5985
      Left            =   90
      TabIndex        =   0
      Top             =   600
      Width           =   8835
      _Version        =   65536
      _ExtentX        =   15584
      _ExtentY        =   10557
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
         Height          =   5775
         Left            =   90
         ScaleHeight     =   5775
         ScaleWidth      =   8595
         TabIndex        =   1
         Top             =   90
         Width           =   8595
      End
   End
End
Attribute VB_Name = "frmPlot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ThisGraph As Object






Private Sub Command1_Click()
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
  num_rows = 4
  ReDim data_x(1 To num_rows)
  ReDim data_y(1 To num_rows)
  data_x(1) = 1#: data_y(1) = 1#
  data_x(2) = 2#: data_y(2) = 2#
  data_x(3) = 3#: data_y(3) = 3#
  data_x(4) = 4#: data_y(4) = 4#
  Call ThisGraph.AddSeriesData( _
      "Series Whatever", CLng(num_rows), data_x, data_y, _
      0, 1#, QBColor(9))
  '
  ' ADD THE SECOND SERIES.
  '
  num_rows = 4
  ReDim data_x(1 To num_rows)
  ReDim data_y(1 To num_rows)
  data_x(1) = 1#: data_y(1) = 1# + 0.1
  data_x(2) = 2#: data_y(2) = 2# - 0.2
  data_x(3) = 3#: data_y(3) = 3# + 0.3
  data_x(4) = 4#: data_y(4) = 4# - 0.1
  Call ThisGraph.AddSeriesData( _
      "Series Whatever", CLng(num_rows), data_x, data_y, _
      1, 1#, QBColor(12))
  '
  ' ACTUALLY MAKE THE DATA BE DISPLAYED.
  '
  Call ThisGraph.Refresh_Graph
End Sub


Private Sub Form_Load()
  'CREATE GRAPH OBJECT.
  Set ThisGraph = New GraphControl
  Set ThisGraph.handle_ctlPicture = picGraph
  'Call ThisGraph.CreateGraph("Title", "X", "Y")
  Call ThisGraph.CreateGraph("", "", "")
  'DRAW GRAPH FOR FIRST TIME.
  'Call CenterOnForm(Me, frmPlot)
  'Me.WindowState = 2  'maximized
  'READY_TO_PLOT = True
  'Call Redraw_Graph
End Sub


