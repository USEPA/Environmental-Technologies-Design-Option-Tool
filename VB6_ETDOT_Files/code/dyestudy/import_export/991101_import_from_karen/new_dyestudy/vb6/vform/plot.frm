VERSION 5.00
Begin VB.Form frmPlot 
   Caption         =   "Plot Example"
   ClientHeight    =   6735
   ClientLeft      =   1860
   ClientTop       =   2670
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
   Begin VB.PictureBox ssframe_GraphHolder 
      BackColor       =   &H00C0C0C0&
      Height          =   5985
      Left            =   90
      ScaleHeight     =   5925
      ScaleWidth      =   8775
      TabIndex        =   0
      Top             =   600
      Width           =   8835
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
'
'(2.) Note:   The new PEC.EXE program was changed so (A) the English is a bit
'improved, and (B) the OUTPT2.TXT file is generated. The new file OUTPT2.TXT
'consists of a row count, and a set of predicted values. You can continue to
'display/load/save the old file OUTPT.TXT as always.
'(3.) Each time the PEC.EXE program is called, have the VB read in the row
'count, and the data. The arrays should be named something like
'Predicted_Theta() and Predicted_E(). These arrays must be loaded and saved
'in the .DYE file. (I recommend that you store the arrays of values instead
'of storing the entire OUTPT2.TXT file in a memo field, but you may find a
'better way.)
'(4.) Add a variable Predicted_Available (boolean) which defaults to False.
'It must be loaded and saved in the .DYE file.
'(5.) Change the main window so it is sizable and maximizable. Move things
'around so a sizable plot can be placed on the window.
'(6.) Using the code from the attached 990930_simple*.zip file, set up a
'simple plot on the DyeStudy main window. To start, just have it plot
'something simple (like the data in the simple example I attached).
'(7.) Make the plot "resizable" so that when the main window is resized, the
'plot is automatically resized. (Refer to the code in NewAdox -- frmPlot for
'a template.)
'(8.) In DyeStudy -- Refresh -- refresh_frmMain(), add code so that when
'Predicted_Available is false, the plot is invisible, and a message is
'displayed indicating something like "Plot is unavailable. Please enter data
'and recalculate."
'(9.) Also in refresh_frmMain(), if Predicted_Available is true, have the
'plot be visible and refresh it with the proper data.
'
'Here 's another quick change that Luke recommended:
'
'(1.) In Project_SetDefaults(), set .alk to 200.
'(2.) In the batch1.adx, batch2.adx, and tis2.adx example files, reset the
'value of .alk to 200 and re-save. But make sure that the "TIC input as" is
'specified as "TIC".
'(3.) The reason for this is because of the following conversion factor that
'applies for a pH range of approximately 6 to 9
'          (alkalinity, g/L as CaCO3) = (TIC, gmol/L) * 50
'          (TIC, gmol/L) = (alkalinity, g/L as CaCO3) / 50

  
  'CREATE GRAPH OBJECT.
  Set ThisGraph = New GraphControl
  Set ThisGraph.handle_ctlPicture = picGraph
  'Call ThisGraph.CreateGraph("Title", "X", "Y")
  Call ThisGraph.CreateGraph("", "", "")
  'DRAW GRAPH FOR FIRST TIME.
  'Call CenterOnForm(Me, frmMain)
  'Me.WindowState = 2  'maximized
  'READY_TO_PLOT = True
  'Call Redraw_Graph
End Sub



