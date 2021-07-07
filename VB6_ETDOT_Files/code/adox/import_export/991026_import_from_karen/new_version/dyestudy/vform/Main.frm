VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Dye Study"
   ClientHeight    =   5655
   ClientLeft      =   2415
   ClientTop       =   3090
   ClientWidth     =   7875
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "Main.frx":0000
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5655
   ScaleWidth      =   7875
   Begin VB.Frame fraMain 
      Height          =   4755
      Left            =   210
      TabIndex        =   5
      Top             =   60
      Width           =   6525
      Begin VB.CommandButton cmdDisplayResults 
         Caption         =   "&Display Results"
         Height          =   555
         Left            =   210
         TabIndex        =   9
         Top             =   840
         Width           =   2325
      End
      Begin VB.TextBox txtData 
         Enabled         =   0   'False
         Height          =   345
         Index           =   0
         Left            =   3990
         TabIndex        =   7
         Text            =   "txtData(0)"
         Top             =   300
         Width           =   1635
      End
      Begin VB.CommandButton cmdEditDyeStudyData 
         Caption         =   "&Edit Dye Study Data"
         Height          =   555
         Left            =   210
         TabIndex        =   6
         Top             =   240
         Width           =   2325
      End
      Begin Threed.SSPanel ssframe_GraphHolder 
         Height          =   1755
         Left            =   2670
         TabIndex        =   11
         Top             =   780
         Width           =   3615
         _Version        =   65536
         _ExtentX        =   6376
         _ExtentY        =   3096
         _StockProps     =   15
         ForeColor       =   -2147483630
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
            Height          =   1275
            Left            =   180
            ScaleHeight     =   1275
            ScaleWidth      =   3075
            TabIndex        =   12
            Top             =   240
            Width           =   3075
         End
      End
      Begin VB.Label lblPlotAxes 
         Caption         =   "On the plot, the x-axis is the dimensionless time (theta), and the y-axis is the dimensionless concentration (E)"
         Height          =   1455
         Left            =   330
         TabIndex        =   13
         Top             =   1710
         Width           =   2025
      End
      Begin VB.Label lblDesc 
         Caption         =   "Last Calculated:"
         Height          =   465
         Index           =   0
         Left            =   2730
         TabIndex        =   8
         Top             =   270
         Width           =   1065
      End
   End
   Begin VB.Frame Invisible 
      Caption         =   "Frame2"
      Height          =   1605
      Left            =   7020
      TabIndex        =   2
      Top             =   1740
      Visible         =   0   'False
      Width           =   2715
      Begin VB.PictureBox Picture1_nolongerused 
         Height          =   9660
         Left            =   2610
         ScaleHeight     =   9600
         ScaleWidth      =   12075
         TabIndex        =   3
         Top             =   600
         Width           =   12135
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   210
         Top             =   390
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin Threed.SSPanel panDirty_to_be_deleted 
         Height          =   285
         Left            =   510
         TabIndex        =   4
         Top             =   300
         Width           =   1845
         _Version        =   65536
         _ExtentX        =   3254
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "Data Unchanged"
         ForeColor       =   -2147483630
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
      End
   End
   Begin Threed.SSPanel sspBottom 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   5250
      Width           =   7875
      _Version        =   65536
      _ExtentX        =   13891
      _ExtentY        =   714
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSPanel sspanel_Dirty 
         Height          =   285
         Left            =   150
         TabIndex        =   1
         Top             =   60
         Width           =   3525
         _Version        =   65536
         _ExtentX        =   6218
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "sspanel_Dirty"
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin Threed.SSPanel sspanel_Status 
         Height          =   285
         Left            =   5550
         TabIndex        =   10
         Top             =   60
         Width           =   3675
         _Version        =   65536
         _ExtentX        =   6482
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "sspanel_Status"
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileItem 
         Caption         =   "&New"
         Index           =   0
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Open ..."
         Index           =   1
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Save"
         Index           =   2
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "Save &As ..."
         Index           =   3
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "P&rinter Setup"
         Index           =   6
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Print"
         Index           =   7
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   190
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&1 Old File #1"
         Index           =   191
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&2 Old File #2"
         Index           =   192
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&3 Old File #3"
         Index           =   193
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&4 Old File #4"
         Index           =   194
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   199
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "E&xit"
         Index           =   200
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpItem 
         Caption         =   "View Version History ..."
         Index           =   80
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "View Disclaimer ..."
         Index           =   85
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "-"
         Index           =   98
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "&About ..."
         Index           =   99
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim USER_HIT_CANCEL As Integer
'''Dim nowproj As Project_Type
Dim frmPrint_loading_now As Boolean
Dim READY_TO_PLOT As Boolean




Const frmMain_declarations_end = True


Sub Avoid_Weird_Focus_Problem()
  Call unitsys_control_MostRecent_Force_lostfocus
  'frmMain.SetFocus
  '
  ' NOTE: IT IS VERY IMPORTANT TO SET FOCUS HERE
  ' TO SOME NON-UNITTEXTBOX CONTROL, I.E. DON'T
  ' SET IT TO txtData(0...3), BUT cboUnits(0)
  ' IS OKAY.
'  cboUnits(0).SetFocus
  'Text1.SetFocus
End Sub





Sub Local_DirtyStatus_Set(DirtyFlag As Boolean, NewSetting As Boolean)
  Call Global_DirtyStatus_Set(Me, DirtyFlag, NewSetting)
End Sub
Sub Local_GenericStatus_Set(NewString As String)
  Call Global_GenericStatus_Set(Me, NewString)
End Sub



Private Sub cmdDyeStudy_Click(Index As Integer)

End Sub


Private Sub DyeStudy_Click()
Dim RetVal As Integer

  RetVal = frmDyeStudy.frmDyeStudy_DoEdit()
  If (RetVal) Then
    'USER HIT OK; ASSUME THEY MODIFIED SOMETHING.
    
    'REFRESH MAIN WINDOW, ALTHOUGH PROBABLY
    'NOTHING ON THE MAIN WINDOW NEEDS REFRESHING.
    Call refresh_frmMain
    
    'THROW DIRTY FLAG.
   
    If nowproj.dirty Then
       'THROW DIRTY FLAG.
       Call Local_DirtyStatus_Set( _
           Project_Is_Dirty, True)
    End If
    Call DirtyFlag_Refresh(nowproj)
  Else
    'RESTORE DIRTY FLAG DISPLAY IF NEEDED.
    Call DirtyFlag_Refresh(nowproj)
  End If
End Sub


Private Sub cmdDisplayPlot_Click()

End Sub

Private Sub cmdDisplayResults_Click()
Dim fn_this As String
  'see if data changed and not calculated
  If Not IsCalculated Then
    Call Show_Message("Data was changed, please calculate first.", _
    vbExclamation, App.title)
  Else
    'look for output.txt and if not there,display message
    fn_this = MAIN_APP_PATH & "\exes\outpt.txt"
    If (FileExists(fn_this) = False) Then
      Call Show_Message("No output file was found, please calculate.", _
      vbExclamation, App.title)
    Else
  '    Call Launch_Notepad(App.Path & "\exes\output.txt")
      Call Launch_Notepad(App.Path & "\exes\outpt.txt")
    End If
  End If
End Sub

Private Sub cmdEditDyeStudyData_Click()
Dim RetVal As Integer

  nowproj.dirty = False
  RetVal = frmDyeStudy.frmDyeStudy_DoEdit()
  If (RetVal) Then
    
    'REFRESH MAIN WINDOW, ALTHOUGH PROBABLY
    'NOTHING ON THE MAIN WINDOW NEEDS REFRESHING.
    Call refresh_frmMain
    nowproj.dirty = True
    Call Local_DirtyStatus_Set( _
         Project_Is_Dirty, True)
  
    Call DirtyFlag_Refresh(nowproj)
  Else
    'RESTORE DIRTY FLAG DISPLAY IF NEEDED.
    nowproj.dirty = False
    Call DirtyFlag_Refresh(nowproj)
  End If
  AppActivate App.title
  
End Sub


Private Sub cmdPlotResults_Click()


  Call refresh_frmMain
  
End Sub


Private Sub Form_Activate()
  If (frmPrint_loading_now) Then
    frmPrint_loading_now = False
    Call Form_Resize
  End If
End Sub


Private Sub Form_Load()
  '
  Call Local_DirtyStatus_Set(Project_Is_Dirty, False)
  Call Local_GenericStatus_Set("")
  Me.Caption = Name_App_Short
  Me.width = 9600
  Me.height = 7600
  Me.picGraph.visible = False
  Call CenterOnScreen(Me)
  frmPrint_loading_now = True
  Set ThisGraph = New GraphControl
  Set ThisGraph.handle_ctlPicture = picGraph
  Call ThisGraph.CreateGraph("", "", "")
  Call file_new
  '
  ' POPULATE LAST-FEW-FILES LIST.
  '
  Call OldFileList_Populate( _
      1, _
      frmMain.mnuFileItem(199), _
      frmMain.mnuFileItem(191), _
      frmMain.mnuFileItem(192), _
      frmMain.mnuFileItem(193), _
      frmMain.mnuFileItem(194))
      
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If (file_query_unload() = False) Then
    Cancel = True
  End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Call frmMain_Close_All_Windows
  Call unitsys_unregister_all_on_form(Me)
End Sub



Private Sub mnuFileItem_Click(Index As Integer)

  Select Case Index
    Case 0:      'New
      If (file_query_unload()) Then
        Call Avoid_Weird_Focus_Problem
        Call file_new
      End If
    Case 1:      'Open ...
      If (file_query_unload()) Then
        Call Avoid_Weird_Focus_Problem
        Call File_OpenAs("")
      End If
    Case 2:      'Save
      If (Current_Filename = "") Then
        Call Avoid_Weird_Focus_Problem
        Call File_SaveAs("")
      Else
        Call Avoid_Weird_Focus_Problem
        Call File_SaveAs(Current_Filename)
      End If
    Case 3:      'Save As ...
      Call Avoid_Weird_Focus_Problem
      Call File_SaveAs("")
    Case 6:       'Select printer
      CommonDialog1.Copies = 1
      CommonDialog1.ShowPrinter
    Case 7:       'Print
      If nowproj.dyestudy_count = 0 Then
        Call Show_Message("There is no data to print", vbExclamation, App.title)
      Else
        If nowproj.DyeStudy(1).time = "" Then
            Call Show_Message("There is no data to print", vbExclamation, App.title)
        Else
            Call Print_DyeStudy
        End If
      End If
    Case 191 To 194:      'Last-few-files list
      If (file_query_unload()) Then
        If (mnuFileItem(Index).visible) Then
          Call File_OpenAs(OldFiles(1, Index - 190))
        End If
      End If
    Case 200:     'Exit
      'NOTE: MDIForm_QueryUnload() TAKES CAKE OF THIS.
      'If we do it here, _two_ message boxes will pop up
      'when the user has data which needs saving !
      'If (file_query_unload()) Then
      '  Unload Me
      'End If
      Unload Me
  End Select
End Sub
Private Sub mnuHelpItem_Click(Index As Integer)
Dim fn_this As String
  Select Case Index
   Case 80:
      fn_this = App.Path & "\dbase\readme.txt"
      If (FileExists(fn_this) = False) Then
        Call Show_Message("The file `" & fn_this & "` is missing.", _
          vbExclamation, App.title)
        Exit Sub
      End If
      Call Launch_Notepad(fn_this)
    Case 85:    'VIEW DISCLAIMER.
      'SHOW THE DISCLAIMER WINDOW.
      splash_mode = 101
      splash_button_pressed = 1
      frmSplash.Show 1
    Case 99:    'ABOUT.
      frmAbout.Show 1
  End Select
End Sub

Private Sub txtData_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = Global_NumericKeyPress(KeyAscii)
End Sub

Private Sub Form_Resize()
Dim XXX As Long
Dim YYY As Long
Dim USE_MARGIN As Long
  USE_MARGIN = 100
  'If (frmPrint_loading_now) Then Exit Sub
  If (Me.WindowState = 1) Then
    'CANNOT RESIZE WHEN MINIMIZED; EXIT OUTTA HERE.
    Exit Sub
  End If
  '
  ' RESIZE fraMain.
  '
  XXX = Me.ScaleWidth - USE_MARGIN * 2
  If (XXX < 100) Then XXX = 100
  YYY = Me.ScaleHeight - sspBottom.height - USE_MARGIN * 2
  If (YYY < 100) Then YYY = 100
  fraMain.Move _
      USE_MARGIN, _
      USE_MARGIN, _
      XXX, _
      YYY
    '
    ' RESIZE ssframe_GraphHolder.
    '
    XXX = fraMain.width - ssframe_GraphHolder.left - USE_MARGIN
    If (XXX < 100) Then XXX = 100
    YYY = fraMain.height - ssframe_GraphHolder.top - USE_MARGIN
    If (YYY < 100) Then YYY = 100
    ssframe_GraphHolder.Move _
        ssframe_GraphHolder.left, _
        ssframe_GraphHolder.top, _
        XXX, _
        YYY
      '
      ' RESIZE picGraph.
      '
      XXX = ssframe_GraphHolder.width - USE_MARGIN * 2
      If (XXX < 100) Then XXX = 100
      YYY = ssframe_GraphHolder.height - USE_MARGIN * 2
      If (YYY < 100) Then YYY = 100
      picGraph.Move _
          USE_MARGIN, _
          USE_MARGIN, _
          XXX, _
          YYY
  '
  ' ACTUALLY REPLOT THE GRAPH.
  '
  Call refresh_frmMain
  
'  If (Me.width < 6000) Then Me.width = 6000
'  If (Me.height < 3000) Then Me.height = 3000
'  'RESIZE ssframe_GraphHolder.
'  XXX = Me.width - (Me.width - Me.ScaleWidth) - _
'      ssframe_GraphHolder.left - USE_MARGIN
'  If (XXX < 1000) Then
'    XXX = 1000
'  End If
'  ssframe_GraphHolder.width = XXX
'
''  XXX = Me.height - (Me.height - Me.ScaleHeight) - _
''      StatusBar1.height - ssframe_GraphHolder.top - USE_MARGIN
'  XXX = Me.height - (Me.height - Me.ScaleHeight) - _
'      ssframe_GraphHolder.top - USE_MARGIN
'
'  If (XXX < 1000) Then
'    XXX = 1000
'  End If
'  ssframe_GraphHolder.height = XXX
'  'RESIZE picGraph.
'  XXX = ssframe_GraphHolder.width - picGraph.left * 2
'  If (XXX < 1000) Then XXX = 1000
'  picGraph.width = XXX
'  XXX = ssframe_GraphHolder.height - picGraph.top * 2
'  If (XXX < 1000) Then XXX = 1000
'  picGraph.height = XXX
'  Call Redraw_Graph
End Sub


Sub Redraw_Graph()
ReDim data_x(10) As Double
ReDim data_y(10) As Double
Dim i As Integer
Dim j As Integer
Dim num_rows As Integer
Dim ThisColor As Integer

'  'ABORT IF NOT READY TO PLOT.
'  If (READY_TO_PLOT = False) Then
'    Exit Sub
'  End If
'  Screen.MousePointer = 11
'  'REMOVE ALL EXISTING GRAPH DATA.
'  Call ThisGraph.DeleteAllSeries
'  'DETERMINE INDEXES OF SERIES TO PLOT.
'  Predicted_Theta = UBound(Predicted, 1)
'  Predicted_E = UBound(Predicted, 2)
'  num_rows = Prj.Predicted_count
'  ReDim series_link(1 To Num_Series)
'  For i = 1 To Num_Series
'    series_link(i) = 0
'    For j = 1 To num_rows
'      If (Trim$(UCase$(Predicted(j).Predicted_Theta)) = _
'          Trim$(UCase$(chkPlot(i).Tag))) Then
'        series_link(i) = j
'        Exit For
'      End If
'    Next j
'  Next i
'  'PLOT THE DATA.
'  ThisColor = 8
'  For i = 1 To Num_Series
'    If (chkPlot(i).Value = 0) Then
'      'DO NOT PLOT THIS SERIES.
'      chkPlot(i).ForeColor = QBColor(0)
'    Else
'      'PLOT THIS SERIES.
'      'PLOT THIS SERIES STEP #1: TRANSFER DATA TO LOCAL X-Y ARRAYS.
'      ReDim data_x(1 To num_rows)
'      ReDim data_y(1 To num_rows)
'      For j = 1 To num_rows
'
'      Next j
'  Next i
''
''  'MISCELLANEOUS FORMATTING STUFF.
'' 'Call graph.Change_X_Number_Format("0.00e+00")
'' 'Call graph.Change_Y_Number_Format("0.00e+00")
'  Call ThisGraph.Refresh_Graph
'  Screen.MousePointer = 0
''
End Sub

