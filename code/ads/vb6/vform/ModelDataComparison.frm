VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0842D103-1E19-101B-9AAF-1A1626551E7C}#1.0#0"; "GRAPH32.OCX"
Begin VB.Form frmModelDataComparison 
   Caption         =   "Data Comparison"
   ClientHeight    =   6795
   ClientLeft      =   345
   ClientTop       =   1845
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   9480
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   9480
      ScaleHeight     =   795
      ScaleWidth      =   1275
      TabIndex        =   13
      Top             =   840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin Threed.SSCommand cmdClose 
      Height          =   435
      Left            =   7950
      TabIndex        =   11
      Top             =   60
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   767
      _StockProps     =   78
      Caption         =   "&Close"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   615
      Left            =   90
      TabIndex        =   7
      Top             =   420
      Width           =   1155
      _Version        =   65536
      _ExtentX        =   2037
      _ExtentY        =   1085
      _StockProps     =   14
      Caption         =   "C Units"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Begin VB.ComboBox cboCUnits 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   240
         Width           =   915
      End
   End
   Begin VB.ComboBox cboGraphType 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6510
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   60
      Width           =   1335
   End
   Begin VB.ComboBox cboCompo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1890
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   60
      Width           =   3195
   End
   Begin VB.ComboBox cboGrid 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6510
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   420
      Width           =   1335
   End
   Begin GraphLib.Graph grpBreak 
      Height          =   5655
      Left            =   90
      TabIndex        =   3
      Top             =   1080
      Width           =   9315
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   2646
      _StockProps     =   96
      BorderStyle     =   1
      RandomData      =   1
      ColorData       =   0
      ExtraData       =   0
      ExtraData[]     =   0
      FontFamily      =   4
      FontSize        =   4
      FontSize[0]     =   200
      FontSize[1]     =   150
      FontSize[2]     =   100
      FontSize[3]     =   100
      FontStyle       =   4
      GraphData       =   0
      GraphData[]     =   0
      LabelText       =   0
      LegendText      =   0
      PatternData     =   0
      SymbolData      =   0
      XPosData        =   0
      XPosData[]      =   0
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   615
      Left            =   1230
      TabIndex        =   9
      Top             =   420
      Width           =   1155
      _Version        =   65536
      _ExtentX        =   2037
      _ExtentY        =   1085
      _StockProps     =   14
      Caption         =   "T Units"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Begin VB.ComboBox cboTUnits 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   240
         Width           =   915
      End
   End
   Begin Threed.SSCommand cmdPrint 
      Height          =   435
      Left            =   7950
      TabIndex        =   12
      Top             =   480
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   767
      _StockProps     =   78
      Caption         =   "&Print Screen"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Select a component:"
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
      Height          =   255
      Left            =   -30
      TabIndex        =   6
      Top             =   120
      Width           =   1875
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Plot Patterns:"
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
      Height          =   255
      Left            =   5250
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Grid Style:"
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
      Height          =   255
      Left            =   5130
      TabIndex        =   4
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "frmModelDataComparison"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Dim Cin() As Double, Td()  As Double, Cd() As Double, Fmin() As Double
Dim UnloadMe As Integer

Dim PopulatingScrollboxes As Integer

'---- Conc Units
Const CBOCUNITS_CC0 = 0
Const CBOCUNITS_mg_L = 1
Const CBOCUNITS_ug_L = 2

'---- Time Units
Const CBOTUNITS_days = 0
Const CBOTUNITS_BVF = 1
Const CBOTUNITS_VTM = 2





Const frmModelDataComparison_declarations_end = True


Private Sub Populate_Scrollboxes()
Dim i As Integer
  PopulatingScrollboxes = True
  'For i = 0 To 2
  '  cboDataset(i).Clear
  '  cboDataset(i).AddItem "Off"
  '  cboDataset(i).AddItem "Symbols"
  '  cboDataset(i).AddItem "Lines"
  '  cboDataset(i).AddItem "Symbols and Lines"
  'Next i
  cboCUnits.Clear
  cboCUnits.AddItem "C/C0"
  cboCUnits.AddItem "mg/L"
  cboCUnits.AddItem Chr$(181) & "g/L"
  cboTUnits.Clear
  cboTUnits.AddItem "days"
  cboTUnits.AddItem "BVF"
  cboTUnits.AddItem "VTM"
  cboGraphType.Clear
  cboGraphType.AddItem "Symbols"
  cboGraphType.AddItem "Lines"
  cboGrid.AddItem "None"
  cboGrid.AddItem "Horizontal"
  cboGrid.AddItem "Vertical"
  cboGrid.AddItem "Both"
  cboCompo.Clear
  Select Case frmCompareData_WhichSet
    Case frmCompareData_WhichSet_PSDM
      For i = 1 To Results.NComponent
        cboCompo.AddItem Trim$(Results.Component(i).Name)
      Next i
    Case frmCompareData_WhichSet_CPHSDM
      cboCompo.AddItem Trim$(CPM_Results.Component.Name)
  End Select
  '---- Read in INI settings
  'cboDataset(0).ListIndex = 1
  'cboDataset(1).ListIndex = 1
  'cboDataset(2).ListIndex = 1
  cboCUnits.ListIndex = 0
  cboTUnits.ListIndex = 0
  cboGraphType.ListIndex = 0
  cboGrid.ListIndex = 0
  cboCompo.ListIndex = 0
  Call UserPrefs_Load
  PopulatingScrollboxes = False
End Sub




Private Sub cboCompo_Click()
  If (Not PopulatingScrollboxes) Then
    Call Draw_Curves(cboCompo.ListIndex + 1)
    'lblErrorC = Format$(Fmin(cboCompo.ListIndex + 1), "0.0000E+00")
  End If
End Sub
Private Sub cboCUnits_Click()
  If (Not PopulatingScrollboxes) Then
    Call Draw_Curves(cboCompo.ListIndex + 1)
  End If
End Sub
Private Sub cboDataset_Click(Index As Integer)
  'If (Not PopulatingScrollboxes) Then
  '  Call Draw_Curves(cboCompo.ListIndex + 1)
  'End If
End Sub
Private Sub cboGraphType_Click()
  If (Not PopulatingScrollboxes) Then
    Select Case cboGraphType.ListIndex
      Case 0 'Symbols
        grpBreak.GraphStyle = 1
      Case 1 'Lines
        grpBreak.GraphStyle = 4
    End Select
    grpBreak.DrawMode = 2
  End If
End Sub
Private Sub cboGrid_Click()
  If (Not PopulatingScrollboxes) Then
    grpBreak.GridStyle = cboGrid.ListIndex
    grpBreak.DrawMode = 2
  End If
End Sub
Private Sub cboTUnits_Click()
  If (Not PopulatingScrollboxes) Then
    Call Draw_Curves(cboCompo.ListIndex + 1)
  End If
End Sub


Private Sub cmdClose_Click()
  Unload Me
End Sub


Private Sub cmdPrint_Click()
    Set Picture1.Picture = CaptureActiveWindow()
    PrintPictureToFitPage Printer, Picture1.Picture
    Printer.EndDoc
    ' Set focus back to form.
    Me.SetFocus

'''Dim i As Integer
'''Dim H As Single
'''Dim W As Single
'''  'printer.ScaleLeft = -1080  'Set a 3/4-inch margin
'''  'printer.ScaleTop = -1080
'''  'printer.CurrentX = 0
'''  'printer.CurrentY = 0
'''  '
'''  '  printer.FontSize = 12
'''  '  printer.FontBold = True
'''  '  printer.FontUnderline = True
'''  '  printer.Print "Input data for the Plug-Flow Pore And Surface Diffusion Model"
'''  '  printer.FontSize = 10
'''  '  printer.FontBold = False
'''  '  printer.FontUnderline = False
'''  '  '-- Print Filename
'''  '  printer.Print
'''  '  printer.Print "From Data File: "; Filename
'''  '---- Print the graph ------------------------
'''  For i = 1 To Number_Component
'''    grpBreak.ThisPoint = i
'''    grpBreak.PatternData = i - 1
'''  Next i
'''  H = grpBreak.Height
'''  W = grpBreak.Width
'''  grpBreak.Visible = False 'Hide it before printing
'''  If (Printer.Width < Printer.Height) Then
'''    grpBreak.Height = CSng(Printer.Height / 2#)
'''    grpBreak.Width = Printer.Width
'''  Else
'''    grpBreak.Height = Printer.Height
'''    grpBreak.Width = Printer.Width
'''  End If
'''  grpBreak.PrintStyle = 2
'''  grpBreak.DrawMode = 5
'''  grpBreak.Height = H
'''  grpBreak.Width = W
'''  grpBreak.Visible = True
'''  grpBreak.PrintStyle = 2
'''  grpBreak.DrawMode = 2
'''  Printer.EndDoc

End Sub


Private Sub Draw_Curves(Component_Index As Integer)
Dim i As Integer, J As Integer
Dim Data_Max As Double, t_factor As Double, Bottom_Title As String
Dim c_factor As Double, Left_Title As String
Dim bigger As Integer
Dim SameX, SameY As Double
Dim LastPointI As Integer
Dim bed_data As BedPropertyType
Dim comp_data As ComponentPropertyType
Dim num_model_points As Integer

  Select Case frmCompareData_WhichSet
    Case frmCompareData_WhichSet_PSDM
      bed_data = Results.Bed
      comp_data = Results.Component(Component_Index)
      num_model_points = Results.npoints
    Case frmCompareData_WhichSet_CPHSDM
      bed_data = CPM_Results.Bed
      comp_data = CPM_Results.Component
      num_model_points = 100
  End Select
  
  Select Case cboTUnits.ListIndex
    Case CBOTUNITS_days:
      t_factor = 1# / 60# / 24#  'mn -> days
      Bottom_Title = "Time(days)"
    Case CBOTUNITS_BVF:
      t_factor = 60# * bed_data.Flowrate / bed_data.length / PI / (bed_data.Diameter / 2#) ^ 2
      Bottom_Title = "Bed Volumes Treated"
    Case CBOTUNITS_VTM:
      t_factor = 60# * bed_data.Flowrate / bed_data.Weight
      Bottom_Title = "m" & Chr$(179) & " treated per kg of adsorbent"
  End Select

  Select Case cboCUnits.ListIndex
    Case CBOCUNITS_CC0:
      c_factor = 1#
      Left_Title = "C/C0"
    Case CBOCUNITS_mg_L:
      c_factor = comp_data.InitialConcentration
      Left_Title = "mg/L"
    Case CBOCUNITS_ug_L:
      c_factor = comp_data.InitialConcentration * 1000#
      Left_Title = Chr$(181) & "g/L"
  End Select

   'Define Graph
   If (Number_Influent_Points = 0) Then
     grpBreak.NumSets = 2
   Else
     grpBreak.NumSets = 3
   End If
   grpBreak.GraphType = 6 'Lines/Symbols
   grpBreak.GraphStyle = 1 'Symbols

'   grpBreak.ThisSet = 1
'   grpBreak.NumPoints = Results.NPoints
'   grpBreak.ThisSet = 2
'   grpBreak.NumPoints = NData_Points
   
   ' The following code where grpBreak.NumPoints is set is a rather
   ' unfortunate kludge, in my opinion.  I could find no other way to
   ' convince/force Visual Basic's graphical interface to accept two sets
   ' of data that were of two different sizes, so I determined which one
   ' was the smaller set and then filled the remainer of the smaller set
   ' with copies of the last data point in it (X,Y) (note, the default
   ' is for the data to hook back to the point (0,0) at the end of its
   ' plotting due to the fact that, by default, the (X,Y) data points
   ' that are unspecified are filled with 0's).
   ' -- If possible, it would be nice to replace this with something
   ' more elegant, but hey, it works. -- Eric J. Oman
   If (num_model_points > NData_Points) Then
     bigger = num_model_points
   Else
     bigger = NData_Points
   End If
   If (Number_Influent_Points > bigger) Then
     bigger = Number_Influent_Points
   End If

   grpBreak.ThisSet = 1
   grpBreak.NumPoints = bigger
   grpBreak.ThisSet = 2
   grpBreak.NumPoints = bigger
   If (Number_Influent_Points = 0) Then
     'Do nothing
   Else
     grpBreak.ThisSet = 3
     grpBreak.NumPoints = bigger
   End If
   
   grpBreak.SymbolData = 2 'triangle
   grpBreak.SymbolData = 6 'square
   If (Number_Influent_Points = 0) Then
     'Do nothing
   Else
     grpBreak.SymbolData = 8 'diamond
   End If

   grpBreak.ColorData = 9 'Blue
   grpBreak.ColorData = 12 'Red
   If (Number_Influent_Points = 0) Then
     'Do nothing
   Else
     grpBreak.ColorData = 10 'Green
   End If
   
   grpBreak.PatternData = 1
   grpBreak.PatternData = 1
   If (Number_Influent_Points = 0) Then
     'Do nothing
   Else
     grpBreak.PatternData = 1
   End If

   grpBreak.AutoInc = 0  'No autoincrementation
    
'**************************************************************
'   grpBreak.ThisSet = 1
'   For I = 1 To grpBreak.NumPoints
'     grpBreak.ThisPoint = I
'     If Cin(Component_Index, I) < 0 Then
'       grpBreak.GraphData = 0#
'     Else
'       grpBreak.GraphData = Cin(Component_Index, I) 'Results.CP(Component, I)
'     End If
'     grpBreak.ThisPoint = I
'     grpBreak.LabelText = ""
'     grpBreak.ThisPoint = I
'     grpBreak.XPosData = Td(I) * factor 'X_Values(I)
'   Next I
'   grpBreak.ThisPoint = 1
'   grpBreak.LegendText = Trim$(Results.Component(Component_Index).Name)
   
   '---- I. Display Effluent Prediction
   grpBreak.ThisSet = 1
   Select Case frmCompareData_WhichSet
     Case frmCompareData_WhichSet_PSDM
       For i = 1 To num_model_points
         grpBreak.ThisPoint = i
         If (Results.CP(Component_Index, i) < 0) Then
           grpBreak.GraphData = 0#
         Else
           grpBreak.GraphData = Results.CP(Component_Index, i) * c_factor
         End If
         ''''grpBreak.LabelText = ""
         grpBreak.XPosData = Results.T(i) * t_factor        'X_Values(I)
       Next i
     Case frmCompareData_WhichSet_CPHSDM
       For i = 1 To num_model_points
         grpBreak.ThisPoint = i
         If (CPM_Results.C_Over_C0(i) < 0) Then
           grpBreak.GraphData = 0#
         Else
           grpBreak.GraphData = CPM_Results.C_Over_C0(i) * c_factor
         End If
         ''''grpBreak.LabelText = ""
         grpBreak.XPosData = CPM_Results.T(i) * 24# * 60# * t_factor
       Next i
   End Select

   grpBreak.ThisPoint = 1
   grpBreak.LegendText = "Effluent Prediction"
   'grpBreak.LegendText = Trim$(Results.Component(Component_Index).Name)

   '---- II. Display Effluent Data
   grpBreak.ThisSet = 2
   For i = 1 To NData_Points
     grpBreak.ThisPoint = i
     If (C_Data_Points(Component_Index, i) < 0) Then
       grpBreak.GraphData = 0#
     Else
       grpBreak.GraphData = C_Data_Points(Component_Index, i) * c_factor
     End If
     ''''grpBreak.LabelText = ""
     grpBreak.XPosData = T_Data_Points(i) * 24# * 60# * t_factor
   Next i
   grpBreak.ThisPoint = 2
   grpBreak.LegendText = "Effluent Data"

   '---- III. Display Influent Data
   If (Number_Influent_Points = 0) Then
     'Do nothing
   Else
     grpBreak.ThisSet = 3
     For i = 1 To Number_Influent_Points
       grpBreak.ThisPoint = i
       If (C_Influent(Component_Index, i) < 0) Then
         grpBreak.GraphData = 0#
       Else
         grpBreak.GraphData = C_Influent(Component_Index, i) / comp_data.InitialConcentration * c_factor
       End If
       ''''grpBreak.LabelText = ""
       grpBreak.XPosData = T_Influent(i) * t_factor
     Next i
     grpBreak.ThisPoint = 3
     grpBreak.LegendText = "Influent Data"
   End If

   '---- Run the kludge mentioned above.
   If (bigger > NData_Points) Then
     grpBreak.ThisSet = 2
     LastPointI = NData_Points
     
     SameX = T_Data_Points(LastPointI) * 24# * 60# * t_factor
     SameY = C_Data_Points(Component_Index, LastPointI) * c_factor
     For i = LastPointI + 1 To bigger
       grpBreak.ThisPoint = i
       grpBreak.GraphData = SameY
       ''''grpBreak.ThisPoint = i
       grpBreak.XPosData = SameX
     Next i
   End If
   Select Case frmCompareData_WhichSet
     Case frmCompareData_WhichSet_PSDM
       If (bigger > num_model_points) Then
         grpBreak.ThisSet = 1
         LastPointI = num_model_points
         SameX = Results.T(LastPointI) * t_factor
         SameY = Results.CP(Component_Index, LastPointI) * c_factor
         For i = LastPointI + 1 To bigger
           grpBreak.ThisPoint = i
           grpBreak.GraphData = SameY
           ''''grpBreak.ThisPoint = i
           grpBreak.XPosData = SameX
         Next i
       End If
     Case frmCompareData_WhichSet_CPHSDM
       If (bigger > num_model_points) Then
         grpBreak.ThisSet = 1
         LastPointI = num_model_points
         SameX = CPM_Results.T(LastPointI) * 24# * 60# * t_factor
         SameY = CPM_Results.C_Over_C0(LastPointI) * c_factor
         For i = LastPointI + 1 To bigger
           grpBreak.ThisPoint = i
           grpBreak.GraphData = SameY
           ''''grpBreak.ThisPoint = i
           grpBreak.XPosData = SameX
         Next i
       End If
   End Select
   If (Number_Influent_Points = 0) Then
     'Do nothing
   Else
     If (bigger > Number_Influent_Points) Then
       grpBreak.ThisSet = 3
       LastPointI = Number_Influent_Points
       SameX = T_Influent(LastPointI) * t_factor
       SameY = C_Influent(Component_Index, LastPointI) / Component(Component_Index).InitialConcentration * c_factor
       For i = LastPointI + 1 To bigger
         grpBreak.ThisPoint = i
         grpBreak.GraphData = SameY
         ''''grpBreak.ThisPoint = i
         grpBreak.XPosData = SameX
       Next i
     End If
   End If

   grpBreak.PatternedLines = 0
   Data_Max = 0
   For J = 1 To grpBreak.NumSets
     grpBreak.ThisSet = J
     For i = 1 To grpBreak.NumPoints
       grpBreak.ThisPoint = i
       If grpBreak.GraphData > Data_Max Then
         Data_Max = grpBreak.GraphData
       End If
     Next i
   Next J
   
   grpBreak.YAxisMax = (Int(Data_Max * 10# + 1)) / 10#
   grpBreak.YAxisTicks = 4
   'grpBreak.GridStyle = 0

   grpBreak.YAxisStyle = 2
   grpBreak.YAxisMin = 0#
   grpBreak.BottomTitle = Bottom_Title
    
   grpBreak.LeftTitle = Left_Title
   grpBreak.DrawMode = 2

End Sub


Private Sub Form_Activate()
  If UnloadMe Then Unload Me
End Sub
Private Sub Form_Load()
Dim J As Integer, i As Integer
  Me.Caption = frmCompareData_caption
  Call Populate_Scrollboxes
  Call CenterOnForm(Me, frmMain)
  ''''Move frmPFPSDM.Left + (frmPFPSDM.Width / 2) - (frmShow_Data_And_Prediction.Width / 2), frmPFPSDM.Top + (frmPFPSDM.Height / 2) - (frmShow_Data_And_Prediction.Height / 2)
  'If Obj_Function() Then
  Screen.MousePointer = 11
  Call Draw_Curves(cboCompo.ListIndex + 1)
  Call cboGraphType_Click   'Wake cboGraphType up
  Screen.MousePointer = 0
  UnloadMe = False
  Call cboGrid_Click
  'Else
  '  UnloadMe = True
  'End If
End Sub
Private Sub Form_Resize()
  'If WindowState = 1 Then
  '  frmPFPSDM.WindowState = 1
  '  frmPlantData.WindowState = 1
  'End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Call UserPrefs_Save
End Sub


Private Sub UserPrefs_Load()
Dim X As Long
  On Error GoTo err_DATAPRED_UserPrefs_Load
  X = CLng(INI_Getsetting("DATAPRED_cboCUnits"))
  If ((X >= 0) And (X <= cboCUnits.ListCount - 1)) Then
    cboCUnits.ListIndex = X
  End If
  X = CLng(INI_Getsetting("DATAPRED_cboTUnits"))
  If ((X >= 0) And (X <= cboTUnits.ListCount - 1)) Then
    cboTUnits.ListIndex = X
  End If
  X = CLng(INI_Getsetting("DATAPRED_cboGraphType"))
  If ((X >= 0) And (X <= cboGraphType.ListCount - 1)) Then
    cboGraphType.ListIndex = X
  End If
  X = CLng(INI_Getsetting("DATAPRED_cboGrid"))
  If ((X >= 0) And (X <= cboGrid.ListCount - 1)) Then
    cboGrid.ListIndex = X
  End If
  Exit Sub
resume_err_DATAPRED_UserPrefs_Load:
  Call UserPrefs_Save
  Exit Sub
err_DATAPRED_UserPrefs_Load:
  Resume resume_err_DATAPRED_UserPrefs_Load
End Sub
Private Sub UserPrefs_Save()
Dim X As Long
  X = cboCUnits.ListIndex
  Call INI_PutSetting("DATAPRED_cboCUnits", Trim$(CStr(X)))
  X = cboTUnits.ListIndex
  Call INI_PutSetting("DATAPRED_cboTUnits", Trim$(CStr(X)))
  X = cboGraphType.ListIndex
  Call INI_PutSetting("DATAPRED_cboGraphType", Trim$(CStr(X)))
  X = cboGrid.ListIndex
  Call INI_PutSetting("DATAPRED_cboGrid", Trim$(CStr(X)))
End Sub





'Private Function Obj_Function() As Integer
'Dim i, J As Integer
'Dim ncomp As Long, ndata  As Long, np As Long, temp As String, Error_Code As Integer
'ReDim Fmin(Results.NComponent) As Double
'ReDim TP(Results.npoints) As Double, CP(Results.NComponent, Results.npoints) As Double
'ReDim Td(NData_Points) As Double, Cd(Results.NComponent, NData_Points) As Double, Cin(Results.NComponent, NData_Points) As Double
'
'  ncomp = CLng(Results.NComponent)
'  ndata = CLng(NData_Points)
'  np = CLng(Results.npoints)
'
'  For i = 1 To Results.npoints
'    TP(i) = Results.T(i)
'    For J = 1 To Results.NComponent
'      CP(J, i) = Results.CP(J, i)
'    Next J
'  Next i
'  For i = 1 To NData_Points
'    Td(i) = T_Data_Points(i) * 24# * 60#
'    For J = 1 To Results.NComponent
'      Cd(J, i) = C_Data_Points(J, i)
'    Next J
'  Next i
'
'On Error GoTo Error_In_OBJFUN
'  'Call OBJFUN(ncomp, ndata, np, Tp(1), CP(1, 1), Td(1), Cd(1, 1), Cin(1, 1), Fmin(1))
'  Obj_Function = True
'  Exit Function
'
'Error_In_OBJFUN:
'  Error_Code = Err
'  temp = "Error " & Format$(Error_Code, "0") & " : " & Error$(Error_Code)
'  MsgBox "Fatal Error with OBJFUN.DLL. Calculations Stoppped." & Chr$(13) & temp, MB_ICONEXCLAMATION, AppName_For_Display_long
'  Obj_Function = False
'  Resume Exit_Obj_Function
'Exit_Obj_Function:
'End Function

