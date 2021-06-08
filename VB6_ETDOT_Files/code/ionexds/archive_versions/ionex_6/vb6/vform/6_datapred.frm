VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Begin VB.Form frmShow_Data_And_Prediction 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Data Comparison"
   ClientHeight    =   6525
   ClientLeft      =   1350
   ClientTop       =   2415
   ClientWidth     =   8100
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
   Icon            =   "6_datapred.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6525
   ScaleWidth      =   8100
   Begin VB.ComboBox cboGraphType 
      Height          =   315
      Left            =   5520
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   1200
      Width           =   2175
   End
   Begin VB.PictureBox grpBreak 
      BackColor       =   &H8000000E&
      Height          =   4455
      Left            =   240
      ScaleHeight     =   4395
      ScaleWidth      =   7515
      TabIndex        =   6
      Top             =   1920
      Width           =   7575
   End
   Begin Threed.SSFrame fraGraphHolder 
      Height          =   4815
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   7815
      _Version        =   65536
      _ExtentX        =   13785
      _ExtentY        =   8493
      _StockProps     =   14
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
   Begin Threed.SSCommand cmdExit 
      Height          =   615
      Left            =   6000
      TabIndex        =   4
      Top             =   240
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   1085
      _StockProps     =   78
      Caption         =   "E&xit"
   End
   Begin Threed.SSFrame fra3D1 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      _Version        =   65536
      _ExtentX        =   8705
      _ExtentY        =   2566
      _StockProps     =   14
      Caption         =   "Select one component"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ComboBox cboCompo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   4635
      End
      Begin VB.Label lblObjective 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3360
         TabIndex        =   3
         Top             =   900
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sum of the squared errors:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   60
         TabIndex        =   2
         Top             =   900
         Width           =   3135
      End
   End
End
Attribute VB_Name = "frmShow_Data_And_Prediction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Option Base 1

Dim Cin() As Double, TD()  As Double, CD() As Double, FMIN() As Double
Dim UnloadMe As Integer

Private Sub cboCompo_Click()
    Call Draw_Curves(cboCompo.ListIndex + 1)
    lblObjective = Format$(FMIN(cboCompo.ListIndex + 1), "0.0000E+00")
End Sub

Private Sub cboGraphType_Click()
  Select Case cboGraphType.ListIndex
    Case 0 'Symbols
'      grpBreak.GraphStyle = 1
    Case 1 'Lines
'      grpBreak.GraphStyle = 4
  End Select
'  grpBreak.DrawMode = 2
End Sub

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub Draw_Curves(Component_Index As Integer)
Dim i As Integer, j As Integer
Dim Data_Max As Double, factor As Double, Bottom_Title As String

    If TimeUnitsOnGraphs = 0 Then   'min
       factor = 1#
       Bottom_Title = "Time (min)"
    ElseIf TimeUnitsOnGraphs = 1 Then   'sec
       factor = 1# * 60#
       Bottom_Title = "Time (sec)"
    ElseIf TimeUnitsOnGraphs = 2 Then   'hrs
       factor = 1# / 60#
       Bottom_Title = "Time (hrs)"
    ElseIf TimeUnitsOnGraphs = 3 Then   'days
       factor = 1# / 60# / 24#
       Bottom_Title = "Time (days)"
    End If

    'Define Graph
'    grpBreak.NumSets = 2
'    grpBreak.GraphType = 6 'Lines/Symbols
'    grpBreak.GraphStyle = 1 'Symbols
'
'    grpBreak.ThisSet = 1
'    grpBreak.NumPoints = NData_Points
'    grpBreak.ThisSet = 2
'    grpBreak.NumPoints = NData_Points

'    grpBreak.SymbolData = 7 'Square
'    grpBreak.SymbolData = 9 'Diamond
'
'    grpBreak.ColorData = 9 'Blue
'    grpBreak.ColorData = 12 'Red
'
'    grpBreak.PatternData = 1
'    grpBreak.PatternData = 1

'    grpBreak.AutoInc = 0  'No autoincrementation
    
'**************************************************************
'      grpBreak.ThisSet = 1
'      For i = 1 To grpBreak.NumPoints
'         grpBreak.ThisPoint = i
'         If Cin(Component_Index, i) < 0 Then
'           grpBreak.GraphData = 0#
'         Else
'           grpBreak.GraphData = Cin(Component_Index, i) 'Results.CP(Component, I)
'         End If
'         grpBreak.ThisPoint = i
'         grpBreak.LabelText = ""
'         grpBreak.ThisPoint = i
'         grpBreak.XPosData = TD(i) * factor 'X_Values(I)
'       Next i
'       grpBreak.ThisPoint = 1
'       grpBreak.LegendText = Trim$(Results.Component(Component_Index).Name)
'
'
'       grpBreak.ThisSet = 2
'       For i = 1 To grpBreak.NumPoints
'         grpBreak.ThisPoint = i
'         If CD(Component_Index, i) < 0 Then
'           grpBreak.GraphData = 0#
'         Else
'           grpBreak.GraphData = CD(Component_Index, i)
'         End If
'         grpBreak.ThisPoint = i
'         grpBreak.LabelText = ""
'         grpBreak.ThisPoint = i
'         grpBreak.XPosData = TD(i) * factor
'       Next i
'       grpBreak.ThisPoint = 2
'       grpBreak.LegendText = "Data Points"
'
''********************************************************************
'    grpBreak.PatternedLines = 0
'    Data_Max = 0
'    For j = 1 To grpBreak.NumSets
'      grpBreak.ThisSet = j
'    For i = 1 To grpBreak.NumPoints
'     grpBreak.ThisPoint = i
'       If grpBreak.GraphData > Data_Max Then
'         Data_Max = grpBreak.GraphData
'        End If
'       Next i
'    Next j
'    grpBreak.YAxisMax = (Int(Data_Max * 10# + 1)) / 10#
'    grpBreak.YAxisTicks = 4
'    grpBreak.GridStyle = 0
'
'    grpBreak.YAxisStyle = 2
'    grpBreak.YAxisMin = 0#
'    grpBreak.BottomTitle = Bottom_Title
'
'    grpBreak.LeftTitle = "C/Ct"
'    grpBreak.DrawMode = 2

End Sub

Private Sub Form_Activate()
  If UnloadMe Then Unload Me
End Sub

Private Sub Form_Load()
Dim j As Integer, i As Integer

    top = Screen.height / 2 - height / 2
    left = Screen.width / 2 - width / 2

    cboGraphType.AddItem "Symbols"
    cboGraphType.AddItem "Lines"

    If Obj_Function() Then
        Screen.MousePointer = 11
        For i = 1 To Results.NComponent
          cboCompo.AddItem Trim$(Results.Component(i).Name)
        Next i
        cboCompo.ListIndex = 0
        Call Draw_Curves(cboCompo.ListIndex + 1)
        cboGraphType.ListIndex = 0
        Screen.MousePointer = 0
        UnloadMe = False
    Else
        UnloadMe = True
    End If
End Sub

Private Sub Form_Resize()
  If WindowState = 1 Then
    frmIonExchangeMain.WindowState = 1
    frmPlantData.WindowState = 1
  End If

End Sub

Private Function Obj_Function() As Integer
Dim i, j As Integer
Dim NCOMP As Long, NDATA  As Long, NP As Long, temp As String, Error_Code As Integer
ReDim FMIN(Results.NComponent) As Double
ReDim TP(Results.NPoints) As Double, CP(Results.NComponent, Results.NPoints) As Double
ReDim TD(NData_Points) As Double, CD(Results.NComponent, NData_Points) As Double, Cin(Results.NComponent, NData_Points) As Double

  NCOMP = CLng(Results.NComponent)
  NDATA = CLng(NData_Points)
  NP = CLng(Results.NPoints)

  For i = 1 To Results.NPoints
    TP(i) = Results.T(i)
    For j = 1 To Results.NComponent
      CP(j, i) = Results.CP(j, i)
    Next j
  Next i
  For i = 1 To NData_Points
    TD(i) = T_Data_Points(i)
    For j = 1 To Results.NComponent
      CD(j, i) = C_Data_Points(j, i)
    Next j
  Next i

On Error GoTo Error_In_OBJFUN
  Call OBJFUN(NCOMP, NDATA, NP, TP(1), CP(1, 1), TD(1), CD(1, 1), Cin(1, 1), FMIN(1))
  Obj_Function = True
  Exit Function

Error_In_OBJFUN:
  Error_Code = Err
  temp = "Error " & Format$(Error_Code, "0") & " : " & Error$(Error_Code)
  MsgBox "Fatal Error with OBJFUN.DLL. Calculations Stoppped." & Chr$(13) & temp, MB_ICONEXCLAMATION, App.title
  Obj_Function = False
  Resume Exit_Obj_Function
Exit_Obj_Function:
End Function

