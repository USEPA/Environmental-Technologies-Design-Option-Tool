VERSION 2.00
Begin Form frmShow_Data_And_Prediction 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Data Comparison"
   ClientHeight    =   6525
   ClientLeft      =   1350
   ClientTop       =   2415
   ClientWidth     =   8100
   Height          =   6930
   Icon            =   DATAPRED.FRX:0000
   Left            =   1290
   LinkTopic       =   "Form1"
   ScaleHeight     =   6525
   ScaleWidth      =   8100
   Top             =   2070
   Width           =   8220
   Begin ComboBox cboGraphType 
      Height          =   300
      Left            =   6480
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1200
      Width           =   1575
   End
   Begin GRAPH grpBreak 
      Height          =   4815
      Left            =   60
      TabIndex        =   3
      Top             =   1620
      Width           =   7935
   End
   Begin SSCommand cmdExit 
      Caption         =   "E&xit"
      Height          =   435
      Left            =   6540
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin SSFrame Frame3D1 
      Caption         =   "Select one component:"
      ForeColor       =   &H00000000&
      Height          =   1455
      Left            =   60
      ShadowColor     =   1  'Black
      TabIndex        =   0
      Top             =   60
      Width           =   6375
      Begin ComboBox cboCompo 
         Height          =   300
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   300
         Width           =   4635
      End
      Begin Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sum of the squared errors:"
         Height          =   255
         Left            =   300
         TabIndex        =   5
         Top             =   840
         Width           =   3135
      End
      Begin Label lblObjective 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   255
         Left            =   3600
         TabIndex        =   4
         Top             =   840
         Width           =   1155
      End
   End
End
Option Explicit
Option Base 1

Dim Cin() As Double, Td()  As Double, Cd() As Double, Fmin() As Double
Dim UnloadMe As Integer

Sub cboCompo_Click ()
    Call Draw_Curves(cboCompo.ListIndex + 1)
    lblObjective = Format$(Fmin(cboCompo.ListIndex + 1), "0.0000E+00")
End Sub

Sub cboGraphType_Click ()
  Select Case cboGraphType.ListIndex
    Case 0 'Symbols
      grpBreak.GraphStyle = 1
    Case 1 'Lines
      grpBreak.GraphStyle = 4
  End Select
  grpBreak.DrawMode = 2
End Sub

Sub cmdExit_Click ()
  Unload Me
End Sub

Sub Draw_Curves (Component_Index As Integer)
Dim I As Integer, J As Integer
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
    grpBreak.NumSets = 2
    grpBreak.GraphType = 6 'Lines/Symbols
    grpBreak.GraphStyle = 1 'Symbols

    grpBreak.ThisSet = 1
    grpBreak.NumPoints = NData_Points
    grpBreak.ThisSet = 2
    grpBreak.NumPoints = NData_Points

    grpBreak.SymbolData = 7 'Square
    grpBreak.SymbolData = 9 'Diamond

    grpBreak.ColorData = 9 'Blue
    grpBreak.ColorData = 12 'Red

    grpBreak.PatternData = 1
    grpBreak.PatternData = 1

    grpBreak.AutoInc = 0  'No autoincrementation
    
'**************************************************************
      grpBreak.ThisSet = 1
      For I = 1 To grpBreak.NumPoints
         grpBreak.ThisPoint = I
         If Cin(Component_Index, I) < 0 Then
           grpBreak.GraphData = 0#
         Else
           grpBreak.GraphData = Cin(Component_Index, I) 'Results.CP(Component, I)
         End If
         grpBreak.ThisPoint = I
         grpBreak.LabelText = ""
         grpBreak.ThisPoint = I
         grpBreak.XPosData = Td(I) * factor 'X_Values(I)
       Next I
       grpBreak.ThisPoint = 1
       grpBreak.LegendText = Trim$(Results.Component(Component_Index).Name)
       
    
       grpBreak.ThisSet = 2
       For I = 1 To grpBreak.NumPoints
         grpBreak.ThisPoint = I
         If Cd(Component_Index, I) < 0 Then
           grpBreak.GraphData = 0#
         Else
           grpBreak.GraphData = Cd(Component_Index, I)
         End If
         grpBreak.ThisPoint = I
         grpBreak.LabelText = ""
         grpBreak.ThisPoint = I
         grpBreak.XPosData = Td(I) * factor
       Next I
       grpBreak.ThisPoint = 2
       grpBreak.LegendText = "Data Points"

'********************************************************************
    grpBreak.PatternedLines = 0
    Data_Max = 0
    For J = 1 To grpBreak.NumSets
      grpBreak.ThisSet = J
    For I = 1 To grpBreak.NumPoints
     grpBreak.ThisPoint = I
       If grpBreak.GraphData > Data_Max Then
         Data_Max = grpBreak.GraphData
        End If
       Next I
    Next J
    grpBreak.YAxisMax = (Int(Data_Max * 10# + 1)) / 10#
    grpBreak.YAxisTicks = 4
    grpBreak.GridStyle = 0

    grpBreak.YAxisStyle = 2
    grpBreak.YAxisMin = 0#
    grpBreak.BottomTitle = Bottom_Title
    
    grpBreak.LeftTitle = "C/Ct"
    grpBreak.DrawMode = 2

End Sub

Sub Form_Activate ()
  If UnloadMe Then Unload Me
End Sub

Sub Form_Load ()
Dim J As Integer, I As Integer

    Top = Screen.Height / 2 - Height / 2
    Left = Screen.Width / 2 - Width / 2

    cboGraphType.AddItem "Symbols"
    cboGraphType.AddItem "Lines"

    If Obj_Function() Then
        Screen.MousePointer = 11
        For I = 1 To Results.NComponent
          cboCompo.AddItem Trim$(Results.Component(I).Name)
        Next I
        cboCompo.ListIndex = 0
        Call Draw_Curves(cboCompo.ListIndex + 1)
        cboGraphType.ListIndex = 0
        Screen.MousePointer = 0
        UnloadMe = False
    Else
        UnloadMe = True
    End If
End Sub

Sub Form_Resize ()
  If WindowState = 1 Then
    frmIonExchangeMain.WindowState = 1
    frmPlantdata.WindowState = 1
  End If

End Sub

Function Obj_Function () As Integer
Dim I, J As Integer
Dim ncomp As Long, ndata  As Long, np As Long, temp As String, Error_Code As Integer
ReDim Fmin(Results.NComponent) As Double
ReDim Tp(Results.NPoints) As Double, CP(Results.NComponent, Results.NPoints) As Double
ReDim Td(NData_Points) As Double, Cd(Results.NComponent, NData_Points) As Double, Cin(Results.NComponent, NData_Points) As Double

  ncomp = CLng(Results.NComponent)
  ndata = CLng(NData_Points)
  np = CLng(Results.NPoints)

  For I = 1 To Results.NPoints
    Tp(I) = Results.T(I)
    For J = 1 To Results.NComponent
      CP(J, I) = Results.CP(J, I)
    Next J
  Next I
  For I = 1 To NData_Points
    Td(I) = T_Data_Points(I)
    For J = 1 To Results.NComponent
      Cd(J, I) = C_Data_Points(J, I)
    Next J
  Next I

On Error GoTo Error_In_OBJFUN
  Call OBJFUN(ncomp, ndata, np, Tp(1), CP(1, 1), Td(1), Cd(1, 1), Cin(1, 1), Fmin(1))
  Obj_Function = True
  Exit Function

Error_In_OBJFUN:
  Error_Code = Err
  temp = "Error " & Format$(Error_Code, "0") & " : " & Error$(Error_Code)
  MsgBox "Fatal Error with OBJFUN.DLL. Calculations Stoppped." & Chr$(13) & temp, MB_ICONEXCLAMATION, Application_Name
  Obj_Function = False
  Resume Exit_Obj_Function
Exit_Obj_Function:
End Function

