VERSION 5.00
Begin VB.Form frmgraphSet 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Graph Setup"
   ClientHeight    =   6315
   ClientLeft      =   1065
   ClientTop       =   1920
   ClientWidth     =   6975
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmgraph.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6315
   ScaleWidth      =   6975
   Begin VB.CommandButton rangecmd 
      Caption         =   "View Range"
      Height          =   315
      Left            =   5340
      TabIndex        =   39
      Top             =   4920
      Width           =   1395
   End
   Begin VB.ComboBox CMBPropUnits 
      Height          =   315
      Left            =   120
      TabIndex        =   37
      Text            =   "CMBPropUnits"
      Top             =   3240
      Width           =   4455
   End
   Begin VB.ListBox LSTUserList 
      Height          =   1425
      Left            =   120
      MultiSelect     =   1  'Simple
      TabIndex        =   27
      Top             =   300
      Width           =   4455
   End
   Begin VB.ComboBox CMBPropMethod 
      Height          =   315
      Left            =   120
      TabIndex        =   25
      Text            =   "CMBPropMethod"
      Top             =   2640
      Width           =   4455
   End
   Begin VB.ComboBox CMBPropFunction 
      Height          =   315
      Left            =   120
      TabIndex        =   24
      Text            =   "CMBProp"
      Top             =   2040
      Width           =   4455
   End
   Begin VB.Frame FRMGraphOptions 
      Caption         =   "Graph Text"
      Height          =   2175
      Left            =   120
      TabIndex        =   3
      Top             =   3600
      Width           =   6735
      Begin VB.TextBox TXTMinT 
         Height          =   285
         Left            =   1320
         TabIndex        =   29
         Text            =   "TXTMinT"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox TXTMaxT 
         Height          =   285
         Left            =   3120
         TabIndex        =   28
         Text            =   "TXTMaxT"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox TXTNumPts 
         Height          =   285
         Left            =   5400
         TabIndex        =   11
         Text            =   "TXTNumPts"
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox TXTXAxis 
         Height          =   285
         Left            =   1320
         TabIndex        =   9
         Top             =   960
         Width           =   5295
      End
      Begin VB.TextBox TXTYAxis 
         Height          =   285
         Left            =   1320
         TabIndex        =   8
         Top             =   600
         Width           =   5295
      End
      Begin VB.TextBox TXTGraphTitle 
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Top             =   240
         Width           =   5295
      End
      Begin VB.Label LBLTFTUnit 
         Caption         =   "LBLTFTUnit"
         Height          =   255
         Index           =   1
         Left            =   4440
         TabIndex        =   33
         Top             =   1350
         Width           =   150
      End
      Begin VB.Label LBLTFTUnit 
         Caption         =   "LBLTFTUnit"
         Height          =   255
         Index           =   0
         Left            =   2640
         TabIndex        =   32
         Top             =   1350
         Width           =   150
      End
      Begin VB.Label LBLTo 
         Caption         =   "to"
         Height          =   255
         Left            =   2830
         TabIndex        =   31
         Top             =   1350
         Width           =   255
      End
      Begin VB.Label LBLGraphFrom 
         Alignment       =   1  'Right Justify
         Caption         =   "Graph from"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1350
         Width           =   1095
      End
      Begin VB.Label LBLNumPts 
         Caption         =   "Number of points to graph within specified temperature range"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1800
         Width           =   5295
      End
      Begin VB.Label LBLXAxis 
         Alignment       =   1  'Right Justify
         Caption         =   "X-Axis Label"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   990
         Width           =   1095
      End
      Begin VB.Label LBLYAxis 
         Alignment       =   1  'Right Justify
         Caption         =   "Y-Axis Label"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   630
         Width           =   1095
      End
      Begin VB.Label LBLGraphTitle 
         Alignment       =   1  'Right Justify
         Caption         =   "Graph Title"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   270
         Width           =   1095
      End
   End
   Begin VB.CommandButton CMDCan 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   5820
      Width           =   1455
   End
   Begin VB.CommandButton CMDGraph 
      Caption         =   "Graph"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   5820
      Width           =   1455
   End
   Begin VB.PictureBox FRMLineSym 
      Height          =   1095
      Left            =   4680
      ScaleHeight     =   1035
      ScaleWidth      =   2115
      TabIndex        =   14
      Top             =   240
      Width           =   2175
      Begin VB.OptionButton OPBLineSym 
         Caption         =   "Lines and Symbols"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   1935
      End
      Begin VB.OptionButton OPBLineSym 
         Caption         =   "Symbols"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   1935
      End
      Begin VB.OptionButton OPBLineSym 
         Caption         =   "Lines "
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.PictureBox FRMYAxis 
      Height          =   1095
      Left            =   4680
      ScaleHeight     =   1035
      ScaleWidth      =   2115
      TabIndex        =   18
      Top             =   1320
      Width           =   2175
      Begin VB.OptionButton OPBYAxis 
         Caption         =   "log (1/Y)"
         Height          =   255
         Index           =   3
         Left            =   960
         TabIndex        =   36
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton OPBYAxis 
         Caption         =   "log Y"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   35
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton OPBYAxis 
         Caption         =   "1/Y"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton OPBYAxis 
         Caption         =   "Y"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.PictureBox FRMXAxis 
      Height          =   1095
      Left            =   4680
      ScaleHeight     =   1035
      ScaleWidth      =   2115
      TabIndex        =   21
      Top             =   2400
      Width           =   2175
      Begin VB.OptionButton OPBXAxis 
         Caption         =   "log (1/T)"
         Height          =   255
         Index           =   3
         Left            =   960
         TabIndex        =   34
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton OPBXAxis 
         Caption         =   "log T"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   26
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton OPBXAxis 
         Caption         =   "1/T"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   23
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton OPBXAxis 
         Caption         =   "T"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Label LBLPropUnits 
      Caption         =   "Property Units"
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   3000
      Width           =   4455
   End
   Begin VB.Label LBLPropMethod 
      Caption         =   "Property Method"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2400
      Width           =   4455
   End
   Begin VB.Label LBLUserList 
      Caption         =   "User List"
      Height          =   255
      Left            =   180
      TabIndex        =   10
      Top             =   60
      Width           =   4455
   End
   Begin VB.Label LBLFuncToGraph 
      Caption         =   "Property"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   4455
   End
End
Attribute VB_Name = "frmgraphSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit









Private Sub CMBPropFunction_Click()
    
    Dim i As Integer
    
    Select Case CMBPropFunction.Text
        Case "Liquid Density as f(T)"
            CurProp = LD
        Case "Vapor Pressure as f(T)"
            CurProp = VP
        Case "Liquid Heat Capacity as f(T)"
            CurProp = LHC
        Case "Vapor Heat Capacity as f(T)"
            CurProp = VHC
        Case "Heat of Vaporization as f(T)"
            CurProp = Hvap
        Case "Surface Tension as f(T)"
            CurProp = ST
        Case "Vapor Viscosity as f(T)"
            CurProp = VV
        Case "Liquid Viscosity as f(T)"
            CurProp = LV
        Case "Liquid Thermal Conductivity as f(T)"
            CurProp = LTC
        Case "Vapor Thermal Conductivity as f(T)"
            CurProp = VTC
    End Select
    
    'Reset graph properties
    TXTGraphTitle.Text = Trim(CMBPropFunction.Text)
    TXTYAxis.Text = Trim(Mid(CMBPropFunction.Text, 1, Len(CMBPropFunction.Text) - 8)) & " (" & Trim(CMBPropUnits.Text) & ")"
        
    'Fill in units to choose from
    CMBPropUnits.Clear
    CMBPropUnits.AddItem Get_DefaultUnit(CurProp)
    i = 1
    Do While Unit1(i) <> "End"
        If Trim(Unit1(i)) = Trim(Get_DefaultUnit(CurProp)) Then
            CMBPropUnits.AddItem Unit2(i)
        End If
        i = i + 1
    Loop
    
    ' now update the y axis label
    CMBPropUnits.ListIndex = 0
     ' REVISION  DMW : 6/6/97  took out the log part of the label since Y-axis not expressed in log
    If OPBYAxis(0).value = True Then
        TXTYAxis.Text = Trim(Mid(CMBPropFunction.Text, 1, Len(CMBPropFunction.Text) - 8)) & " (" & Trim(CMBPropUnits.Text) & ")"
    ElseIf OPBYAxis(1).value = True Then
        TXTYAxis.Text = "1/" & Trim(Mid(CMBPropFunction.Text, 1, Len(CMBPropFunction.Text) - 8)) & " (" & Trim(CMBPropUnits.Text) & ")"
    ElseIf OPBYAxis(2).value = True Then
        TXTYAxis.Text = Trim(Mid(CMBPropFunction.Text, 1, Len(CMBPropFunction.Text) - 8)) & " (" & Trim(CMBPropUnits.Text) & ")"
       ' TXTYAxis.Text = "log " & Trim(Mid(CMBPropFunction.Text, 1, Len(CMBPropFunction.Text) - 8)) & " (" & Trim(CMBPropUnits.Text) & ")"
    ElseIf OPBYAxis(3).value = True Then
        TXTYAxis.Text = "1/" & Trim(Mid(CMBPropFunction.Text, 1, Len(CMBPropFunction.Text) - 8)) & " (" & Trim(CMBPropUnits.Text) & ")"
        'TXTYAxis.Text = "log (1/" & Trim(Mid(CMBPropFunction.Text, 1, Len(CMBPropFunction.Text) - 8)) & " (" & Trim(CMBPropUnits.Text) & "))"
    End If
    
End Sub


Private Sub CMBPropUnits_Click()

    TXTYAxis.Text = Trim(Mid(CMBPropFunction.Text, 1, Len(CMBPropFunction.Text) - 8)) & " (" & Trim(CMBPropUnits.Text) & ")"

End Sub


Private Sub CMDCan_Click()
             
    Call DisplayProps
    
    Unload FRMGraphSet
        
End Sub

Private Sub CMDGraph_Click()

    Dim i As Integer
    Dim Test As Boolean
    
    Test = False
    For i = 0 To FRMGraphSet!LSTUserList.ListCount - 1
        If FRMGraphSet!LSTUserList.Selected(i) = True Then
            Test = True
        End If
    Next i
    
    
    If FRMGraphSet!TXTMinT.Text = "" Or FRMGraphSet!TXTMaxT.Text = "" Then
        MsgBox "Please specify a temperature range", 48, "Temperature Range Not Set"
        Exit Sub
    End If
    
    If Val(FRMGraphSet!TXTNumPts.Text) > 100 Then
        MsgBox "Maximum number of points is 100", 48, "Too Many Points"
        Exit Sub
    End If
    
    If Test = True Then
        Call Graph
        FRMGraph.Show 1
    Else
        MsgBox "Please select a chemical to graph", 48, "No Chemicals Selected"
    End If
    
End Sub

Private Sub Form_Load()
            
    Dim i As Integer
    Dim TempCAS As Long
    
    'Fill in chemicals to select from
    
    'TempCAS = FRMMain!Data2.Recordset("CAS")
    FRMMain!Data2.Recordset.MoveFirst
    Do While Not FRMMain!Data2.Recordset.EOF
        LSTUserList.AddItem FRMMain!Data2.Recordset("Name")
        FRMMain!Data2.Recordset.MoveNext
    Loop
    FRMMain!Data2.Recordset.FindFirst "CAS =" & TempCAS
    
    'Set default number of points to graph
    TXTNumPts.Text = "5"
    
    'Fill the combo box with properties to choose from
    CMBPropFunction.Clear
    CMBPropFunction.AddItem "Liquid Density as f(T)"
    CMBPropFunction.AddItem "Vapor Pressure as f(T)"
    CMBPropFunction.AddItem "Liquid Heat Capacity as f(T)"
    CMBPropFunction.AddItem "Vapor Heat Capacity as f(T)"
    CMBPropFunction.AddItem "Heat of Vaporization as f(T)"
    CMBPropFunction.AddItem "Surface Tension as f(T)"
    CMBPropFunction.AddItem "Vapor Viscosity as f(T)"
    CMBPropFunction.AddItem "Liquid Viscosity as f(T)"
    CMBPropFunction.AddItem "Liquid Thermal Conductivity as f(T)"
    CMBPropFunction.AddItem "Vapor Thermal Conductivity as f(T)"
    
    'Fill in methods to choose from
    CMBPropMethod.Clear
    If Path801 <> NULLPATH And DIPPR801 = True Then
        CMBPropMethod.AddItem "801 Database"
    End If
    If Path911 <> NULLPATH And DIPPR911 = True Then
        CMBPropMethod.AddItem "911 Database"
    End If
    CMBPropMethod.AddItem "User Input"
    
    'Set initial selections
    CMBPropFunction.ListIndex = 0
    CMBPropMethod.ListIndex = 0
  
    'Fill in units to choose from
    CMBPropUnits.Clear
    CMBPropUnits.AddItem Get_DefaultUnit(CurProp)
    i = 1
    Do While Unit1(i) <> "End"
        If Trim(Unit1(i)) = Trim(Get_DefaultUnit(CurProp)) Then
            CMBPropUnits.AddItem Unit2(i)
        End If
        i = i + 1
    Loop
    CMBPropUnits.ListIndex = 0
        
    'Set temperature range and units
    TXTXAxis.Text = "T (K)"
    TXTMinT.Text = "100"
    TXTMaxT.Text = "500"
    LBLTFTUnit(0).caption = "K"
    LBLTFTUnit(1).caption = "K"
    
    'Set X and Y Axis options
    OPBLineSym(0).value = True
    OPBYAxis(0).value = True
    OPBXAxis(0).value = True
            
    'Center form on the screen
    CenterForm Me
    
End Sub

Private Sub LBLTFTUnit_Click(Index As Integer)

    Dim i As Integer
       
    GraphConvert = True
    
    FRMUnits!CMBUnits.Clear

'msh FRMUnits!CMBUnits.AddItem ConvertToDefault(-1)
    FRMUnits!CMBUnits.AddItem Get_DefaultUnit(-1)

    i = 1
    Do While Unit1(i) <> "End"
'msh    If Trim(Unit1(i)) = Trim(ConvertToDefault(-1)) Then
        If Trim(Unit1(i)) = Trim(Get_DefaultUnit(-1)) Then
            FRMUnits!CMBUnits.AddItem Unit2(i)
'msh    ElseIf Trim(Unit2(i)) = Trim(ConvertToDefault(CurProp)) Then
        ElseIf Trim(Unit2(i)) = Trim(Get_DefaultUnit(CurProp)) Then
            FRMUnits!CMBUnits.AddItem Unit1(i)
        End If
        i = i + 1
    Loop
        
    FRMUnits.Show 1
    
    ' now update the x axis label
    If OPBXAxis(0).value = True Then
        TXTXAxis.Text = "T (" & Trim(LBLTFTUnit(0).caption) & ")"
    ElseIf OPBXAxis(1).value = True Then
        TXTXAxis.Text = "1/T (" & Trim(LBLTFTUnit(0).caption) & ")"
    ElseIf OPBXAxis(2).value = True Then
        TXTXAxis.Text = "log T (" & Trim(LBLTFTUnit(0).caption) & ")"
    ElseIf OPBXAxis(3).value = True Then
        TXTXAxis.Text = "log (1/T (" & Trim(LBLTFTUnit(0).caption) & "))"
    End If
End Sub


Private Sub OPBXAxis_Click(Index As Integer)

    If OPBXAxis(0).value = True Then
        TXTXAxis.Text = "T (" & Trim(LBLTFTUnit(0).caption) & ")"
    ElseIf OPBXAxis(1).value = True Then
        TXTXAxis.Text = "1/T (" & Trim(LBLTFTUnit(0).caption) & ")"
    ElseIf OPBXAxis(2).value = True Then
        TXTXAxis.Text = "log T (" & Trim(LBLTFTUnit(0).caption) & ")"
    ElseIf OPBXAxis(3).value = True Then
        TXTXAxis.Text = "log (1/T (" & Trim(LBLTFTUnit(0).caption) & "))"
    End If

End Sub


Private Sub OPBYAxis_Click(Index As Integer)

    ' REVISION  DMW : 6/6/97  took out the log part of the label since Y-axis not expressed in log
    If OPBYAxis(0).value = True Then
        TXTYAxis.Text = Trim(Mid(CMBPropFunction.Text, 1, Len(CMBPropFunction.Text) - 8)) & " (" & Trim(CMBPropUnits.Text) & ")"
    ElseIf OPBYAxis(1).value = True Then
        TXTYAxis.Text = "1/" & Trim(Mid(CMBPropFunction.Text, 1, Len(CMBPropFunction.Text) - 8)) & " (" & Trim(CMBPropUnits.Text) & ")"
    ElseIf OPBYAxis(2).value = True Then
        TXTYAxis.Text = Trim(Mid(CMBPropFunction.Text, 1, Len(CMBPropFunction.Text) - 8)) & " (" & Trim(CMBPropUnits.Text) & ")"
       ' TXTYAxis.Text = "log " & Trim(Mid(CMBPropFunction.Text, 1, Len(CMBPropFunction.Text) - 8)) & " (" & Trim(CMBPropUnits.Text) & ")"
    ElseIf OPBYAxis(3).value = True Then
        TXTYAxis.Text = "1/" & Trim(Mid(CMBPropFunction.Text, 1, Len(CMBPropFunction.Text) - 8)) & " (" & Trim(CMBPropUnits.Text) & ")"
        'TXTYAxis.Text = "log (1/" & Trim(Mid(CMBPropFunction.Text, 1, Len(CMBPropFunction.Text) - 8)) & " (" & Trim(CMBPropUnits.Text) & "))"
    End If

End Sub


Private Sub rangecmd_Click()
' this opens up a form which tells the user
' the legal range for the chemicals in the list
' for the selected property

FRMInfo.rangegrd.Rows = FRMGraphSet!LSTUserList.ListCount + 1
FRMInfo.rangegrd.Cols = 5
Call fill_range_form(CurProp)    ' a function in modgraph.bas
' if this form is used for other purposes, make those frames invisible
' (some rearranging will be required if it's used generically)
' then make the range form visible
FRMInfo.rangepnl.Visible = True
FRMInfo.Show 1
End Sub

Private Sub rangelbl_DblClick()
' this opens up a form which tells the user
' the legal range for the chemicals in the list
' for the selected property

FRMInfo.rangegrd.Rows = FRMGraphSet!LSTUserList.ListCount + 1
FRMInfo.rangegrd.Cols = 5
Call fill_range_form(CurProp)    ' a function in modgraph.bas
' if this form is used for other purposes, make those frames invisible
' (some rearranging will be required if it's used generically)
' then make the range form visible
FRMInfo.rangepnl.Visible = True
FRMInfo.Show 1
End Sub

