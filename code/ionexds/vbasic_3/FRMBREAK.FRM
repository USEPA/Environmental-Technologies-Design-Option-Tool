VERSION 2.00
Begin Form frmbreak 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Results for the Plug-Flow Pore Diffusion Model"
   ClientHeight    =   6690
   ClientLeft      =   135
   ClientTop       =   675
   ClientWidth     =   9390
   Height          =   7095
   Left            =   75
   LinkTopic       =   "Form1"
   ScaleHeight     =   6690
   ScaleWidth      =   9390
   Top             =   330
   Width           =   9510
   Begin SSFrame Frame3D4 
      Caption         =   "Display Time In:"
      ForeColor       =   &H00000000&
      Height          =   675
      Left            =   7740
      ShadowColor     =   1  'Black
      TabIndex        =   41
      Top             =   2880
      Width           =   1575
      Begin ComboBox cboTime 
         Height          =   300
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   300
         Width           =   1455
      End
   End
   Begin SSFrame Frame3D3 
      Caption         =   "Grid Style:"
      ForeColor       =   &H00000000&
      Height          =   675
      Left            =   7740
      ShadowColor     =   1  'Black
      TabIndex        =   39
      Top             =   2040
      Width           =   1575
      Begin ComboBox cboGrid 
         Height          =   300
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   300
         Width           =   1455
      End
   End
   Begin SSCommand cmdFile 
      Caption         =   "Print to &File"
      Enabled         =   0   'False
      Height          =   435
      Left            =   7800
      TabIndex        =   25
      Top             =   6180
      Visible         =   0   'False
      Width           =   1455
   End
   Begin SSFrame Frame3D2 
      Caption         =   "C/Co as a function of:"
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   120
      TabIndex        =   31
      Top             =   6000
      Width           =   7575
      Begin SSOption optType 
         Caption         =   "&Time"
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   32
         Top             =   240
         Width           =   1215
      End
      Begin SSOption optType 
         Caption         =   "&BVT"
         Height          =   255
         Index           =   1
         Left            =   3840
         TabIndex        =   33
         Top             =   240
         Width           =   975
      End
      Begin SSOption optType 
         Caption         =   "Treatment C&apacity"
         Height          =   255
         Index           =   2
         Left            =   5280
         TabIndex        =   34
         Top             =   240
         Width           =   2055
      End
   End
   Begin SSFrame Frame3D1 
      Caption         =   "Results for:"
      ForeColor       =   &H00000000&
      Height          =   1935
      Left            =   120
      ShadowColor     =   1  'Black
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   7575
      Begin SSCommand cmdTreat 
         Caption         =   "Treatment Objective"
         Height          =   315
         Left            =   120
         TabIndex        =   26
         Top             =   1560
         Width           =   2415
      End
      Begin ComboBox cboCompo 
         Height          =   300
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   4815
      End
      Begin Label lblData 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "999.99"
         Height          =   255
         Index           =   15
         Left            =   6120
         TabIndex        =   38
         Top             =   1560
         Width           =   1215
      End
      Begin Label lblData 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "999.99"
         Height          =   255
         Index           =   14
         Left            =   4920
         TabIndex        =   37
         Top             =   1560
         Width           =   1215
      End
      Begin Label lblData 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "999.99"
         Height          =   255
         Index           =   13
         Left            =   3720
         TabIndex        =   36
         Top             =   1560
         Width           =   1215
      End
      Begin Label lblData 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "999.99"
         Height          =   255
         Index           =   12
         Left            =   2520
         TabIndex        =   35
         Top             =   1560
         Width           =   1215
      End
      Begin Label lblLegend 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "C (mg/L)"
         Height          =   255
         Index           =   7
         Left            =   6120
         TabIndex        =   30
         Top             =   600
         Width           =   1215
      End
      Begin Label lblData 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "999.99"
         Height          =   255
         Index           =   11
         Left            =   6120
         TabIndex        =   29
         Top             =   1320
         Width           =   1215
      End
      Begin Label lblData 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "999.99"
         Height          =   255
         Index           =   10
         Left            =   6120
         TabIndex        =   28
         Top             =   1080
         Width           =   1215
      End
      Begin Label lblData 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "999.99"
         Height          =   255
         Index           =   9
         Left            =   6120
         TabIndex        =   27
         Top             =   840
         Width           =   1215
      End
      Begin Label lblData 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "999.99"
         Height          =   255
         Index           =   8
         Left            =   4920
         TabIndex        =   15
         Top             =   1320
         Width           =   1215
      End
      Begin Label lblData 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "999.99"
         Height          =   255
         Index           =   7
         Left            =   4920
         TabIndex        =   16
         Top             =   1080
         Width           =   1215
      End
      Begin Label lblData 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "999.99"
         Height          =   255
         Index           =   6
         Left            =   4920
         TabIndex        =   17
         Top             =   840
         Width           =   1215
      End
      Begin Label lblData 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "999.99"
         Height          =   255
         Index           =   5
         Left            =   3720
         TabIndex        =   19
         Top             =   1320
         Width           =   1215
      End
      Begin Label lblData 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "999.99"
         Height          =   255
         Index           =   4
         Left            =   3720
         TabIndex        =   20
         Top             =   1080
         Width           =   1215
      End
      Begin Label lblData 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "999.99"
         Height          =   255
         Index           =   3
         Left            =   3720
         TabIndex        =   21
         Top             =   840
         Width           =   1215
      End
      Begin Label lblData 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "999.99"
         Height          =   255
         Index           =   2
         Left            =   2520
         TabIndex        =   22
         Top             =   1320
         Width           =   1215
      End
      Begin Label lblData 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "999.99"
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   23
         Top             =   1080
         Width           =   1215
      End
      Begin Label lblData 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "999.99"
         Height          =   255
         Index           =   0
         Left            =   2520
         TabIndex        =   14
         Top             =   840
         Width           =   1215
      End
      Begin Label lblLegend 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "95% of influent conc."
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   2415
      End
      Begin Label lblLegend 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "50% of influent conc."
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   2415
      End
      Begin Label lblLegend 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "5% of influent conc."
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   2415
      End
      Begin Label lblLegend 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tr. Capacity"
         Height          =   255
         Index           =   3
         Left            =   4920
         TabIndex        =   10
         Top             =   600
         Width           =   1215
      End
      Begin Label lblLegend 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "BVT"
         Height          =   255
         Index           =   2
         Left            =   3720
         TabIndex        =   9
         Top             =   600
         Width           =   1215
      End
      Begin Label lblLegend 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Time (days)"
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   8
         Top             =   600
         Width           =   1215
      End
      Begin Label lblLegend 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   2415
      End
   End
   Begin SSCommand cmdExcel 
      Caption         =   "&Excel..."
      Height          =   435
      Left            =   7800
      TabIndex        =   3
      Top             =   4020
      Visible         =   0   'False
      Width           =   1455
   End
   Begin SSCommand cmdSave 
      Caption         =   "Save &Curves"
      Height          =   435
      Left            =   7800
      TabIndex        =   7
      Top             =   4560
      Width           =   1455
   End
   Begin CommonDialog CMDialog1 
      Left            =   8400
      Top             =   720
   End
   Begin SSCommand cmdSelect 
      Caption         =   "&Select Printer"
      Height          =   435
      Left            =   7800
      TabIndex        =   18
      Top             =   5100
      Width           =   1455
   End
   Begin CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   435
      Left            =   7800
      TabIndex        =   24
      Top             =   5640
      Width           =   1455
   End
   Begin CommandButton cmdOK 
      Caption         =   "E&xit"
      Height          =   435
      Left            =   7800
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin GRAPH grpBreak 
      Height          =   3855
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2040
      Width           =   7575
   End
End
Option Explicit
Option Base 1

'Dim Treatment_Objective(NumSelectedComponents_PFPDM) As Throughput
'Dim Flag_TO(NumSelectedComponents_PFPDM) As Integer
Dim Filename_Input As String

Const TAB_1 = 10
Const TAB_2 = 20
Const TAB_3 = 30
Const TAB_SAVE_INTERVAL = 10

Sub cboCompo_Click ()

  Exit Sub

End Sub

Sub cboGrid_Click ()
   grpBreak.GridStyle = cboGrid.ListIndex
   grpBreak.DrawMode = 2
End Sub

Sub cboTime_Click ()
    Dim I As Integer
    Dim FoundTrue As Integer

    FoundTrue = False
    For I = 0 To 2
        If optType(I).Value = True Then
           FoundTrue = True
           optType(I).Value = False
           optType(I).Value = True
           Exit For
        End If
    Next I

    If Not FoundTrue Then
       optType(0).Value = True
    End If

    TimeUnitsOnGraphs = cboTime.ListIndex

End Sub

Sub cmdExcel_Click ()
  frmExcel.Show 1
End Sub

Sub cmdExcel_KeyPress (KeyAscii As Integer)
 Call Key_Pressed_On_Control(KeyAscii)

End Sub

Sub cmdFile_Click ()
End Sub

Sub cmdFile_KeyPress (KeyAscii As Integer)
  Call Key_Pressed_On_Control(KeyAscii)
End Sub

Sub cmdOK_Click ()
    Unload Me
End Sub

Sub cmdOK_KeyPress (KeyAscii As Integer)
 Call Key_Pressed_On_Control(KeyAscii)

End Sub

Sub cmdPrint_Click ()
 
Dim Error_Code As Integer, temp As String
Dim I As Integer, H  As Single, W As Single
Dim Eq1 As String, J As Integer, MTZ As String, f As Double

On Error GoTo Print_Error

'---- Print the graph ------------------------
    For I = 1 To Number_Component
     grpBreak.ThisPoint = I
     grpBreak.PatternData = I - 1
    Next I

    H = grpBreak.Height
    W = grpBreak.Width

    grpBreak.Visible = False 'Hide it before printing

    If Printer.Width < Printer.Height Then
      grpBreak.Height = CSng(Printer.Height / 2#)
      grpBreak.Width = Printer.Width
    Else
      grpBreak.Height = Printer.Height
      grpBreak.Width = Printer.Width
    End If

    grpBreak.PrintStyle = 2
    grpBreak.DrawMode = 5

    grpBreak.Height = H
    grpBreak.Width = W

    grpBreak.Visible = True

    grpBreak.PrintStyle = 2
    grpBreak.DrawMode = 2

    Call PrintIonExchange

    Exit Sub

Print_Error:
  Error_Code = Err
  temp = "Error " & Format$(Error_Code, "0") & ": " & Error$(Error_Code)
  MsgBox "An error uccured while printing." & Chr$(13) & temp, MB_ICONEXCLAMATION, Application_Name
  Resume Exit_Print
Exit_Print:

End Sub

Sub cmdPrint_KeyPress (KeyAscii As Integer)
 Call Key_Pressed_On_Control(KeyAscii)

End Sub

Sub cmdSave_Click ()
Dim f As Integer, I As Integer, J As Integer, temp As String
Dim Filename_PFS As String
Dim TimeToDisplay As Double
Dim ValueToDisplay As Double
 
On Error GoTo Save_Results_PF_Error

   CMDialog1.Filename = ""
   CMDialog1.Filter = "All Files (*.*)|*.*|Text Files (*.txt)|*.txt|Data Files (*.iex)|*.iex"
   CMDialog1.FilterIndex = 2
   CMDialog1.DialogTitle = "Save curves from PFPDM"
'------Begin Modification Hokanson: 12-Aug2000
   CMDialog1.CancelError = True
'------End Modification Hokanson: 11-Aug2000
   CMDialog1.Action = 2

   f = FileNameIsValid(Filename_PFS, CMDialog1)
   If Not (f) Then Exit Sub

   'Save, T, BVF, Usage rate, C/C0
    f = FreeFile
    Open Filename_PFS For Output As f
    Print #f, "Results file:  PFPDM for Windows - Version " & Format$(NVersion, "0.00")
    Print #f,
    Print #f,

     If cboTime.ListIndex = 0 Then       'min
        Print #f, "Time-min"; Tab(TAB_1); "BVF"; Tab(TAB_2); "UR-m3/kg";
     ElseIf cboTime.ListIndex = 1 Then   's
        Print #f, "Time-s"; Tab(TAB_1); "BVF"; Tab(TAB_2); "UR-m3/kg";
     ElseIf cboTime.ListIndex = 2 Then   'hr
        Print #f, "Time-hr"; Tab(TAB_1); "BVF"; Tab(TAB_2); "UR-m3/kg";
     ElseIf cboTime.ListIndex = 3 Then   'd
        Print #f, "Time-d"; Tab(TAB_1); "BVF"; Tab(TAB_2); "UR-m3/kg";
     End If

    For I = 1 To Results.NComponent
     Print #f, Tab(TAB_3 + (I - 1) * TAB_SAVE_INTERVAL); Trim$(Results.Component(I).Name);
    Next I
    
    Print #f,
    Print #f,

    For I = 1 To Results.NPoints
     If cboTime.ListIndex = 0 Then       'min
        TimeToDisplay = Results.T(I)
     ElseIf cboTime.ListIndex = 1 Then   's
        TimeToDisplay = Results.T(I) * 60#
     ElseIf cboTime.ListIndex = 2 Then   'hr
        TimeToDisplay = Results.T(I) / 60#
     ElseIf cboTime.ListIndex = 3 Then   'd
        TimeToDisplay = Results.T(I) / 60# / 24#
     End If

     Print #f, Format$(TimeToDisplay, GetTheFormat(TimeToDisplay));
     ValueToDisplay = Results.T(I) * 60 * Results.Bed.Flowrate.Value / Results.Bed.Length / Pi / (Results.Bed.Diameter / 2) ^ 2
     Print #f, Tab(TAB_1); Format$(ValueToDisplay, GetTheFormat(ValueToDisplay));
     ValueToDisplay = Results.T(I) * 60 * Results.Bed.Flowrate.Value / Results.Bed.Weight
     Print #f, Tab(TAB_2); Format$(ValueToDisplay, GetTheFormat(ValueToDisplay));

     For J = 1 To Results.NComponent
       ValueToDisplay = Results.CP(J, I)
       Print #f, Tab(TAB_3 + (J - 1) * TAB_SAVE_INTERVAL); Format$(ValueToDisplay, GetTheFormat(ValueToDisplay));
     Next J
     Print #f,
    
    Next I
    Close f
    CMDialog1.Filename = ""
 Exit Sub

Save_Results_PF_Error:
  If Err = 32755 Then
  Else
    MsgBox "Error occurred trying to save results, please retry.", MB_ICONEXCLAMATION, Application_Name
  End If
  Resume Exit_Save_Results_PF
Exit_Save_Results_PF:

End Sub

Sub cmdSave_KeyPress (KeyAscii As Integer)
 Call Key_Pressed_On_Control(KeyAscii)

End Sub

Sub cmdSelect_Click ()
Dim Error_Code As Integer, temp As String
On Error GoTo Select_Print_Error
  CMDialog1.Flags = PD_PRINTSETUP
'------Begin Modification Hokanson: 12-Aug2000
  CMDialog1.CancelError = True
'------End Modification Hokanson: 11-Aug2000
  CMDialog1.Action = 5
  Exit Sub
Select_Print_Error:
  If Err = 32755 Then
  Else
     Error_Code = Err
     temp = "Error " & Format$(Error_Code, "0") & ": " & Error$(Error_Code)
     MsgBox "An error occured while selecting the printer." & Chr$(13) & temp, MB_ICONEXCLAMATION, Application_Name
  End If
  Resume Exit_Select_Print
Exit_Select_Print:
End Sub

Sub cmdSelect_KeyPress (KeyAscii As Integer)
 Call Key_Pressed_On_Control(KeyAscii)
End Sub

Sub Draw_PFPSDM ()
Dim I As Integer, J As Integer
Dim Data_Max As Double, factor As Double, Bottom_Title As String
ReDim X_Values(Number_Points_Max) As Double

'Copy the results
  If optType(0) Then  'Time
     If cboTime.ListIndex = 0 Then       'min
        factor = 1#
        Bottom_Title = "Time (min)"
     ElseIf cboTime.ListIndex = 1 Then   's
        factor = 1# * 60#
        Bottom_Title = "Time (s)"
     ElseIf cboTime.ListIndex = 2 Then   'hr
        factor = 1# / 60#
        Bottom_Title = "Time (hr)"
     ElseIf cboTime.ListIndex = 3 Then   'd
        factor = 1# / 60# / 24#
        Bottom_Title = "Time (d)"
     End If
    
  Else
    If optType(1) Then   'BVF         mn * (mn/s) * (m3/s) / m / (m2) -> dimensionless
      factor = 60# * Results.Bed.Flowrate.Value / Results.Bed.Length / Pi / (Results.Bed.Diameter / 2#) ^ 2
      Bottom_Title = "Bed Volumes Treated"
    Else   'Treatment Capacity
      factor = 60# * Results.Bed.Flowrate.Value / Results.Bed.Weight   'mn * (s/mn) * (m3/s) / (kg) -> m3/kg
      Bottom_Title = "m3 treated per kg of resin"
    End If
  End If
  'Results.T(I,1) time is in mn
  'Results.T(I,2) is BVF
  For I = 1 To Number_Points_Max
    X_Values(I) = Results.T(I) * factor
  Next I

     
    'Define Graph
    grpBreak.NumSets = Results.NComponent
    grpBreak.GraphType = 6
    grpBreak.GraphStyle = 4

    For J = 1 To grpBreak.NumSets
     grpBreak.ThisSet = J
     grpBreak.NumPoints = Results.NPoints
    Next J

    grpBreak.AutoInc = 0
    
    For J = 1 To grpBreak.NumSets
       grpBreak.ThisSet = J
       For I = 1 To grpBreak.NumPoints
         grpBreak.ThisPoint = I
         If Results.CP(J, I) < 0 Then
           grpBreak.GraphData = 0#
         Else
           grpBreak.GraphData = Results.CP(J, I)
         End If
         grpBreak.ThisPoint = I
         grpBreak.LabelText = ""
         grpBreak.ThisPoint = I
         grpBreak.XPosData = X_Values(I)
       Next I
       grpBreak.ThisPoint = J
       grpBreak.LegendText = Trim$(Results.Component(J).Name)
       grpBreak.ThisPoint = J
       grpBreak.PatternData = J - 1
    Next J

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
    
'    grpBreak.LeftTitle = "C/Co"
    grpBreak.LeftTitle = "C/Ct"
    grpBreak.DrawMode = 2

End Sub

Sub Form_Load ()
Dim J As Integer, I As Integer
   'Set Window
    Top = Screen.Height / 2 - Height / 2
    Left = Screen.Width / 2 - Width / 2
'      Me.HelpContextID = Hlp_Results_for
    Screen.MousePointer = 11
    cboGrid.AddItem "None"
    cboGrid.AddItem "Horizontal"
    cboGrid.AddItem "Vertical"
    cboGrid.AddItem "Both"
    cboGrid.ListIndex = 0

    For I = 1 To Results.NComponent
       cboCompo.AddItem Trim$(Results.Component(I).Name)
'       Treatment_Objective(I) = Results.ThroughPut_05(I)
'       If Treatment_Objective(I).c <> -1 Then
'         Flag_TO(I) = True
'       Else
'         Flag_TO(I) = False
'       End If
    Next I

    Screen.MousePointer = 0
    cboCompo.ListIndex = 0

    cboTime.AddItem "min"
    cboTime.AddItem "s"
    cboTime.AddItem "hr"
    cboTime.AddItem "d"
    cboTime.ListIndex = TimeUnitsOnGraphs


End Sub

Sub Key_Pressed_On_Control (Ascii_Code As Integer)
  Select Case Ascii_Code
    Case 64, 97'A,a
      optType(2) = True
    Case 66, 98'B,a
      optType(1) = True
    Case 67, 99 'C,c
      cmdSave_Click
    Case 69, 101'E,e
      cmdExcel_Click
    Case 70, 102 'F,f
      cmdFile_Click
    Case 80, 112 'P,p
      cmdPrint_Click
    Case 83, 115'S,s
     cmdSelect_Click
    Case 84, 116'T,t
     optType(0) = True
    Case 88, 120'X,x
     cmdOK_Click
  End Select
End Sub

Sub optType_Click (Index As Integer, Value As Integer)
  Call Draw_PFPSDM
End Sub

Sub optType_KeyPress (Index As Integer, KeyAscii As Integer)
  Call Key_Pressed_On_Control(KeyAscii)
End Sub

