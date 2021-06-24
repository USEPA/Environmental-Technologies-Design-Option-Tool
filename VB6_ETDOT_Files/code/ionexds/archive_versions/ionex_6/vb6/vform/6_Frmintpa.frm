VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{B16553C3-06DB-101B-85B2-0000C009BE81}#1.0#0"; "Spin32.ocx"
Begin VB.Form frmOptionsInputParameters 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Algorithm Parameters"
   ClientHeight    =   5340
   ClientLeft      =   885
   ClientTop       =   1545
   ClientWidth     =   5895
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
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5340
   ScaleWidth      =   5895
   Begin Threed.SSCommand cmdCancel 
      Height          =   375
      Left            =   4440
      TabIndex        =   26
      Top             =   4800
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "&Cancel"
   End
   Begin Threed.SSFrame Frame3D1 
      Height          =   2415
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Width           =   5535
      _Version        =   65536
      _ExtentX        =   9763
      _ExtentY        =   4260
      _StockProps     =   14
      Caption         =   "Other Parameters:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtTime 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   3000
         TabIndex        =   20
         Text            =   "txtTime"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtTime 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   3000
         TabIndex        =   19
         Text            =   "txtTime"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtTime 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   3000
         TabIndex        =   18
         Text            =   "txtTime"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.ComboBox cboTimeParametersUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   0
         Left            =   4320
         Style           =   2  'Dropdown List
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   480
         Width           =   1095
      End
      Begin VB.ComboBox cboTimeParametersUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   1
         Left            =   4320
         Style           =   2  'Dropdown List
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   840
         Width           =   1095
      End
      Begin VB.ComboBox cboTimeParametersUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   2
         Left            =   4320
         Style           =   2  'Dropdown List
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtTime 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   3000
         TabIndex        =   14
         Text            =   "txtTime"
         Top             =   1560
         Width           =   1215
      End
      Begin VB.ComboBox cboTimeParametersUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   4
         Left            =   4320
         Style           =   2  'Dropdown List
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox txtTime 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   3000
         TabIndex        =   12
         Text            =   "txtTime"
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label lblTime 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Run Time:"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   25
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label lblTime 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "First point displayed:"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   24
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label lblTime 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Time Step:"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   960
         TabIndex        =   23
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label lblTime 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Integrator Error Criteria, EPS"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   22
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label lblTime 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Initial Integrator Time Step, DH0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   21
         Top             =   1920
         Width           =   2775
      End
   End
   Begin Threed.SSFrame fraPoint 
      Height          =   1215
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   5295
      _Version        =   65536
      _ExtentX        =   9340
      _ExtentY        =   2143
      _StockProps     =   14
      Caption         =   "Number of collocation points:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Spin.SpinButton spnPoint 
         Height          =   255
         Index           =   1
         Left            =   3600
         TabIndex        =   10
         Top             =   720
         Width           =   135
         _Version        =   65536
         _ExtentX        =   238
         _ExtentY        =   450
         _StockProps     =   73
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
      End
      Begin Spin.SpinButton spnPoint 
         Height          =   255
         Index           =   0
         Left            =   3600
         TabIndex        =   11
         Top             =   360
         Width           =   135
         _Version        =   65536
         _ExtentX        =   238
         _ExtentY        =   450
         _StockProps     =   73
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
      End
      Begin VB.Label lblText 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Axial direction"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   9
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblText 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Radial Direction"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   8
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label lblNPoint 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   3000
         TabIndex        =   7
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblNPoint 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   6
         Top             =   720
         Width           =   615
      End
   End
   Begin Threed.SSFrame fraNumberOfBeds 
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      _Version        =   65536
      _ExtentX        =   9340
      _ExtentY        =   1508
      _StockProps     =   14
      Caption         =   "Number of Beds:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Spin.SpinButton spnNumberOfBeds 
         Height          =   255
         Left            =   3720
         TabIndex        =   5
         Top             =   360
         Width           =   135
         _Version        =   65536
         _ExtentX        =   238
         _ExtentY        =   450
         _StockProps     =   73
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
      End
      Begin VB.TextBox txtNumberOfBeds 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3000
         TabIndex        =   3
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Beds (in series)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   420
         Width           =   2595
      End
   End
   Begin Threed.SSCommand cmdOK 
      Height          =   375
      Left            =   2880
      TabIndex        =   27
      Top             =   4800
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "&OK"
   End
End
Attribute VB_Name = "frmOptionsInputParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Option Base 1

Dim MCT As Integer, Time_Step As Double
Dim NCT As Integer, FirstPt As Double, EndT As Double
Dim OldTimeParameters As TimeParametersType
Dim Temp_Text As String
Dim IsError As Integer
'------Begin Modification Hokanson: 11-Aug2000
Dim OldUnits(1 To 5) As Integer
'------End Modification Hokanson: 11-Aug2000
Dim OldNumberOfBeds As Integer
'------Begin Modification Hokanson: 11-Aug2000
Dim OldEPS_ErrorCriteriaForDGEARIntegrator As Double
Dim OldDH0_InitialTimeStepForDGEARIntegrator As Double
'------End Modification Hokanson: 11-Aug2000

Private Sub cboTimeParametersUnits_Click(index As Integer)
    Dim ValueToDisplay As Double

    Select Case index

       Case 0   'Total Run Time
            Select Case cboTimeParametersUnits(0).ListIndex
               Case TIME_MIN   'min
                    ValueToDisplay = NowProj.TimeParameters.FinalTime
               Case TIME_S     's
                    ValueToDisplay = NowProj.TimeParameters.FinalTime * TimeConversionFactor(TIME_S)
               Case TIME_HR    'hr
                    ValueToDisplay = NowProj.TimeParameters.FinalTime * TimeConversionFactor(TIME_HR)
               Case TIME_D     'd
                    ValueToDisplay = NowProj.TimeParameters.FinalTime * TimeConversionFactor(TIME_D)
            End Select
            txtTime(0).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
            TimeUnitsOnGraphs = cboTimeParametersUnits(0).ListIndex

       Case 1   'Inital Time
            Select Case cboTimeParametersUnits(1).ListIndex
               Case TIME_MIN   'min
                    ValueToDisplay = NowProj.TimeParameters.InitialTime
               Case TIME_S     's
                    ValueToDisplay = NowProj.TimeParameters.InitialTime * TimeConversionFactor(TIME_S)
               Case TIME_HR    'hr
                    ValueToDisplay = NowProj.TimeParameters.InitialTime * TimeConversionFactor(TIME_HR)
               Case TIME_D     'd
                    ValueToDisplay = NowProj.TimeParameters.InitialTime * TimeConversionFactor(TIME_D)
            End Select
            txtTime(1).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       Case 2   'Time step
            Select Case cboTimeParametersUnits(2).ListIndex
               Case TIME_MIN   'min
                    ValueToDisplay = NowProj.TimeParameters.TimeStep
               Case TIME_S     's
                    ValueToDisplay = NowProj.TimeParameters.TimeStep * TimeConversionFactor(TIME_S)
               Case TIME_HR    'hr
                    ValueToDisplay = NowProj.TimeParameters.TimeStep * TimeConversionFactor(TIME_HR)
               Case TIME_D     'd
                    ValueToDisplay = NowProj.TimeParameters.TimeStep * TimeConversionFactor(TIME_D)
            End Select
            txtTime(2).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

    End Select

End Sub

Private Sub cmdCancel_Click()
    Dim i As Integer
  
    NowProj.Bed.NumberOfBeds = OldNumberOfBeds
    NowProj.TimeParameters = OldTimeParameters

'------Begin Modification Hokanson: 11-Aug2000
    EPS_ErrorCriteriaForDGEARIntegrator = OldEPS_ErrorCriteriaForDGEARIntegrator
    DH0_InitialTimeStepForDGEARIntegrator = OldDH0_InitialTimeStepForDGEARIntegrator
'------End Modification Hokanson: 11-Aug2000

    'Set units back to original
    For i = 1 To 3
        cboTimeParametersUnits(i - 1).ListIndex = OldUnits(i)
    Next i
'------Begin Modification Hokanson: 11-Aug2000
    cboTimeParametersUnits(5 - 1).ListIndex = OldUnits(5)
'------End Modification Hokanson: 11-Aug2000

    frmOptionsInputParameters.Hide

End Sub

Private Sub cmdCancel_KeyPress(KeyAscii As Integer)
Call Key_Pressed_On_Control(KeyAscii)
End Sub

Private Sub cmdOK_Click()
    Dim NewTimeStep As Double, ValueToDisplay As Double, CurrentUnits As Integer

'    Input_Exist = True
    NowProj.NumAxialCollocationPoints = MCT
    NowProj.NumRadialCollocationPoints = NCT
    If NowProj.TimeParameters.InitialTime > NowProj.TimeParameters.FinalTime Then
      MsgBox "The first point is greater than the final point.", MB_ICONEXCLAMATION, App.title
      Exit Sub
    ElseIf NowProj.TimeParameters.TimeStep < ((NowProj.TimeParameters.FinalTime - NowProj.TimeParameters.InitialTime) / (Number_Points_Max - 1)) Then
      MsgBox "Time step is too small. The maximum number of points is " & Trim$(Str$(Number_Points_Max)) & ".", MB_ICONEXCLAMATION, App.title
      Exit Sub
    End If

    If (NowProj.Bed.NumberOfBeds = 1) Or (NowProj.TimeParameters.InitialTime < 0.00011) Then

    Else   'For beds in series, initial time must be approximately zero
       NowProj.TimeParameters.InitialTime = 0.0001
       NewTimeStep = (NowProj.TimeParameters.FinalTime - NowProj.TimeParameters.InitialTime) / (Number_Points_Max - 5)
       If NowProj.TimeParameters.TimeStep < NewTimeStep Then NowProj.TimeParameters.TimeStep = NewTimeStep
       MsgBox "For beds in series, the initial time must be approximately zero.  The initial time will automatically be adjusted to reflect this.  If necessary, the time step will also be adjusted.", MB_ICONINFORMATION
       CurrentUnits = cboTimeParametersUnits(1).ListIndex
       If CurrentUnits = 0 Then
          ValueToDisplay = NowProj.TimeParameters.InitialTime
       Else
          ValueToDisplay = NowProj.TimeParameters.InitialTime * TimeConversionFactor(CurrentUnits)
       End If
       txtTime(1) = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
       CurrentUnits = cboTimeParametersUnits(2).ListIndex
       If CurrentUnits = 0 Then
          ValueToDisplay = NowProj.TimeParameters.TimeStep
       Else
          ValueToDisplay = NowProj.TimeParameters.TimeStep * TimeConversionFactor(CurrentUnits)
       End If
       txtTime(2) = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
       Exit Sub
    End If

    frmOptionsInputParameters.Hide

End Sub

Private Sub cmdOK_KeyPress(KeyAscii As Integer)
Call Key_Pressed_On_Control(KeyAscii)
End Sub

Private Sub Form_Activate()
    Dim ValueToDisplay As Double
    Dim CurrentUnits As Integer

    OldNumberOfBeds = NowProj.Bed.NumberOfBeds
    txtNumberOfBeds = Format$(NowProj.Bed.NumberOfBeds, "0")

    NCT = NowProj.NumRadialCollocationPoints
    MCT = NowProj.NumAxialCollocationPoints
    lblNPoint(0) = Format$(MCT, "0")
    lblNPoint(1) = Format$(NCT, "0")

    CurrentUnits = cboTimeParametersUnits(0).ListIndex
    OldUnits(1) = CurrentUnits
    ValueToDisplay = NowProj.TimeParameters.FinalTime
    If CurrentUnits <> 0 Then ValueToDisplay = ValueToDisplay * TimeConversionFactor(CurrentUnits)
    txtTime(0) = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

    CurrentUnits = cboTimeParametersUnits(1).ListIndex
    OldUnits(2) = CurrentUnits
    ValueToDisplay = NowProj.TimeParameters.InitialTime
    If CurrentUnits <> 0 Then ValueToDisplay = ValueToDisplay * TimeConversionFactor(CurrentUnits)
    txtTime(1) = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

    CurrentUnits = cboTimeParametersUnits(2).ListIndex
    OldUnits(3) = CurrentUnits
    ValueToDisplay = NowProj.TimeParameters.TimeStep
    If CurrentUnits <> 0 Then ValueToDisplay = ValueToDisplay * TimeConversionFactor(CurrentUnits)
    txtTime(2) = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

'------Begin Modification Hokanson: 11-Aug2000
    OldEPS_ErrorCriteriaForDGEARIntegrator = EPS_ErrorCriteriaForDGEARIntegrator
    ValueToDisplay = EPS_ErrorCriteriaForDGEARIntegrator
    txtTime(3) = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

    OldDH0_InitialTimeStepForDGEARIntegrator = DH0_InitialTimeStepForDGEARIntegrator
    CurrentUnits = cboTimeParametersUnits(4).ListIndex
    OldUnits(5) = CurrentUnits
    ValueToDisplay = DH0_InitialTimeStepForDGEARIntegrator
    If CurrentUnits <> 0 Then ValueToDisplay = ValueToDisplay * TimeConversionFactor(CurrentUnits)
    txtTime(4) = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
'------End Modification Hokanson: 11-Aug2000
 
    OldTimeParameters = NowProj.TimeParameters

End Sub

Private Sub Form_Load()
    top = Screen.height / 2 - height / 2
    left = Screen.width / 2 - width / 2
'    Me.HelpContextID = Hlp_Algorithm_Parameter

    
End Sub

Private Sub Key_Pressed_On_Control(Ascii_Code As Integer)
  Select Case Ascii_Code
    Case 67, 99 'C,c
      cmdCancel_Click
    Case 79, 111 'O,o
      cmdOK_Click
  End Select
End Sub

Private Sub spnNumberOfBeds_SpinDown()

    If NowProj.Bed.NumberOfBeds = 1 Then
       Exit Sub
    Else
       NowProj.Bed.NumberOfBeds = NowProj.Bed.NumberOfBeds - 1
       txtNumberOfBeds = Format$(NowProj.Bed.NumberOfBeds, "0")
    End If

End Sub

Private Sub spnNumberOfBeds_SpinUp()

    If NowProj.Bed.NumberOfBeds = Maximum_Beds_In_Series Then
       Exit Sub
    Else
       NowProj.Bed.NumberOfBeds = NowProj.Bed.NumberOfBeds + 1
       txtNumberOfBeds = Format$(NowProj.Bed.NumberOfBeds, "0")
    End If

End Sub

Private Sub spnPoint_SpinDown(index As Integer)
   Select Case index
    Case 0
    If MCT > 1 Then
     MCT = MCT - 1
     lblNPoint(0) = Format$(MCT, "0")
     End If
    Case 1
    If NCT > 1 Then
     NCT = NCT - 1
     lblNPoint(1) = Format$(NCT, "0")
    End If
   End Select
End Sub

Private Sub spnPoint_SpinUp(index As Integer)
   Select Case index
    Case 0
    If MCT < MAX_AXIAL_COLLOCATION_POINTS Then
     MCT = MCT + 1
     lblNPoint(0) = Format$(MCT, "0")
     End If
    Case 1
    If NCT < MAX_RADIAL_COLLOCATION_POINTS Then
     NCT = NCT + 1
     lblNPoint(1) = Format$(NCT, "0")
    End If
   End Select

End Sub

Private Sub txtNumberOfBeds_GotFocus()
    Call TextGetFocus(txtNumberOfBeds, Temp_Text)
End Sub

Private Sub txtNumberOfBeds_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
       KeyAscii = 0
       SendKeys "{Tab}"
       Exit Sub
    End If

    Call NumberCheck(KeyAscii)

End Sub

Private Sub txtNumberOfBeds_LostFocus()
    Dim NewValue As Long, msg As String

    NewValue = CLng(txtNumberOfBeds)
    If (NewValue < 1) Or (NewValue > Maximum_Beds_In_Series) Then
       txtNumberOfBeds = Temp_Text
       msg = "Specified number of beds in series (" & Trim$(Str$(NewValue)) & ") was out of range (minimum = 1, maximum = " & Trim$(Str$(Maximum_Beds_In_Series)) & ").  Incorrect value was replaced by previous value."
       MsgBox msg, MB_ICONSTOP, "Error"
    Else
       NowProj.Bed.NumberOfBeds = NewValue
       txtNumberOfBeds = Format$(NewValue, "0")
    End If

End Sub

Private Sub txtTime_GotFocus(index As Integer)
    Call TextGetFocus(txtTime(index), Temp_Text)
End Sub

Private Sub txtTime_KeyPress(index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
       KeyAscii = 0
       SendKeys "{Tab}"
       Exit Sub
    End If

    Call NumberCheck(KeyAscii)

End Sub

Private Sub txtTime_LostFocus(index As Integer)
    Dim ValueChanged As Integer, NewValue As Double
    Dim msg As String, CurrentUnits As Integer
    Dim OldValue As Double, ValueToDisplay As Double
    Dim i As Integer

    Call TextHandleError(IsError, txtTime(index), Temp_Text)

    If Not IsError Then
       NewValue = CDbl(txtTime(index).Text)
       'Convert NewValue to Standard Units if Necessary
       Select Case index
          Case 0   'Total Run Time
               OldValue = NowProj.TimeParameters.FinalTime
               CurrentUnits = cboTimeParametersUnits(0).ListIndex
               If CurrentUnits <> 0 Then
                  NewValue = NewValue / TimeConversionFactor(CurrentUnits)
               End If
          Case 1   'Initial Time
               OldValue = NowProj.TimeParameters.InitialTime
               CurrentUnits = cboTimeParametersUnits(1).ListIndex
               If CurrentUnits <> 0 Then
                  NewValue = NewValue / TimeConversionFactor(CurrentUnits)
               End If
          Case 2   'Time Step
               OldValue = NowProj.TimeParameters.TimeStep
               CurrentUnits = cboTimeParametersUnits(2).ListIndex
               If CurrentUnits <> 0 Then
                  NewValue = NewValue / TimeConversionFactor(CurrentUnits)
               End If
'------Begin Modification Hokanson: 11-Aug2000
          Case 4   'DH0: Initial Time Step for DGEAR Integrator
               OldValue = DH0_InitialTimeStepForDGEARIntegrator
               CurrentUnits = cboTimeParametersUnits(4).ListIndex
               If CurrentUnits <> 0 Then
                  NewValue = NewValue / TimeConversionFactor(CurrentUnits)
               End If
 '------End Modification Hokanson: 11-Aug2000
       End Select

       Select Case index
          Case 0    'Total Run Time
             Call NumberChanged(ValueChanged, OldValue, NewValue)
             If ValueChanged Then
                If HaveValue(NewValue) Then
                   NowProj.TimeParameters.FinalTime = NewValue

                Else
                   txtTime(0).Text = Temp_Text
                   txtTime(0).SetFocus
                   Exit Sub
                End If
             End If

          Case 1    'Initial Time
             Call NumberChanged(ValueChanged, OldValue, NewValue)
             If ValueChanged Then
                If HaveValue(NewValue) Then
                   NowProj.TimeParameters.InitialTime = NewValue

                Else
                   txtTime(1).Text = Temp_Text
                   txtTime(1).SetFocus
                   Exit Sub
                End If
             End If

          Case 2    'Time Step
             Call NumberChanged(ValueChanged, OldValue, NewValue)
             If ValueChanged Then
                If HaveValue(NewValue) Then
                   NowProj.TimeParameters.TimeStep = NewValue

                Else
                   txtTime(2).Text = Temp_Text
                   txtTime(2).SetFocus
                   Exit Sub
                End If
             End If

'------Begin Modification Hokanson: 11-Aug2000
          Case 3    'EPS Error Criteria for DGEAR Integrator
             Call NumberChanged(ValueChanged, OldValue, NewValue)
             If ValueChanged Then
                If HaveValue(NewValue) Then
                   EPS_ErrorCriteriaForDGEARIntegrator = NewValue

                Else
                   txtTime(3).Text = Temp_Text
                   txtTime(3).SetFocus
                   Exit Sub
                End If
             End If

          Case 4    'DH0 Initial Time Step for DGEAR Integrator
             Call NumberChanged(ValueChanged, OldValue, NewValue)
             If ValueChanged Then
                If HaveValue(NewValue) Then
                   DH0_InitialTimeStepForDGEARIntegrator = NewValue

                Else
                   txtTime(4).Text = Temp_Text
                   txtTime(4).SetFocus
                   Exit Sub
                End If
             End If
'------End Modification Hokanson: 11-Aug2000

       End Select

    End If

End Sub

