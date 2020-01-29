VERSION 2.00
Begin Form frmOptionsInputParameters 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Algorithm Parameters"
   ClientHeight    =   5340
   ClientLeft      =   885
   ClientTop       =   1545
   ClientWidth     =   5895
   ControlBox      =   0   'False
   Height          =   5745
   Left            =   825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   5895
   Top             =   1200
   Width           =   6015
   Begin SSFrame fraNumberOfBeds 
      Caption         =   "Number of Beds"
      ForeColor       =   &H00000000&
      Height          =   675
      Left            =   600
      TabIndex        =   17
      Top             =   120
      Width           =   4575
      Begin SpinButton spnNumberOfBeds 
         Height          =   285
         Left            =   3600
         Top             =   240
         Width           =   255
      End
      Begin TextBox txtNumberOfBeds 
         Height          =   285
         Left            =   2880
         TabIndex        =   18
         Top             =   240
         Width           =   735
      End
      Begin Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Beds (in series)"
         Height          =   255
         Left            =   180
         TabIndex        =   19
         Top             =   300
         Width           =   2595
      End
   End
   Begin SSFrame Frame3D1 
      Caption         =   "Other Parameters:"
      ForeColor       =   &H00000000&
      Height          =   2235
      Left            =   120
      TabIndex        =   5
      Top             =   2340
      Width           =   5655
      Begin TextBox txtTime 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   4
         Left            =   3120
         TabIndex        =   23
         Text            =   "txtTime"
         Top             =   1800
         Width           =   1215
      End
      Begin ComboBox cboTimeParametersUnits 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   4
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1095
      End
      Begin TextBox txtTime 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   3
         Left            =   3120
         TabIndex        =   20
         Text            =   "txtTime"
         Top             =   1440
         Width           =   1215
      End
      Begin ComboBox cboTimeParametersUnits 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   2
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1095
      End
      Begin ComboBox cboTimeParametersUnits 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   1
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   720
         Width           =   1095
      End
      Begin ComboBox cboTimeParametersUnits 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   0
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   360
         Width           =   1095
      End
      Begin TextBox txtTime 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   2
         Left            =   3120
         TabIndex        =   11
         Text            =   "txtTime"
         Top             =   1080
         Width           =   1215
      End
      Begin TextBox txtTime 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   1
         Left            =   3120
         TabIndex        =   10
         Text            =   "txtTime"
         Top             =   720
         Width           =   1215
      End
      Begin TextBox txtTime 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   0
         Left            =   3120
         TabIndex        =   9
         Text            =   "txtTime"
         Top             =   360
         Width           =   1215
      End
      Begin Label lblTime 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Initial Integrator Time Step, DH0"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   24
         Top             =   1800
         Width           =   2895
      End
      Begin Label lblTime 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Integrator Error Criteria, EPS"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   21
         Top             =   1440
         Width           =   2895
      End
      Begin Label lblTime 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Time Step:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   2895
      End
      Begin Label lblTime 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "First point displayed:"
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   2895
      End
      Begin Label lblTime 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Run Time:"
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   2895
      End
   End
   Begin SSCommand cmdOK 
      Caption         =   "&OK"
      Height          =   435
      Left            =   4320
      TabIndex        =   13
      Top             =   4740
      Width           =   1455
   End
   Begin SSCommand cmdCancel 
      Caption         =   "&Cancel"
      Height          =   435
      Left            =   120
      TabIndex        =   12
      Top             =   4740
      Width           =   1455
   End
   Begin SSFrame fraPoint 
      Caption         =   "Number of collocation points:"
      ForeColor       =   &H00000000&
      Height          =   1332
      Left            =   600
      TabIndex        =   0
      Top             =   900
      Width           =   4572
      Begin SpinButton spnPoint 
         Height          =   252
         Index           =   1
         Left            =   3120
         Top             =   840
         Width           =   252
      End
      Begin SpinButton spnPoint 
         Height          =   252
         Index           =   0
         Left            =   3120
         Top             =   360
         Width           =   252
      End
      Begin Label lblNPoint 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   252
         Index           =   1
         Left            =   2520
         TabIndex        =   4
         Top             =   840
         Width           =   612
      End
      Begin Label lblNPoint 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   252
         Index           =   0
         Left            =   2520
         TabIndex        =   3
         Top             =   360
         Width           =   612
      End
      Begin Label lblText 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Radial Direction"
         Height          =   252
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   1932
      End
      Begin Label lblText 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Axial direction"
         Height          =   252
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1932
      End
   End
End
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

Sub cboTimeParametersUnits_Click (Index As Integer)
    Dim ValueToDisplay As Double

    Select Case Index

       Case 0   'Total Run Time
            Select Case cboTimeParametersUnits(0).ListIndex
               Case TIME_MIN   'min
                    ValueToDisplay = TimeParameters.FinalTime
               Case TIME_S     's
                    ValueToDisplay = TimeParameters.FinalTime * TimeConversionFactor(TIME_S)
               Case TIME_HR    'hr
                    ValueToDisplay = TimeParameters.FinalTime * TimeConversionFactor(TIME_HR)
               Case TIME_D     'd
                    ValueToDisplay = TimeParameters.FinalTime * TimeConversionFactor(TIME_D)
            End Select
            txtTime(0).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
            TimeUnitsOnGraphs = cboTimeParametersUnits(0).ListIndex

       Case 1   'Inital Time
            Select Case cboTimeParametersUnits(1).ListIndex
               Case TIME_MIN   'min
                    ValueToDisplay = TimeParameters.InitialTime
               Case TIME_S     's
                    ValueToDisplay = TimeParameters.InitialTime * TimeConversionFactor(TIME_S)
               Case TIME_HR    'hr
                    ValueToDisplay = TimeParameters.InitialTime * TimeConversionFactor(TIME_HR)
               Case TIME_D     'd
                    ValueToDisplay = TimeParameters.InitialTime * TimeConversionFactor(TIME_D)
            End Select
            txtTime(1).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       Case 2   'Time step
            Select Case cboTimeParametersUnits(2).ListIndex
               Case TIME_MIN   'min
                    ValueToDisplay = TimeParameters.TimeStep
               Case TIME_S     's
                    ValueToDisplay = TimeParameters.TimeStep * TimeConversionFactor(TIME_S)
               Case TIME_HR    'hr
                    ValueToDisplay = TimeParameters.TimeStep * TimeConversionFactor(TIME_HR)
               Case TIME_D     'd
                    ValueToDisplay = TimeParameters.TimeStep * TimeConversionFactor(TIME_D)
            End Select
            txtTime(2).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

    End Select

End Sub

Sub cmdCancel_Click ()
    Dim i As Integer
  
    Bed.NumberOfBeds = OldNumberOfBeds
    TimeParameters = OldTimeParameters

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

Sub cmdCancel_KeyPress (KeyAscii As Integer)
Call Key_Pressed_On_Control(KeyAscii)
End Sub

Sub cmdOK_Click ()
    Dim NewTimeStep As Double, ValueToDisplay As Double, CurrentUnits As Integer

'    Input_Exist = True
    NumAxialCollocationPoints = MCT
    NumRadialCollocationPoints = NCT
    If TimeParameters.InitialTime > TimeParameters.FinalTime Then
      MsgBox "The first point is greater than the final point.", MB_ICONEXCLAMATION, Application_Name
      Exit Sub
    ElseIf TimeParameters.TimeStep < ((TimeParameters.FinalTime - TimeParameters.InitialTime) / (Number_Points_Max - 1)) Then
      MsgBox "Time step is too small. The maximum number of points is " & Trim$(Str$(Number_Points_Max)) & ".", MB_ICONEXCLAMATION, Application_Name
      Exit Sub
    End If

    If (Bed.NumberOfBeds = 1) Or (TimeParameters.InitialTime < .00011) Then

    Else   'For beds in series, initial time must be approximately zero
       TimeParameters.InitialTime = .0001
       NewTimeStep = (TimeParameters.FinalTime - TimeParameters.InitialTime) / (Number_Points_Max - 5)
       If TimeParameters.TimeStep < NewTimeStep Then TimeParameters.TimeStep = NewTimeStep
       MsgBox "For beds in series, the initial time must be approximately zero.  The initial time will automatically be adjusted to reflect this.  If necessary, the time step will also be adjusted.", MB_ICONINFORMATION
       CurrentUnits = cboTimeParametersUnits(1).ListIndex
       If CurrentUnits = 0 Then
          ValueToDisplay = TimeParameters.InitialTime
       Else
          ValueToDisplay = TimeParameters.InitialTime * TimeConversionFactor(CurrentUnits)
       End If
       txtTime(1) = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
       CurrentUnits = cboTimeParametersUnits(2).ListIndex
       If CurrentUnits = 0 Then
          ValueToDisplay = TimeParameters.TimeStep
       Else
          ValueToDisplay = TimeParameters.TimeStep * TimeConversionFactor(CurrentUnits)
       End If
       txtTime(2) = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
       Exit Sub
    End If

    frmOptionsInputParameters.Hide

End Sub

Sub cmdOK_KeyPress (KeyAscii As Integer)
Call Key_Pressed_On_Control(KeyAscii)
End Sub

Sub Form_Activate ()
    Dim ValueToDisplay As Double
    Dim CurrentUnits As Integer

    OldNumberOfBeds = Bed.NumberOfBeds
    txtNumberOfBeds = Format$(Bed.NumberOfBeds, "0")

    NCT = NumRadialCollocationPoints
    MCT = NumAxialCollocationPoints
    lblNPoint(0) = Format$(MCT, "0")
    lblNPoint(1) = Format$(NCT, "0")

    CurrentUnits = cboTimeParametersUnits(0).ListIndex
    OldUnits(1) = CurrentUnits
    ValueToDisplay = TimeParameters.FinalTime
    If CurrentUnits <> 0 Then ValueToDisplay = ValueToDisplay * TimeConversionFactor(CurrentUnits)
    txtTime(0) = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

    CurrentUnits = cboTimeParametersUnits(1).ListIndex
    OldUnits(2) = CurrentUnits
    ValueToDisplay = TimeParameters.InitialTime
    If CurrentUnits <> 0 Then ValueToDisplay = ValueToDisplay * TimeConversionFactor(CurrentUnits)
    txtTime(1) = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

    CurrentUnits = cboTimeParametersUnits(2).ListIndex
    OldUnits(3) = CurrentUnits
    ValueToDisplay = TimeParameters.TimeStep
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
 
    OldTimeParameters = TimeParameters

End Sub

Sub Form_Load ()
    Top = Screen.Height / 2 - Height / 2
    Left = Screen.Width / 2 - Width / 2
'    Me.HelpContextID = Hlp_Algorithm_Parameter

    
End Sub

Sub Key_Pressed_On_Control (Ascii_Code As Integer)
  Select Case Ascii_Code
    Case 67, 99 'C,c
      cmdCancel_Click
    Case 79, 111'O,o
      cmdOK_Click
  End Select
End Sub

Sub spnNumberOfBeds_SpinDown ()

    If Bed.NumberOfBeds = 1 Then
       Exit Sub
    Else
       Bed.NumberOfBeds = Bed.NumberOfBeds - 1
       txtNumberOfBeds = Format$(Bed.NumberOfBeds, "0")
    End If

End Sub

Sub spnNumberOfBeds_SpinUp ()

    If Bed.NumberOfBeds = Maximum_Beds_In_Series Then
       Exit Sub
    Else
       Bed.NumberOfBeds = Bed.NumberOfBeds + 1
       txtNumberOfBeds = Format$(Bed.NumberOfBeds, "0")
    End If

End Sub

Sub spnPoint_SpinDown (Index As Integer)
   Select Case Index
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

Sub spnPoint_SpinUp (Index As Integer)
   Select Case Index
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

Sub txtNumberOfBeds_GotFocus ()
    Call TextGetFocus(txtNumberOfBeds, Temp_Text)
End Sub

Sub txtNumberOfBeds_KeyPress (KeyAscii As Integer)

    If KeyAscii = 13 Then
       KeyAscii = 0
       SendKeys "{Tab}"
       Exit Sub
    End If

    Call NumberCheck(KeyAscii)

End Sub

Sub txtNumberOfBeds_LostFocus ()
    Dim NewValue As Long, msg As String

    NewValue = CLng(txtNumberOfBeds)
    If (NewValue < 1) Or (NewValue > Maximum_Beds_In_Series) Then
       txtNumberOfBeds = Temp_Text
       msg = "Specified number of beds in series (" & Trim$(Str$(NewValue)) & ") was out of range (minimum = 1, maximum = " & Trim$(Str$(Maximum_Beds_In_Series)) & ").  Incorrect value was replaced by previous value."
       MsgBox msg, MB_ICONSTOP, "Error"
    Else
       Bed.NumberOfBeds = NewValue
       txtNumberOfBeds = Format$(NewValue, "0")
    End If

End Sub

Sub txtTime_GotFocus (Index As Integer)
    Call TextGetFocus(txtTime(Index), Temp_Text)
End Sub

Sub txtTime_KeyPress (Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
       KeyAscii = 0
       SendKeys "{Tab}"
       Exit Sub
    End If

    Call NumberCheck(KeyAscii)

End Sub

Sub txtTime_LostFocus (Index As Integer)
    Dim ValueChanged As Integer, NewValue As Double
    Dim msg As String, CurrentUnits As Integer
    Dim OldValue As Double, ValueToDisplay As Double
    Dim i As Integer

    Call TextHandleError(IsError, txtTime(Index), Temp_Text)

    If Not IsError Then
       NewValue = CDbl(txtTime(Index).Text)
       'Convert NewValue to Standard Units if Necessary
       Select Case Index
          Case 0   'Total Run Time
               OldValue = TimeParameters.FinalTime
               CurrentUnits = cboTimeParametersUnits(0).ListIndex
               If CurrentUnits <> 0 Then
                  NewValue = NewValue / TimeConversionFactor(CurrentUnits)
               End If
          Case 1   'Initial Time
               OldValue = TimeParameters.InitialTime
               CurrentUnits = cboTimeParametersUnits(1).ListIndex
               If CurrentUnits <> 0 Then
                  NewValue = NewValue / TimeConversionFactor(CurrentUnits)
               End If
          Case 2   'Time Step
               OldValue = TimeParameters.TimeStep
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

       Select Case Index
          Case 0    'Total Run Time
             Call NumberChanged(ValueChanged, OldValue, NewValue)
             If ValueChanged Then
                If HaveValue(NewValue) Then
                   TimeParameters.FinalTime = NewValue

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
                   TimeParameters.InitialTime = NewValue

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
                   TimeParameters.TimeStep = NewValue

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

