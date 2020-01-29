VERSION 5.00
Begin VB.Form vp_form 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vapor Pressure"
   ClientHeight    =   4665
   ClientLeft      =   1215
   ClientTop       =   2805
   ClientWidth     =   8640
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
   ScaleHeight     =   4665
   ScaleWidth      =   8640
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   612
      Left            =   4920
      TabIndex        =   1
      Top             =   960
      Width           =   3492
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      Caption         =   "Accept Selected Vapor Pressure"
      Height          =   612
      Left            =   4920
      TabIndex        =   0
      Top             =   225
      Width           =   3492
   End
   Begin VB.TextBox txtVPmaximumT 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   7320
      TabIndex        =   10
      Text            =   "not visible"
      Top             =   3960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtVPminimumT 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   6120
      TabIndex        =   9
      Text            =   "not visible"
      Top             =   3960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtVPTemperature 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   4920
      TabIndex        =   8
      Top             =   3960
      Width           =   975
   End
   Begin VB.TextBox txtVaporPressureValue 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   2880
      TabIndex        =   7
      Top             =   3960
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   4
      Left            =   360
      TabIndex        =   5
      Top             =   3960
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   3
      Left            =   360
      TabIndex        =   4
      Top             =   3480
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   2
      Left            =   360
      TabIndex        =   3
      Top             =   3000
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label lblVPmaximumT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "not visible"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   2
      Left            =   7320
      TabIndex        =   35
      Top             =   3480
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label lblVPmaximumT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   7320
      TabIndex        =   34
      Top             =   3000
      Width           =   972
   End
   Begin VB.Label lblVPmaximumT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   7320
      TabIndex        =   33
      Top             =   2520
      Width           =   972
   End
   Begin VB.Label lblVPminimumT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "not visible"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   2
      Left            =   6120
      TabIndex        =   32
      Top             =   3480
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label lblVPminimumT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   6120
      TabIndex        =   31
      Top             =   3000
      Width           =   972
   End
   Begin VB.Label lblVPminimumT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   6120
      TabIndex        =   30
      Top             =   2520
      Width           =   972
   End
   Begin VB.Label lblVPTemperature 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   2
      Left            =   4920
      TabIndex        =   29
      Top             =   3480
      Width           =   972
   End
   Begin VB.Label lblVPTemperature 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   4920
      TabIndex        =   28
      Top             =   3000
      Width           =   972
   End
   Begin VB.Label lblVPTemperature 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   4920
      TabIndex        =   27
      Top             =   2520
      Width           =   972
   End
   Begin VB.Label lblVaporPressureValue 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   2
      Left            =   2880
      TabIndex        =   26
      Top             =   3480
      Width           =   1812
   End
   Begin VB.Label lblVaporPressureValue 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   2880
      TabIndex        =   25
      Top             =   3000
      Width           =   1812
   End
   Begin VB.Label lblVaporPressureValue 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   2880
      TabIndex        =   24
      Top             =   2520
      Width           =   1812
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   8400
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   4560
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   2
      Height          =   1335
      Left            =   240
      Top             =   240
      Width           =   4335
   End
   Begin VB.Label lblCurrentValues 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   23
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label lblCurrentValues 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   22
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label lblCurrentInformation 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Source"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   21
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label lblCurrentInformation 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Value"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   20
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Current Vapor Pressure Information"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   19
      Top             =   360
      Width           =   4095
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      Height          =   2655
      Left            =   240
      Top             =   1800
      Width           =   8175
   End
   Begin VB.Label lblVPTempLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Temp."
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4920
      TabIndex        =   18
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label lblSourceLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "User Input"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   720
      TabIndex        =   6
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label lblSourceLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Superfund"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   720
      TabIndex        =   17
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Label lblSourceLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Antoine's Equation"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   16
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label lblSourceLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DIPPR801"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   15
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Source"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   720
      TabIndex        =   14
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label lblVPmaxTLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tmax"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7320
      TabIndex        =   13
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label lblVPminTLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tmin"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6120
      TabIndex        =   12
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label lblVPLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Vapor Pressure"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2880
      TabIndex        =   11
      Top             =   1920
      Width           =   1815
   End
End
Attribute VB_Name = "vp_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PastVPInputValue As String
Dim PastVPInputTemp As String

Private Sub cmdCancel_Click()
    Dim i As Integer
    Dim SelectedOption As Integer   'Option selected permanently by the user (the option corresponding to the number on the main screen)

    Select Case phprop.VaporPressure.CurrentSelection.choice
       Case VAPOR_PRESSURE_DATABASE
          For i = 1 To 3
              If Option1(i).Enabled = True Then
                 SelectedOption = i
                 Exit For
              End If
          Next i
       Case VAPOR_PRESSURE_INPUT
          SelectedOption = 4
          txtVaporPressureValue(3).Text = PastVPInputValue
          txtVPTemperature(3).Text = PastVPInputTemp
       Case Else
          vp_form.Hide
          Exit Sub
    End Select

    If Not Option1(SelectedOption).Value Then Option1(SelectedOption).Value = True

    vp_form.Hide
End Sub

Private Sub cmdok_Click()
    Dim OptionSelected As Integer
    Dim ValueToDisplayIndex As Integer
    Dim i As Integer
    Dim NumContaminantInList As Integer

'*** Pass new selected value back to main screen
    For i = 1 To 4
        If Option1(i).Value Then
           OptionSelected = i
           Exit For
        End If
    Next i

    Select Case OptionSelected
       Case 1, 2, 3
          ValueToDisplayIndex = VAPOR_PRESSURE_DATABASE
       Case 4
          If Not PROPAVAILABLE(VAPOR_PRESSURE_INPUT) Then
             MsgBox "User Input can not be selected without first entering a value", MB_ICONSTOP, "Error"
             txtVaporPressureValue(3).SetFocus
             Exit Sub
          End If
          ValueToDisplayIndex = VAPOR_PRESSURE_INPUT
    End Select

    If ValueToDisplayIndex <> phprop.VaporPressure.CurrentSelection.choice Then
       phprop.VaporPressure.CurrentSelection.choice = ValueToDisplayIndex
       Call DisplayVaporPressureMainScreen(ValueToDisplayIndex)
    ElseIf ValueToDisplayIndex = VAPOR_PRESSURE_INPUT Then
       Call DisplayVaporPressureMainScreen(ValueToDisplayIndex)
    End If

    vp_form.Hide

'Recalculate Henry's Constant using selected vapor pressure

'          frmWaitForCalculations.Show
'          frmWaitForCalculations.Refresh

          contam_prop_form!lblContaminantProperties(2).Caption = ""

          Screen.MousePointer = 11   'Hourglass

          Call CalculateHenrysConstant
          contam_prop_form.Refresh

          Screen.MousePointer = 0    'Arrow

'          frmWaitForCalculations.Hide

          NumContaminantInList = contam_prop_form!cboSelectContaminant.ListIndex + 1
          PropContaminant(NumContaminantInList) = phprop

End Sub

Private Sub Form_Activate()
    
  Call centerform_relative(contam_prop_form, Me)
    
    PastVPInputValue = txtVaporPressureValue(3).Text
    PastVPInputTemp = txtVPTemperature(3).Text
End Sub

Private Sub Form_Load()

  Call centerform_relative(contam_prop_form, Me)
    
  If (DemoMode) Then cmdOK.Enabled = False


End Sub

Private Sub lblSourceLabel_Click(Index As Integer)
    Dim i As Integer

    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.VaporPressure.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If
End Sub

Private Sub lblVaporPressureValue_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim msg As String
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.VaporPressure.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

    If Button <> 2 Then Exit Sub

    If lblVaporPressureValue(Index).Caption = "Not Available" Then
       Select Case Index
          Case 0   'DIPPR801
               msg = "Vapor Pressure from DIPPR801 is not available in the StEPP database."
          Case 1   'Antoine's Equation
               msg = "Vapor Pressure from Antoine's Equation is not available in the StEPP database."
          Case 2   'Superfund
               msg = "Vapor Pressure from Superfund is not available in the StEPP database."
       End Select

       MsgBox msg, MB_ICONINFORMATION, Trim$(phprop.Name) & " - Data Unavailable"
       Exit Sub
    End If

    If phprop.VaporPressure.database.error = 0 Then Exit Sub

    MsgBox ErrorMsg(phprop.VaporPressure.database.error), MB_ICONINFORMATION, Trim$(phprop.Name) & " - Warning"

End Sub

Private Sub lblVPmaximumT_Click(Index As Integer)
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.VaporPressure.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

End Sub

Private Sub lblVPminimumT_Click(Index As Integer)
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.VaporPressure.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

End Sub

Private Sub lblVPTemperature_Click(Index As Integer)
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.VaporPressure.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

End Sub

Private Sub Option1_Click(Index As Integer)
    Dim i As Integer, SourceIndex As Integer

    SourceIndex = Index - 1
    If SourceIndex = hilight.VaporPressure.PreviousIndex Then Exit Sub
    lblSourceLabel(SourceIndex).BackColor = &H800000
    lblSourceLabel(SourceIndex).ForeColor = &H80000005
    i = hilight.VaporPressure.PreviousIndex
    hilight.VaporPressure.PreviousIndex = SourceIndex
    If i = -1 Then Exit Sub
    If Option1(i + 1).Enabled = False Then Exit Sub
    lblSourceLabel(i).BackColor = &H80000005
    lblSourceLabel(i).ForeColor = &H80000008

End Sub

Private Sub txtVaporPressureValue_GotFocus(Index As Integer)
  Call GotFocus_Handle(Me, txtVaporPressureValue(Index), Temp_Text)
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.VaporPressure.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If
End Sub

Private Sub txtVaporPressureValue_KeyPress(Index As Integer, keyascii As Integer)

    If keyascii = 13 Then
       keyascii = 0
       txtVPTemperature(Index).SetFocus
       Exit Sub
    End If
    Call NumberCheck(keyascii)

End Sub

Private Sub txtVaporPressureValue_LostFocus(Index As Integer)
    Dim msg As String, response As Integer
    Dim Answer As Integer
    Dim IsError As Integer
    Dim ValueChanged As Integer
    Dim EnglishValue As Double, SIValue As Double
    Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtVaporPressureValue(Index))) Then
     Exit Sub
   End If

   flag_ok = True

    If txtVaporPressureValue(Index).Text = "" Then
       phprop.VaporPressure.input.Value = -1#
       PROPAVAILABLE(VAPOR_PRESSURE_INPUT) = False
       Call LostFocus_Handle(Me, txtVaporPressureValue(Index), flag_ok)
       Exit Sub
    End If

    Call TextHandleError(IsError, txtVaporPressureValue(Index), Temp_Text)
    If Not IsError Then
       If Not HaveNumber(CDbl(txtVaporPressureValue(Index).Text)) Then
          txtVaporPressureValue(Index).Text = Temp_Text
          txtVaporPressureValue(Index).SetFocus
       Call LostFocus_Handle(Me, txtVaporPressureValue(Index), flag_ok)
          Exit Sub
       End If

       Call TextNumberChanged(ValueChanged, txtVaporPressureValue(Index), Temp_Text)

       If ValueChanged Then
          If CurrentUnits = SIUnits Then
             phprop.VaporPressure.input.Value = CDbl(txtVaporPressureValue(Index).Text)
          Else
             EnglishValue = CDbl(txtVaporPressureValue(Index).Text)
             Call VPENSI(SIValue, EnglishValue)
             phprop.VaporPressure.input.Value = SIValue
          End If
          
          PROPAVAILABLE(VAPOR_PRESSURE_INPUT) = True
       Else
       Call LostFocus_Handle(Me, txtVaporPressureValue(Index), flag_ok)
          Exit Sub
       End If

    End If
       Call LostFocus_Handle(Me, txtVaporPressureValue(Index), flag_ok)

End Sub

Private Sub txtVPmaximumT_GotFocus(Index As Integer)
  Call GotFocus_Handle(Me, txtVPmaximumT(Index), Temp_Text)

End Sub

Private Sub txtVPmaximumT_LostFocus(Index As Integer)
Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtVPmaximumT(Index))) Then
     Exit Sub
   End If

   flag_ok = True
  Call LostFocus_Handle(Me, txtVPmaximumT(Index), flag_ok)

End Sub

Private Sub txtVPminimumT_GotFocus(Index As Integer)
  Call GotFocus_Handle(Me, txtVPminimumT(Index), Temp_Text)

End Sub

Private Sub txtVPminimumT_LostFocus(Index As Integer)
Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtVPminimumT(Index))) Then
     Exit Sub
   End If

   flag_ok = True
  Call LostFocus_Handle(Me, txtVPminimumT(Index), flag_ok)

End Sub

Private Sub txtVPTemperature_GotFocus(Index As Integer)
  Call GotFocus_Handle(Me, txtVPTemperature(Index), Temp_Text)
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.VaporPressure.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If
End Sub

Private Sub txtVPTemperature_KeyPress(Index As Integer, keyascii As Integer)

    If keyascii = 13 Then
       keyascii = 0
       cmdOK.SetFocus
       Exit Sub
    End If
    Call NumberCheck(keyascii)

End Sub

Private Sub txtVPTemperature_LostFocus(Index As Integer)

    Dim msg As String, response As Integer
    Dim Answer As Integer
    Dim IsError As Integer
    Dim ValueChanged As Integer
    Dim EnglishValue As Double, SIValue As Double
    Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtVPTemperature(Index))) Then
     Exit Sub
   End If

   flag_ok = True

    If txtVPTemperature(Index).Text = "" Then
       phprop.VaporPressure.input.temperature = -1E+25
        Call LostFocus_Handle(Me, txtVPTemperature(Index), flag_ok)
       Exit Sub
    End If

    Call TextHandleError(IsError, txtVPTemperature(Index), Temp_Text)
    If Not IsError Then
       If Not HaveNumber(CDbl(txtVPTemperature(Index).Text)) Then
          txtVPTemperature(Index).Text = Temp_Text
          txtVPTemperature(Index).SetFocus
        Call LostFocus_Handle(Me, txtVPTemperature(Index), flag_ok)
          Exit Sub
       End If

       Call TextNumberChanged(ValueChanged, txtVPTemperature(Index), Temp_Text)

       If ValueChanged Then
          If CurrentUnits = SIUnits Then
             phprop.VaporPressure.input.temperature = CDbl(txtVPTemperature(Index).Text)
          Else
             EnglishValue = CDbl(txtVPTemperature(Index).Text)
             Call TEMPENSI(SIValue, EnglishValue)
             phprop.VaporPressure.input.temperature = SIValue
          End If
          
       Else
        Call LostFocus_Handle(Me, txtVPTemperature(Index), flag_ok)
          Exit Sub
       End If

    End If
        Call LostFocus_Handle(Me, txtVPTemperature(Index), flag_ok)

End Sub

