VERSION 5.00
Begin VB.Form frmWaterSurfaceTension 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Water Surface Tension"
   ClientHeight    =   3585
   ClientLeft      =   795
   ClientTop       =   3570
   ClientWidth     =   8475
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
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3585
   ScaleWidth      =   8475
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   2880
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   255
   End
   Begin VB.TextBox txtH2OSurfTensmaximumT 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   7320
      TabIndex        =   7
      Text            =   "not visible"
      Top             =   2880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtH2OSurfTensminimumT 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   6120
      TabIndex        =   6
      Text            =   "not visible"
      Top             =   2880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtH2OSurfTensTemperature 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   4920
      TabIndex        =   5
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox txtWaterSurfaceTensionValue 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   3240
      TabIndex        =   4
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   615
      Left            =   4920
      TabIndex        =   1
      Top             =   840
      Width           =   3495
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      Caption         =   "Accept Selected H2O Surface Tension"
      Height          =   615
      Left            =   4920
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   1692
      Left            =   120
      Top             =   1680
      Width           =   8292
   End
   Begin VB.Label lblH2OSurfTensmaximumT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   7320
      TabIndex        =   23
      Top             =   2400
      Width           =   972
   End
   Begin VB.Label lblH2OSurfTensminimumT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   6120
      TabIndex        =   22
      Top             =   2400
      Width           =   972
   End
   Begin VB.Label lblH2OSurfTensTemperature 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   4920
      TabIndex        =   21
      Top             =   2400
      Width           =   972
   End
   Begin VB.Label lblWaterSurfaceTensionValue 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   3240
      TabIndex        =   20
      Top             =   2400
      Width           =   1452
   End
   Begin VB.Label lblSource 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "User Input"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   19
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label lblSource 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cummins, et. al. (1983)"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   18
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   8400
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Source"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   600
      TabIndex        =   17
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label lblH2OSTLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Surf. Tens."
      ForeColor       =   &H00000000&
      Height          =   400
      Left            =   3240
      TabIndex        =   16
      Top             =   1750
      Width           =   1455
   End
   Begin VB.Label lblH2OSTTempLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Temp."
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4920
      TabIndex        =   15
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label lblH2OSTmaxTLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tmax"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7320
      TabIndex        =   14
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label lblH2OSTminTLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tmin"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6120
      TabIndex        =   13
      Top             =   1800
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4440
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   2
      Height          =   1335
      Left            =   120
      Top             =   120
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
      Left            =   2040
      TabIndex        =   12
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label lblCurrentValues 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   11
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label lblCurrentInformation 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Source"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   10
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label lblCurrentInformation 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Value"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Current Water Surface Tension Information"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "frmWaterSurfaceTension"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PastH2OSurfTensInputValue As String
Dim PastH2OSurfTensInputTemp As String

Private Sub cmdCancel_Click()
    Dim SelectedOption As Integer   'Option selected permanently by the user (the option corresponding to the number on the main screen)

    Select Case phprop.WaterSurfaceTension.CurrentSelection.choice
       Case WATER_SURF_TENSION_CORRELATION
          SelectedOption = 1
       Case WATER_SURF_TENSION_INPUT
          SelectedOption = 2
          txtwatersurfacetensionvalue(1).Text = PastH2OSurfTensInputValue
          txth2osurftenstemperature(1).Text = PastH2OSurfTensInputTemp
       Case Else
          frmWaterSurfaceTension.Hide
          Exit Sub
    End Select

    If Not Option1(SelectedOption).Value Then Option1(SelectedOption).Value = True

    frmWaterSurfaceTension.Hide
End Sub

Private Sub cmdok_Click()
    Dim OptionSelected As Integer
    Dim ValueToDisplayIndex As Integer
    Dim i As Integer

'*** Pass new selected value back to main screen
    For i = 1 To 2
        If Option1(i).Value Then
           OptionSelected = i
           Exit For
        End If
    Next i

    Select Case OptionSelected
       Case 1
          ValueToDisplayIndex = WATER_SURF_TENSION_CORRELATION
       Case 2
          If Not PROPAVAILABLE(WATER_SURF_TENSION_INPUT) Then
             MsgBox "User Input can not be selected without first entering a value", MB_ICONSTOP, "Error"
             txtwatersurfacetensionvalue(1).SetFocus
             Exit Sub
          End If
          ValueToDisplayIndex = WATER_SURF_TENSION_INPUT
    End Select

    If ValueToDisplayIndex <> phprop.WaterSurfaceTension.CurrentSelection.choice Then
       phprop.WaterSurfaceTension.CurrentSelection.choice = ValueToDisplayIndex
       Call DisplayWaterSurfaceTensionMainScreen(ValueToDisplayIndex)
    ElseIf ValueToDisplayIndex = WATER_SURF_TENSION_INPUT Then
       Call DisplayWaterSurfaceTensionMainScreen(ValueToDisplayIndex)
    End If

    frmWaterSurfaceTension.Hide

End Sub

Private Sub Form_Activate()
    
  Call centerform_relative(contam_prop_form, Me)
    
    PastH2OSurfTensInputValue = txtwatersurfacetensionvalue(1).Text
    PastH2OSurfTensInputTemp = txth2osurftenstemperature(1).Text
End Sub

Private Sub Form_Load()

    If (DemoMode) Then cmdok.Enabled = False

  Call centerform_relative(contam_prop_form, Me)

End Sub

Private Sub lblH2OSurfTensmaximumT_Click(Index As Integer)
    
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.WaterSurfaceTension.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

End Sub

Private Sub lblH2OSurfTensminimumT_Click(Index As Integer)
    
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.WaterSurfaceTension.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

End Sub

Private Sub lblH2OSurfTensTemperature_Click(Index As Integer)
    
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.WaterSurfaceTension.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

End Sub

Private Sub lblSource_Click(Index As Integer)
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.WaterSurfaceTension.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

End Sub

Private Sub lblWaterSurfaceTensionValue_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ErrorFlag As Integer
    
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.WaterSurfaceTension.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

    If Button <> 2 Then Exit Sub

    Select Case Index
       Case 0  'Correlation from Cummins and Westrick (1983)
          ErrorFlag = phprop.WaterSurfaceTension.correlation.error
    End Select

    If ErrorFlag = 0 Then Exit Sub

    If ErrorFlag < 0 Then
       MsgBox ErrorMsg(ErrorFlag), MB_ICONINFORMATION, "Water - Data Unavailable"
    End If

    If ErrorFlag > 0 Then
       MsgBox ErrorMsg(ErrorFlag), MB_ICONINFORMATION, "Water - Warning"
    End If

End Sub

Private Sub Option1_Click(Index As Integer)
    Dim i As Integer, SourceIndex As Integer

    SourceIndex = Index - 1
    If SourceIndex = hilight.WaterSurfaceTension.PreviousIndex Then Exit Sub
    lblSource(SourceIndex).BackColor = &H800000
    lblSource(SourceIndex).ForeColor = &H80000005
    i = hilight.WaterSurfaceTension.PreviousIndex
    hilight.WaterSurfaceTension.PreviousIndex = SourceIndex
    If i = -1 Then Exit Sub
    If Option1(i + 1).Enabled = False Then Exit Sub
    lblSource(i).BackColor = &H80000005
    lblSource(i).ForeColor = &H80000008

End Sub

Private Sub txtH2OSurfTensmaximumT_GotFocus(Index As Integer)
  Call gotfocus_handle(Me, txtH2oSurftensmaximumT(Index), Temp_Text)

End Sub

Private Sub txtH2OSurfTensmaximumT_KeyPress(Index As Integer, keyascii As Integer)
    If Index <> 1 Then keyascii = 0
End Sub

Private Sub txtH2OSurfTensmaximumT_LostFocus(Index As Integer)
Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtH2oSurftensmaximumT(Index))) Then
     Exit Sub
   End If

   flag_ok = True
  Call LostFocus_Handle(Me, txtH2oSurftensmaximumT(Index), flag_ok)

End Sub

Private Sub txtH2OSurfTensminimimumT_KeyPress(Index As Integer, keyascii As Integer)
    If Index <> 1 Then keyascii = 0
End Sub

Private Sub txtH2OSurfTensminimumT_GotFocus(Index As Integer)
  Call gotfocus_handle(Me, txth2osurftensminimumT(Index), Temp_Text)


End Sub

Private Sub txtH2OSurfTensminimumT_LostFocus(Index As Integer)
Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txth2osurftensminimumT(Index))) Then
     Exit Sub
   End If

   flag_ok = True
  Call LostFocus_Handle(Me, txth2osurftensminimumT(Index), flag_ok)


End Sub

Private Sub txtH2OSurfTensTemperature_GotFocus(Index As Integer)
  Call gotfocus_handle(Me, txth2osurftenstemperature(Index), Temp_Text)
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.WaterSurfaceTension.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If


End Sub

Private Sub txtH2OSurfTensTemperature_KeyPress(Index As Integer, keyascii As Integer)
    
    If keyascii = 13 Then
       keyascii = 0
       cmdok.SetFocus
       Exit Sub
    End If
    Call NumberCheck(keyascii)
   
End Sub

Private Sub txtH2OSurfTensTemperature_LostFocus(Index As Integer)
    Dim msg As String, response As Integer
    Dim Answer As Integer
    Dim IsError As Integer
    Dim ValueChanged As Integer
    Dim EnglishValue As Double, SIValue As Double
    Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txth2osurftenstemperature(Index))) Then
     Exit Sub
   End If

   flag_ok = True

    If txth2osurftenstemperature(Index).Text = "" Then
       phprop.WaterSurfaceTension.input.temperature = -1E+25
       Call LostFocus_Handle(Me, txth2osurftenstemperature(Index), flag_ok)
       Exit Sub
    End If

    Call TextHandleError(IsError, txth2osurftenstemperature(Index), Temp_Text)
    If Not IsError Then
       If Not HaveNumber(CDbl(txth2osurftenstemperature(Index).Text)) Then
          txth2osurftenstemperature(Index).Text = Temp_Text
          txth2osurftenstemperature(Index).SetFocus
           Call LostFocus_Handle(Me, txth2osurftenstemperature(Index), flag_ok)
          Exit Sub
       End If

       Call TextNumberChanged(ValueChanged, txth2osurftenstemperature(Index), Temp_Text)

       If ValueChanged Then
          If CurrentUnits = SIUnits Then
             phprop.WaterSurfaceTension.input.temperature = CDbl(txth2osurftenstemperature(Index).Text)
          Else
             EnglishValue = CDbl(txth2osurftenstemperature(Index).Text)
             Call TEMPENSI(SIValue, EnglishValue)
             phprop.WaterSurfaceTension.input.temperature = SIValue
          End If

       Else
       Call LostFocus_Handle(Me, txth2osurftenstemperature(Index), flag_ok)
          Exit Sub
       End If

    End If
       Call LostFocus_Handle(Me, txth2osurftenstemperature(Index), flag_ok)

End Sub

Private Sub txtWaterSurfaceTensionValue_GotFocus(Index As Integer)
  Call gotfocus_handle(Me, txtwatersurfacetensionvalue(Index), Temp_Text)
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.WaterSurfaceTension.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

End Sub

Private Sub txtWaterSurfaceTensionValue_KeyPress(Index As Integer, keyascii As Integer)
    
    If keyascii = 13 Then
       keyascii = 0
       txth2osurftenstemperature(Index).SetFocus
       Exit Sub
    End If
    Call NumberCheck(keyascii)

End Sub

Private Sub txtWaterSurfaceTensionValue_LostFocus(Index As Integer)
    Dim msg As String, response As Integer
    Dim Answer As Integer
    Dim IsError As Integer
    Dim ValueChanged As Integer
    Dim EnglishValue As Double, SIValue As Double
Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtwatersurfacetensionvalue(Index))) Then
     Exit Sub
   End If

   flag_ok = True

    If txtwatersurfacetensionvalue(Index).Text = "" Then
       phprop.WaterSurfaceTension.input.Value = -1#
       PROPAVAILABLE(WATER_SURF_TENSION_INPUT) = False
       Call LostFocus_Handle(Me, txtwatersurfacetensionvalue(Index), flag_ok)
       Exit Sub
    End If

    Call TextHandleError(IsError, txtwatersurfacetensionvalue(Index), Temp_Text)
    If Not IsError Then
       If Not HaveNumber(CDbl(txtwatersurfacetensionvalue(Index).Text)) Then
          txtwatersurfacetensionvalue(Index).Text = Temp_Text
          txtwatersurfacetensionvalue(Index).SetFocus
          Call LostFocus_Handle(Me, txtwatersurfacetensionvalue(Index), flag_ok)
          Exit Sub
       End If

       Call TextNumberChanged(ValueChanged, txtwatersurfacetensionvalue(Index), Temp_Text)

       If ValueChanged Then
          If CurrentUnits = SIUnits Then
             phprop.WaterSurfaceTension.input.Value = CDbl(txtwatersurfacetensionvalue(Index).Text)
          Else
             EnglishValue = CDbl(txtwatersurfacetensionvalue(Index).Text)
             Call WSTENSI(SIValue, EnglishValue)
             phprop.WaterSurfaceTension.input.Value = SIValue
          End If

          PROPAVAILABLE(WATER_SURF_TENSION_INPUT) = True
       Else
           Call LostFocus_Handle(Me, txtwatersurfacetensionvalue(Index), flag_ok)
           Exit Sub
       End If
    End If
       Call LostFocus_Handle(Me, txtwatersurfacetensionvalue(Index), flag_ok)

End Sub

