VERSION 5.00
Begin VB.Form frmAirDensity 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Air Density"
   ClientHeight    =   3585
   ClientLeft      =   675
   ClientTop       =   1530
   ClientWidth     =   8505
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
   ScaleHeight     =   3585
   ScaleWidth      =   8505
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
   Begin VB.TextBox txtAirDensitymaximumT 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   7320
      TabIndex        =   7
      Text            =   "not visible"
      Top             =   2880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtAirDensityminimumT 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   6120
      TabIndex        =   6
      Text            =   "not visible"
      Top             =   2880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtAirDensityTemperature 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   4920
      TabIndex        =   5
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox txtAirDensityValue 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   3240
      TabIndex        =   4
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      Caption         =   "Accept Selected Air Density"
      Height          =   615
      Left            =   4860
      TabIndex        =   0
      Top             =   120
      Width           =   3495
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
   Begin VB.Label lblAirDensitymaximumT 
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
   Begin VB.Label lblAirDensityminimumT 
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
   Begin VB.Label lblAirDensityTemperature 
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
   Begin VB.Label lblAirDensityValue 
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
      Caption         =   "Ideal Gas Law"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   18
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Current Air Density Information"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   240
      Width           =   4095
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
      TabIndex        =   16
      Top             =   720
      Width           =   1575
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
      TabIndex        =   15
      Top             =   1080
      Width           =   1575
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
      TabIndex        =   14
      Top             =   1080
      Width           =   2295
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
      TabIndex        =   13
      Top             =   720
      Width           =   2295
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   2
      Height          =   1335
      Left            =   120
      Top             =   120
      Width           =   4335
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4440
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label lblAirDensminTLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tmin"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6120
      TabIndex        =   12
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label lblAirDensmaxTLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tmax"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7320
      TabIndex        =   11
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label lblAirDensTempLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Temp."
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4920
      TabIndex        =   10
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label lblAirDensLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Air Dens."
      ForeColor       =   &H00000000&
      Height          =   400
      Left            =   3240
      TabIndex        =   9
      Top             =   1750
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Source"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   1695
      Left            =   120
      Top             =   1680
      Width           =   8295
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   8400
      Y1              =   2160
      Y2              =   2160
   End
End
Attribute VB_Name = "frmAirDensity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PastAirDensInputValue As String
Dim PastAirDensInputTemp As String

Private Sub cmdCancel_Click()
    Dim SelectedOption As Integer   'Option selected permanently by the user (the option corresponding to the number on the main screen)

    Select Case phprop.AirDensity.CurrentSelection.choice
       Case AIR_DENSITY_CORRELATION
          SelectedOption = 1
       Case AIR_DENSITY_INPUT
          SelectedOption = 2
          txtairdensityvalue(1).Text = PastAirDensInputValue
          txtairdensitytemperature(1).Text = PastAirDensInputTemp
       Case Else
          frmAirDensity.Hide
          Exit Sub
    End Select

    If Not Option1(SelectedOption).Value Then Option1(SelectedOption).Value = True

    frmAirDensity.Hide
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
          ValueToDisplayIndex = AIR_DENSITY_CORRELATION
       Case 2
          If Not PROPAVAILABLE(AIR_DENSITY_INPUT) Then
             MsgBox "User Input can not be selected without first entering a value", MB_ICONSTOP, "Error"
             txtairdensityvalue(1).SetFocus
             Exit Sub
          End If
          ValueToDisplayIndex = AIR_DENSITY_INPUT
    End Select

    If ValueToDisplayIndex <> phprop.AirDensity.CurrentSelection.choice Then
       phprop.AirDensity.CurrentSelection.choice = ValueToDisplayIndex
       Call DisplayAirDensityMainScreen(ValueToDisplayIndex)
    ElseIf ValueToDisplayIndex = AIR_DENSITY_INPUT Then
       Call DisplayAirDensityMainScreen(ValueToDisplayIndex)
    End If

    frmAirDensity.Hide
End Sub

Private Sub Form_Activate()
    
  Call centerform_relative(contam_prop_form, Me)
    
    PastAirDensInputValue = txtairdensityvalue(1).Text
    PastAirDensInputTemp = txtairdensitytemperature(1).Text
End Sub

Private Sub Form_Load()

    If (DemoMode) Then cmdok.Enabled = False

  Call centerform_relative(contam_prop_form, Me)

End Sub

Private Sub lblAirDensitymaximumT_Click(Index As Integer)
    
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.AirDensity.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

End Sub

Private Sub lblAirDensityminimumT_Click(Index As Integer)
    
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.AirDensity.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

End Sub

Private Sub lblAirDensityTemperature_Click(Index As Integer)
    
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.AirDensity.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

End Sub

Private Sub lblAirDensityValue_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ErrorFlag As Integer
    
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.AirDensity.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

    If Button <> 2 Then Exit Sub
    
    Select Case Index
       Case 0  'Ideal Gas Law
          ErrorFlag = phprop.AirDensity.correlation.error
    End Select

    If ErrorFlag = 0 Then Exit Sub

    If ErrorFlag < 0 Then
       MsgBox ErrorMsg(ErrorFlag), MB_ICONINFORMATION, "Air - Data Unavailable"
    End If

    If ErrorFlag > 0 Then
       MsgBox ErrorMsg(ErrorFlag), MB_ICONINFORMATION, "Air - Warning"
    End If

End Sub

Private Sub lblSource_Click(Index As Integer)
    
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.AirDensity.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

End Sub

Private Sub Option1_Click(Index As Integer)
    Dim i As Integer, SourceIndex As Integer

    SourceIndex = Index - 1
    If SourceIndex = hilight.AirDensity.PreviousIndex Then Exit Sub
    lblSource(SourceIndex).BackColor = &H800000
    lblSource(SourceIndex).ForeColor = &H80000005
    i = hilight.AirDensity.PreviousIndex
    hilight.AirDensity.PreviousIndex = SourceIndex
    If i = -1 Then Exit Sub
    If Option1(i + 1).Enabled = False Then Exit Sub
    lblSource(i).BackColor = &H80000005
    lblSource(i).ForeColor = &H80000008

End Sub

Private Sub txtAirDensitymaximumT_GotFocus(Index As Integer)
  Call gotfocus_handle(Me, txtairdensitymaximumT(Index), Temp_Text)

End Sub

Private Sub txtAirDensitymaximumT_KeyPress(Index As Integer, keyascii As Integer)
    If Index <> 1 Then keyascii = 0
End Sub

Private Sub txtAirDensitymaximumT_LostFocus(Index As Integer)
Dim ConversionFactor As Double
Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtairdensitymaximumT(Index))) Then
     Exit Sub
   End If

   flag_ok = True
  Call LostFocus_Handle(Me, txtairdensitymaximumT(Index), flag_ok)

End Sub

Private Sub txtAirDensityminimumT_GotFocus(Index As Integer)
  Call gotfocus_handle(Me, txtairdensityminimumT(Index), Temp_Text)

End Sub

Private Sub txtAirDensityminimumT_LostFocus(Index As Integer)
Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtairdensityminimumT(Index))) Then
     Exit Sub
   End If

   flag_ok = True
  Call LostFocus_Handle(Me, txtairdensityminimumT(Index), flag_ok)

End Sub

Private Sub txtAirDensityTemperature_GotFocus(Index As Integer)
  Call gotfocus_handle(Me, txtairdensitytemperature(Index), Temp_Text)

    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.AirDensity.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If
End Sub

Private Sub txtAirDensityTemperature_KeyPress(Index As Integer, keyascii As Integer)
    
    If keyascii = 13 Then
       keyascii = 0
       cmdok.SetFocus
       Exit Sub
    End If
    Call NumberCheck(keyascii)
   
End Sub

Private Sub txtAirDensityTemperature_LostFocus(Index As Integer)
    Dim msg As String, response As Integer
    Dim Answer As Integer
    Dim IsError As Integer
    Dim ValueChanged As Integer
    Dim EnglishValue As Double, SIValue As Double
    Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtairdensitytemperature(Index))) Then
     Exit Sub
   End If

   flag_ok = True

    If txtairdensitytemperature(Index).Text = "" Then
       phprop.AirDensity.input.temperature = -1E+25
       Call LostFocus_Handle(Me, txtairdensitytemperature(Index), flag_ok)
       Exit Sub
    End If

    Call TextHandleError(IsError, txtairdensitytemperature(Index), Temp_Text)
    If Not IsError Then
       If Not HaveNumber(CDbl(txtairdensitytemperature(Index).Text)) Then
          txtairdensitytemperature(Index).Text = Temp_Text
          txtairdensitytemperature(Index).SetFocus
          Call LostFocus_Handle(Me, txtairdensitytemperature(Index), flag_ok)
          Exit Sub
       End If

       Call TextNumberChanged(ValueChanged, txtairdensitytemperature(Index), Temp_Text)

       If ValueChanged Then
          If CurrentUnits = SIUnits Then
             phprop.AirDensity.input.temperature = CDbl(txtairdensitytemperature(Index).Text)
          Else
             EnglishValue = CDbl(txtairdensitytemperature(Index).Text)
             Call TEMPENSI(SIValue, EnglishValue)
             phprop.AirDensity.input.temperature = SIValue
          End If

          
       Else
          Call LostFocus_Handle(Me, txtairdensitytemperature(Index), flag_ok)
          Exit Sub
       End If

    End If
  Call LostFocus_Handle(Me, txtairdensitytemperature(Index), flag_ok)

End Sub

Private Sub txtAirDensityValue_GotFocus(Index As Integer)
  Call gotfocus_handle(Me, txtairdensityvalue(Index), Temp_Text)

    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.AirDensity.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If
       
End Sub

Private Sub txtAirDensityValue_KeyPress(Index As Integer, keyascii As Integer)
    
    If keyascii = 13 Then
       keyascii = 0
       txtairdensitytemperature(Index).SetFocus
       Exit Sub
    End If
    Call NumberCheck(keyascii)
   
End Sub

Private Sub txtAirDensityValue_LostFocus(Index As Integer)
    Dim msg As String, response As Integer
    Dim Answer As Integer
    Dim IsError As Integer
    Dim ValueChanged As Integer
    Dim EnglishValue As Double, SIValue As Double
    Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtairdensityvalue(Index))) Then
     Exit Sub
   End If

   flag_ok = True

    If txtairdensityvalue(Index).Text = "" Then
       phprop.AirDensity.input.Value = -1#
       PROPAVAILABLE(AIR_DENSITY_INPUT) = False
       Call LostFocus_Handle(Me, txtairdensityvalue(Index), flag_ok)
       Exit Sub
    End If

    Call TextHandleError(IsError, txtairdensityvalue(Index), Temp_Text)
    If Not IsError Then
       If Not HaveNumber(CDbl(txtairdensityvalue(Index).Text)) Then
          txtairdensityvalue(Index).Text = Temp_Text
          Call LostFocus_Handle(Me, txtairdensityvalue(Index), flag_ok)
          Exit Sub
       End If

       Call TextNumberChanged(ValueChanged, txtairdensityvalue(Index), Temp_Text)

       If ValueChanged Then
          If CurrentUnits = SIUnits Then
             phprop.AirDensity.input.Value = CDbl(txtairdensityvalue(Index).Text)
          Else
             EnglishValue = CDbl(txtairdensityvalue(Index).Text)
             Call ADENENSI(SIValue, EnglishValue)
             phprop.AirDensity.input.Value = SIValue
          End If
          
          PROPAVAILABLE(AIR_DENSITY_INPUT) = True
       Else
          Call LostFocus_Handle(Me, txtairdensityvalue(Index), flag_ok)
          Exit Sub
       End If
    End If
'    txtAirDensityValue(Index).SetFocus
  Call LostFocus_Handle(Me, txtairdensityvalue(Index), flag_ok)

End Sub

