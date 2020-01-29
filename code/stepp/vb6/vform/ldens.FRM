VERSION 5.00
Begin VB.Form ldens_form 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Liquid Density"
   ClientHeight    =   4410
   ClientLeft      =   1035
   ClientTop       =   1515
   ClientWidth     =   8565
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
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4410
   ScaleWidth      =   8565
   Begin VB.TextBox txtLDmaximumT 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   7320
      TabIndex        =   19
      Text            =   "No Visible"
      Top             =   3360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtLDminimumT 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   6240
      TabIndex        =   18
      Text            =   "No Visible"
      Top             =   3360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtLDTemperature 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   5040
      TabIndex        =   17
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      Caption         =   "Accept Selected Liquid Density"
      Height          =   615
      Left            =   4920
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
   Begin VB.TextBox txtLiquidDensityValue 
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   2
      Left            =   3240
      TabIndex        =   16
      Top             =   3360
      Width           =   1572
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   2
      Left            =   240
      TabIndex        =   8
      Top             =   2880
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   3
      Left            =   240
      TabIndex        =   14
      Top             =   3360
      Width           =   252
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   252
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   $"ldens.frx":0000
      ForeColor       =   &H80000008&
      Height          =   400
      Left            =   120
      TabIndex        =   30
      Top             =   3960
      Width           =   8295
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   2172
      Left            =   120
      Top             =   1680
      Width           =   8292
   End
   Begin VB.Label lblLDmaximumT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "not visible"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   7320
      TabIndex        =   13
      Top             =   2880
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label lblLDmaximumT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   7320
      TabIndex        =   7
      Top             =   2400
      Width           =   972
   End
   Begin VB.Label lblLDminimumT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "not visible"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   6240
      TabIndex        =   12
      Top             =   2880
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label lblLDminimumT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   6240
      TabIndex        =   6
      Top             =   2400
      Width           =   972
   End
   Begin VB.Label lblLDTemperature 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   5040
      TabIndex        =   11
      Top             =   2880
      Width           =   972
   End
   Begin VB.Label lblLDTemperature 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   5040
      TabIndex        =   5
      Top             =   2400
      Width           =   972
   End
   Begin VB.Label lblLiquidDensityValue 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   3180
      TabIndex        =   10
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label lblLiquidDensityValue 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   3240
      TabIndex        =   4
      Top             =   2400
      Width           =   1572
   End
   Begin VB.Label lblLDensminTLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tmin"
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   6240
      TabIndex        =   29
      Top             =   1800
      Width           =   972
   End
   Begin VB.Label lblLDensmaxTLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tmax"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7320
      TabIndex        =   28
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label lblLDensTempLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Temp."
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   5040
      TabIndex        =   27
      Top             =   1800
      Width           =   972
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Current Liquid Density Information"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   26
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
      TabIndex        =   25
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
      TabIndex        =   24
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
      TabIndex        =   23
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
      TabIndex        =   22
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
   Begin VB.Label lblLDensLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Liq. Dens."
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   3240
      TabIndex        =   21
      Top             =   1800
      Width           =   1572
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Source"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   600
      TabIndex        =   20
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label lblSource 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Database"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   3
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label lblSource 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Group Contribution Method"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   9
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label lblSource 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "User Input"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   600
      TabIndex        =   15
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   8400
      Y1              =   2160
      Y2              =   2160
   End
End
Attribute VB_Name = "ldens_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PastLDensInputValue As String
Dim PastLDensInputTemp As String

Private Sub cmdCancel_Click()
    Dim SelectedOption As Integer   'Option selected permanently by the user (the option corresponding to the number on the main screen)

    Select Case phprop.LiquidDensity.CurrentSelection.choice
       Case LIQUID_DENSITY_DATABASE
          SelectedOption = 1
       Case LIQUID_DENSITY_UNIFAC
          SelectedOption = 2
       Case LIQUID_DENSITY_INPUT
          SelectedOption = 3
          txtliquiddensityvalue(2).Text = PastLDensInputValue
          txtldtemperature(2).Text = PastLDensInputTemp
       Case Else
          ldens_form.Hide
          Exit Sub
    End Select

    If Not Option1(SelectedOption).Value Then Option1(SelectedOption).Value = True

    ldens_form.Hide
End Sub

Private Sub cmdok_Click()
    Dim OptionSelected As Integer
    Dim ValueToDisplayIndex As Integer
    Dim i As Integer

'*** Pass new selected value back to main screen
    For i = 1 To 3
        If Option1(i).Value Then
           OptionSelected = i
           Exit For
        End If
    Next i

    Select Case OptionSelected
       Case 1
          ValueToDisplayIndex = LIQUID_DENSITY_DATABASE
       Case 2
          ValueToDisplayIndex = LIQUID_DENSITY_UNIFAC
       Case 3
          If Not PROPAVAILABLE(LIQUID_DENSITY_INPUT) Then
             MsgBox "User Input can not be selected without first entering a value", MB_ICONSTOP, "Error"
             txtliquiddensityvalue(2).SetFocus
             Exit Sub
          End If
          ValueToDisplayIndex = LIQUID_DENSITY_INPUT
    End Select

    If ValueToDisplayIndex <> phprop.LiquidDensity.CurrentSelection.choice Then
       phprop.LiquidDensity.CurrentSelection.choice = ValueToDisplayIndex
       Call DisplayLiquidDensityMainScreen(ValueToDisplayIndex)
    ElseIf ValueToDisplayIndex = LIQUID_DENSITY_INPUT Then
       Call DisplayLiquidDensityMainScreen(ValueToDisplayIndex)
    End If

    ldens_form.Hide

End Sub

Private Sub Form_Activate()
    
  Call centerform_relative(contam_prop_form, Me)
    
    PastLDensInputValue = txtliquiddensityvalue(2).Text
    PastLDensInputTemp = txtldtemperature(2).Text
End Sub

Private Sub Form_Load()

  Call centerform_relative(contam_prop_form, Me)

    If (DemoMode) Then cmdok.Enabled = False


End Sub

Private Sub lblLDmaximumT_Click(Index As Integer)
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.LiquidDensity.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

End Sub

Private Sub lblLDminimumT_Click(Index As Integer)
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.LiquidDensity.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

End Sub

Private Sub lblLDTemperature_Click(Index As Integer)
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.LiquidDensity.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

End Sub

Private Sub lblLiquidDensityValue_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ErrorFlag As Integer
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.LiquidDensity.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

    If Button <> 2 Then Exit Sub

    Select Case Index
       Case 0  'Source = Database
          ErrorFlag = phprop.LiquidDensity.database.error
       Case 1  'Source = Group Contribution Method
          ErrorFlag = phprop.LiquidDensity.UNIFAC.error
    End Select

    If ErrorFlag = 0 Then Exit Sub

    If ErrorFlag < 0 Then
       MsgBox ErrorMsg(ErrorFlag), MB_ICONINFORMATION, Trim$(phprop.Name) & " - Data Unavailable"
    End If

    If ErrorFlag > 0 Then
       MsgBox ErrorMsg(ErrorFlag), MB_ICONINFORMATION, Trim$(phprop.Name) & " - Warning"
    End If

End Sub

Private Sub lblSource_Click(Index As Integer)
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.LiquidDensity.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

End Sub

Private Sub Option1_Click(Index As Integer)
    Dim i As Integer, SourceIndex As Integer

    SourceIndex = Index - 1
    If SourceIndex = hilight.LiquidDensity.PreviousIndex Then Exit Sub
    lblSource(SourceIndex).BackColor = &H800000
    lblSource(SourceIndex).ForeColor = &H80000005
    i = hilight.LiquidDensity.PreviousIndex
    hilight.LiquidDensity.PreviousIndex = SourceIndex
    If i = -1 Then Exit Sub
    If Option1(i + 1).Enabled = False Then Exit Sub
    lblSource(i).BackColor = &H80000005
    lblSource(i).ForeColor = &H80000008

End Sub

Private Sub txtLDmaximumT_GotFocus(Index As Integer)
  Call gotfocus_handle(Me, txtldmaximumT(Index), Temp_Text)

End Sub

Private Sub txtLDmaximumT_KeyPress(Index As Integer, keyascii As Integer)
    If Index <> 2 Then keyascii = 0
End Sub

Private Sub txtLDmaximumT_LostFocus(Index As Integer)
Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtldmaximumT(Index))) Then
     Exit Sub
   End If

   flag_ok = True
  Call LostFocus_Handle(Me, txtldmaximumT(Index), flag_ok)

End Sub

Private Sub txtLDminimimumT_KeyPress(Index As Integer, keyascii As Integer)
    If Index <> 2 Then keyascii = 0
End Sub

Private Sub txtLDminimumT_GotFocus(Index As Integer)
  Call gotfocus_handle(Me, txtldminimumT(Index), Temp_Text)

End Sub

Private Sub txtLDminimumT_LostFocus(Index As Integer)
Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtldminimumT(Index))) Then
     Exit Sub
   End If

   flag_ok = True
  Call LostFocus_Handle(Me, txtldminimumT(Index), flag_ok)

End Sub

Private Sub txtLDTemperature_GotFocus(Index As Integer)
  Call gotfocus_handle(Me, txtldtemperature(Index), Temp_Text)
    
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.LiquidDensity.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If
End Sub

Private Sub txtLDTemperature_KeyPress(Index As Integer, keyascii As Integer)
    
    If keyascii = 13 Then
       keyascii = 0
       cmdok.SetFocus
       Exit Sub
    End If
    Call NumberCheck(keyascii)
   
End Sub

Private Sub txtLDTemperature_LostFocus(Index As Integer)
    Dim msg As String, response As Integer
    Dim Answer As Integer
    Dim IsError As Integer
    Dim ValueChanged As Integer
    Dim EnglishValue As Double, SIValue As Double
    Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtldtemperature(Index))) Then
     Exit Sub
   End If

   flag_ok = True

    If txtldtemperature(Index).Text = "" Then
       phprop.LiquidDensity.input.temperature = -1E+25
       Call LostFocus_Handle(Me, txtldtemperature(Index), flag_ok)
       Exit Sub
    End If

    Call TextHandleError(IsError, txtldtemperature(Index), Temp_Text)
    If Not IsError Then
       If Not HaveNumber(CDbl(txtldtemperature(Index).Text)) Then
          txtldtemperature(Index).Text = Temp_Text
          txtldtemperature(Index).SetFocus
          Call LostFocus_Handle(Me, txtldtemperature(Index), flag_ok)
          Exit Sub
       End If

       Call TextNumberChanged(ValueChanged, txtldtemperature(Index), Temp_Text)

       If ValueChanged Then
          If CurrentUnits = SIUnits Then
             phprop.LiquidDensity.input.temperature = CDbl(txtldtemperature(Index).Text)
          Else
             EnglishValue = CDbl(txtldtemperature(Index).Text)
             Call TEMPENSI(SIValue, EnglishValue)
             phprop.LiquidDensity.input.temperature = SIValue
          End If
          
       Else
          Call LostFocus_Handle(Me, txtldtemperature(Index), flag_ok)
          Exit Sub
       End If

    End If
  Call LostFocus_Handle(Me, txtldtemperature(Index), flag_ok)

End Sub

Private Sub txtLiquidDensityValue_GotFocus(Index As Integer)
  Call gotfocus_handle(Me, txtliquiddensityvalue(Index), Temp_Text)
    
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.LiquidDensity.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If
End Sub

Private Sub txtLiquidDensityValue_KeyPress(Index As Integer, keyascii As Integer)
    
    If keyascii = 13 Then
       keyascii = 0
       txtldtemperature(Index).SetFocus
       Exit Sub
    End If
    Call NumberCheck(keyascii)
    
End Sub

Private Sub txtLiquidDensityValue_LostFocus(Index As Integer)
    Dim msg As String, response As Integer
    Dim Answer As Integer
    Dim IsError As Integer
    Dim ValueChanged As Integer
    Dim EnglishValue As Double, SIValue As Double
    Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtliquiddensityvalue(Index))) Then
     Exit Sub
   End If

   flag_ok = True

    If txtliquiddensityvalue(Index).Text = "" Then
       phprop.LiquidDensity.input.Value = -1#
       PROPAVAILABLE(LIQUID_DENSITY_INPUT) = False
       Call LostFocus_Handle(Me, txtliquiddensityvalue(Index), flag_ok)
       Exit Sub
    End If

    Call TextHandleError(IsError, txtliquiddensityvalue(Index), Temp_Text)
    If Not IsError Then
       If Not HaveNumber(CDbl(txtliquiddensityvalue(Index).Text)) Then
          txtliquiddensityvalue(Index).Text = Temp_Text
          txtliquiddensityvalue(Index).SetFocus
          Call LostFocus_Handle(Me, txtliquiddensityvalue(Index), flag_ok)
          Exit Sub
       End If

       Call TextNumberChanged(ValueChanged, txtliquiddensityvalue(Index), Temp_Text)

       If ValueChanged Then
          If CurrentUnits = SIUnits Then
             phprop.LiquidDensity.input.Value = CDbl(txtliquiddensityvalue(Index).Text)
          Else
             EnglishValue = CDbl(txtliquiddensityvalue(Index).Text)
             Call LDENENSI(SIValue, EnglishValue)
             phprop.LiquidDensity.input.Value = SIValue
          End If

          PROPAVAILABLE(LIQUID_DENSITY_INPUT) = True
       Else
          Call LostFocus_Handle(Me, txtliquiddensityvalue(Index), flag_ok)
          Exit Sub
       End If
    End If
       Call LostFocus_Handle(Me, txtliquiddensityvalue(Index), flag_ok)

End Sub

