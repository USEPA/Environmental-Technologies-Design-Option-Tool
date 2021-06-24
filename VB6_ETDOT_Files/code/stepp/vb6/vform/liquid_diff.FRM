VERSION 5.00
Begin VB.Form liquid_diff_form 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Liquid Diffusivity"
   ClientHeight    =   4515
   ClientLeft      =   480
   ClientTop       =   1905
   ClientWidth     =   8490
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
   ScaleHeight     =   4515
   ScaleWidth      =   8490
   Begin VB.TextBox txtLiqDiffTemperature 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   6120
      TabIndex        =   17
      Top             =   3840
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   1080
      TabIndex        =   14
      Top             =   3840
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   10
      Top             =   3360
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   1080
      TabIndex        =   6
      Top             =   2880
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   2
      Top             =   2400
      Width           =   255
   End
   Begin VB.TextBox txtLiquidDiffusivityValue 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   4080
      TabIndex        =   16
      Top             =   3840
      Width           =   1815
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
      Caption         =   "Accept Selected Liquid Diffusivity"
      Height          =   615
      Left            =   4920
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   2652
      Left            =   960
      Top             =   1680
      Width           =   6492
   End
   Begin VB.Label lblLiqDiffTemperature 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   2
      Left            =   6120
      TabIndex        =   13
      Top             =   3360
      Width           =   1212
   End
   Begin VB.Label lblLiqDiffTemperature 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   6120
      TabIndex        =   9
      Top             =   2880
      Width           =   1212
   End
   Begin VB.Label lblLiqDiffTemperature 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   6120
      TabIndex        =   5
      Top             =   2400
      Width           =   1212
   End
   Begin VB.Label lblLiquidDiffusivityValue 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   2
      Left            =   4080
      TabIndex        =   12
      Top             =   3360
      Width           =   1812
   End
   Begin VB.Label lblLiquidDiffusivityValue 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   4080
      TabIndex        =   8
      Top             =   2880
      Width           =   1812
   End
   Begin VB.Label lblLiquidDiffusivityValue 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   4080
      TabIndex        =   4
      Top             =   2400
      Width           =   1812
   End
   Begin VB.Line Line2 
      X1              =   960
      X2              =   7440
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label lblSource 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "User Input"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   1440
      TabIndex        =   15
      Top             =   3840
      Width           =   2415
   End
   Begin VB.Label lblSource 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Wilke-Chang"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   1440
      TabIndex        =   11
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Label lblSource 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Polson, 1950"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   7
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label lblSource 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hayduk and Laudie"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   1440
      TabIndex        =   3
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label lblLiqDiffTempLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Temp."
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6120
      TabIndex        =   25
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Source"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1440
      TabIndex        =   24
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label lblLiqDiffLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Liquid Diff."
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4080
      TabIndex        =   23
      Top             =   1800
      Width           =   1815
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
      TabIndex        =   22
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
      TabIndex        =   21
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
      TabIndex        =   20
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
      TabIndex        =   19
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Current Liquid Diffusivity Information"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "liquid_diff_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PastLiqDiffInputValue As String
Dim PastLiqDiffInputTemp As String

Private Sub cmdCancel_Click()
    Dim SelectedOption As Integer   'Option selected permanently by the user (the option corresponding to the number on the main screen)

    Select Case phprop.LiquidDiffusivity.CurrentSelection.choice
       Case LIQUID_DIFFUSIVITY_HAYDUKLAUDIE
          SelectedOption = 1
       Case LIQUID_DIFFUSIVITY_POLSON
          SelectedOption = 2
       Case LIQUID_DIFFUSIVITY_WILKECHANG
          SelectedOption = 3
       Case LIQUID_DIFFUSIVITY_INPUT
          SelectedOption = 4
          txtLiquidDiffusivityvalue(3).Text = PastLiqDiffInputValue
          txtliqdifftemperature(3).Text = PastLiqDiffInputTemp
       Case Else
          liquid_diff_form.Hide
          Exit Sub
    End Select

    If Not Option1(SelectedOption).Value Then Option1(SelectedOption).Value = True

    liquid_diff_form.Hide
End Sub

Private Sub cmdok_Click()
    Dim OptionSelected As Integer
    Dim ValueToDisplayIndex As Integer
    Dim i As Integer

'*** Pass new selected value back to main screen
    For i = 1 To 4
        If Option1(i).Value Then
           OptionSelected = i
           Exit For
        End If
    Next i

    Select Case OptionSelected
       Case 1
          ValueToDisplayIndex = LIQUID_DIFFUSIVITY_HAYDUKLAUDIE
       Case 2
          ValueToDisplayIndex = LIQUID_DIFFUSIVITY_POLSON
       Case 3
          ValueToDisplayIndex = LIQUID_DIFFUSIVITY_WILKECHANG
       Case 4
          If Not PROPAVAILABLE(LIQUID_DIFFUSIVITY_INPUT) Then
             MsgBox "User Input can not be selected without first entering a value", MB_ICONSTOP, "Error"
             txtLiquidDiffusivityvalue(3).SetFocus
             Exit Sub
          End If
          ValueToDisplayIndex = LIQUID_DIFFUSIVITY_INPUT
    End Select

    If ValueToDisplayIndex <> phprop.LiquidDiffusivity.CurrentSelection.choice Then
       phprop.LiquidDiffusivity.CurrentSelection.choice = ValueToDisplayIndex
       Call DisplayLiquidDiffusivityMainScreen(ValueToDisplayIndex)
    ElseIf ValueToDisplayIndex = LIQUID_DIFFUSIVITY_INPUT Then
       Call DisplayLiquidDiffusivityMainScreen(ValueToDisplayIndex)
    End If

    liquid_diff_form.Hide

End Sub

Private Sub Form_Activate()
    
  Call centerform_relative(contam_prop_form, Me)
    
    PastLiqDiffInputValue = txtLiquidDiffusivityvalue(3).Text
    PastLiqDiffInputTemp = txtliqdifftemperature(3).Text
End Sub

Private Sub Form_Load()

  Call centerform_relative(contam_prop_form, Me)

    If (DemoMode) Then cmdok.Enabled = False


End Sub

Private Sub lblLiqDiffTemperature_Click(Index As Integer)
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.LiquidDiffusivity.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

End Sub

Private Sub lblLiquidDiffusivityValue_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ErrorFlag As Integer
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.LiquidDiffusivity.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

    If Button <> 2 Then Exit Sub

    Select Case Index
       Case 0  'Source = Hayduk & Laudie Correlation
          ErrorFlag = phprop.LiquidDiffusivity.haydukLaudie.error
       Case 1  'Source = Polson Method
          ErrorFlag = phprop.LiquidDiffusivity.polson.error
       Case 2  'Source = Wilke-Chang Correlation
          ErrorFlag = phprop.LiquidDiffusivity.wilkeChang.error
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
       If Index = hilight.LiquidDiffusivity.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

End Sub

Private Sub Option1_Click(Index As Integer)
    Dim i As Integer, SourceIndex As Integer

    SourceIndex = Index - 1
    If SourceIndex = hilight.LiquidDiffusivity.PreviousIndex Then Exit Sub
    lblSource(SourceIndex).BackColor = &H800000
    lblSource(SourceIndex).ForeColor = &H80000005
    i = hilight.LiquidDiffusivity.PreviousIndex
    hilight.LiquidDiffusivity.PreviousIndex = SourceIndex
    If i = -1 Then Exit Sub
    If Option1(i + 1).Enabled = False Then Exit Sub
    lblSource(i).BackColor = &H80000005
    lblSource(i).ForeColor = &H80000008

End Sub

Private Sub txtLiqDiffTemperature_GotFocus(Index As Integer)

  Call gotfocus_handle(Me, txtliqdifftemperature(Index), Temp_Text)
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.LiquidDiffusivity.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

End Sub

Private Sub txtLiqDiffTemperature_KeyPress(Index As Integer, keyascii As Integer)
    
    If keyascii = 13 Then
       keyascii = 0
       cmdok.SetFocus
       Exit Sub
    End If
    Call NumberCheck(keyascii)

End Sub

Private Sub txtLiqDiffTemperature_LostFocus(Index As Integer)
    Dim msg As String, response As Integer
    Dim Answer As Integer
    Dim IsError As Integer
    Dim ValueChanged As Integer
    Dim EnglishValue As Double, SIValue As Double
    Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtliqdifftemperature(Index))) Then
     Exit Sub
   End If

   flag_ok = True

    If txtliqdifftemperature(Index).Text = "" Then
       phprop.LiquidDiffusivity.input.temperature = -1E+25
        Call LostFocus_Handle(Me, txtliqdifftemperature(Index), flag_ok)
       Exit Sub
    End If

    Call TextHandleError(IsError, txtliqdifftemperature(Index), Temp_Text)
    If Not IsError Then
       If Not HaveNumber(CDbl(txtliqdifftemperature(Index).Text)) Then
          txtliqdifftemperature(Index).Text = Temp_Text
          txtliqdifftemperature(Index).SetFocus
        Call LostFocus_Handle(Me, txtliqdifftemperature(Index), flag_ok)
          Exit Sub
       End If

       Call TextNumberChanged(ValueChanged, txtliqdifftemperature(Index), Temp_Text)

       If ValueChanged Then
          If CurrentUnits = SIUnits Then
             phprop.LiquidDiffusivity.input.temperature = CDbl(txtliqdifftemperature(Index).Text)
          Else
             EnglishValue = CDbl(txtliqdifftemperature(Index).Text)
             Call TEMPENSI(SIValue, EnglishValue)
             phprop.LiquidDiffusivity.input.temperature = SIValue
          End If
          
       Else
        Call LostFocus_Handle(Me, txtliqdifftemperature(Index), flag_ok)
          Exit Sub
       End If

    End If
        Call LostFocus_Handle(Me, txtliqdifftemperature(Index), flag_ok)

End Sub

Private Sub txtLiquidDiffusivityValue_GotFocus(Index As Integer)
  Call gotfocus_handle(Me, txtLiquidDiffusivityvalue(Index), Temp_Text)
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.LiquidDiffusivity.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

End Sub

Private Sub txtLiquidDiffusivityValue_KeyPress(Index As Integer, keyascii As Integer)
    
    If keyascii = 13 Then
       keyascii = 0
       txtliqdifftemperature(Index).SetFocus
       Exit Sub
    End If
    Call NumberCheck(keyascii)
   
End Sub

Private Sub txtLiquidDiffusivityValue_LostFocus(Index As Integer)
    Dim msg As String, response As Integer
    Dim Answer As Integer
    Dim IsError As Integer
    Dim ValueChanged As Integer
    Dim EnglishValue As Double, SIValue As Double
Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtLiquidDiffusivityvalue(Index))) Then
     Exit Sub
   End If

   flag_ok = True

    If txtLiquidDiffusivityvalue(Index).Text = "" Then
       phprop.LiquidDiffusivity.input.Value = -1#
       PROPAVAILABLE(LIQUID_DIFFUSIVITY_INPUT) = False
       Call LostFocus_Handle(Me, txtLiquidDiffusivityvalue(Index), flag_ok)
       Exit Sub
    End If

    Call TextHandleError(IsError, txtLiquidDiffusivityvalue(Index), Temp_Text)
    If Not IsError Then
       If Not HaveNumber(CDbl(txtLiquidDiffusivityvalue(Index).Text)) Then
          txtLiquidDiffusivityvalue(Index).Text = Temp_Text
          txtLiquidDiffusivityvalue(Index).SetFocus
       Call LostFocus_Handle(Me, txtLiquidDiffusivityvalue(Index), flag_ok)
          Exit Sub
       End If

       Call TextNumberChanged(ValueChanged, txtLiquidDiffusivityvalue(Index), Temp_Text)

       If ValueChanged Then
          If CurrentUnits = SIUnits Then
             phprop.LiquidDiffusivity.input.Value = CDbl(txtLiquidDiffusivityvalue(Index).Text)
          Else
             EnglishValue = CDbl(txtLiquidDiffusivityvalue(Index).Text)
             Call LDIFENSI(SIValue, EnglishValue)
             phprop.LiquidDiffusivity.input.Value = SIValue
          End If
         
          PROPAVAILABLE(LIQUID_DIFFUSIVITY_INPUT) = True
       Else
          Call LostFocus_Handle(Me, txtLiquidDiffusivityvalue(Index), flag_ok)
          Exit Sub
       End If
    End If
       Call LostFocus_Handle(Me, txtLiquidDiffusivityvalue(Index), flag_ok)

End Sub

