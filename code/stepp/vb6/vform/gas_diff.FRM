VERSION 5.00
Begin VB.Form gas_diff_form 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gas Diffusivity"
   ClientHeight    =   3570
   ClientLeft      =   2010
   ClientTop       =   1845
   ClientWidth     =   8520
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
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3570
   ScaleWidth      =   8520
   Begin VB.TextBox txtGasDiffTemperature 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   6120
      TabIndex        =   5
      Text            =   " "
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox txtGasDiffusivityValue 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   4080
      TabIndex        =   4
      Text            =   " "
      Top             =   2880
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   1080
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
      Left            =   1080
      TabIndex        =   2
      Top             =   2400
      Width           =   255
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
      Caption         =   "Accept Selected Gas Diffusivity"
      Height          =   615
      Left            =   4920
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label lblGasDiffTemperature 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   6120
      TabIndex        =   17
      Top             =   2400
      Width           =   1212
   End
   Begin VB.Label lblGasDiffusivityValue 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   4080
      TabIndex        =   16
      Top             =   2400
      Width           =   1812
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
      Left            =   1440
      TabIndex        =   15
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label lblSource 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Wilke-Lee"
      ForeColor       =   &H00000000&
      Height          =   252
      Index           =   0
      Left            =   1440
      TabIndex        =   14
      Top             =   2400
      Width           =   2412
   End
   Begin VB.Line Line2 
      X1              =   960
      X2              =   7440
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   1695
      Left            =   960
      Top             =   1680
      Width           =   6495
   End
   Begin VB.Label lblGasDiffTempLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Temp."
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6120
      TabIndex        =   13
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
      TabIndex        =   12
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label lblGasDiffLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Gas Diffusivity"
      ForeColor       =   &H00000000&
      Height          =   400
      Left            =   4080
      TabIndex        =   11
      Top             =   1750
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
      TabIndex        =   10
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
      TabIndex        =   9
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
      TabIndex        =   8
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
      TabIndex        =   7
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Current Gas Diffusivity Information"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "gas_diff_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PastGasDiffInputValue As String
Dim PastGasDiffInputTemp As String

Private Sub cmdCancel_Click()
    Dim SelectedOption As Integer   'Option selected permanently by the user (the option corresponding to the number on the main screen)

    Select Case phprop.GasDiffusivity.CurrentSelection.choice
       Case GAS_DIFFUSIVITY_WILKELEE
          SelectedOption = 1
       Case GAS_DIFFUSIVITY_INPUT
          SelectedOption = 2
          txtgasdiffusivityvalue(1).Text = PastGasDiffInputValue
          txtgasdifftemperature(1).Text = PastGasDiffInputTemp
       Case Else
          gas_diff_form.Hide
          Exit Sub
    End Select

    If Not Option1(SelectedOption).Value Then Option1(SelectedOption).Value = True

    gas_diff_form.Hide
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
          ValueToDisplayIndex = GAS_DIFFUSIVITY_WILKELEE
       Case 2
          If Not PROPAVAILABLE(GAS_DIFFUSIVITY_INPUT) Then
             MsgBox "User Input can not be selected without first entering a value", MB_ICONSTOP, "Error"
             txtgasdiffusivityvalue(1).SetFocus
             Exit Sub
          End If
          ValueToDisplayIndex = GAS_DIFFUSIVITY_INPUT
    End Select

    If ValueToDisplayIndex <> phprop.GasDiffusivity.CurrentSelection.choice Then
       phprop.GasDiffusivity.CurrentSelection.choice = ValueToDisplayIndex
       Call DisplayGasDiffusivityMainScreen(ValueToDisplayIndex)
    ElseIf ValueToDisplayIndex = GAS_DIFFUSIVITY_INPUT Then
       Call DisplayGasDiffusivityMainScreen(ValueToDisplayIndex)
    End If

    gas_diff_form.Hide

End Sub

Private Sub Form_Activate()

  Call centerform_relative(contam_prop_form, Me)

    PastGasDiffInputValue = txtgasdiffusivityvalue(1).Text
    PastGasDiffInputTemp = txtgasdifftemperature(1).Text
End Sub

Private Sub Form_Load()

  Call centerform_relative(contam_prop_form, Me)

    If (DemoMode) Then cmdok.Enabled = False

End Sub

Private Sub lblGasDiffTemperature_Click(Index As Integer)
    
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.GasDiffusivity.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

End Sub

Private Sub lblGasDiffusivityValue_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ErrorFlag As Integer
    
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.GasDiffusivity.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

    If Button <> 2 Then Exit Sub

    Select Case Index
       Case 0  'Source = Wilke-Lee Modification of Hirschfelder-Bird-Spotz Method
          ErrorFlag = phprop.GasDiffusivity.wilkeLee.error
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
       If Index = hilight.GasDiffusivity.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

End Sub

Private Sub Option1_Click(Index As Integer)
    Dim i As Integer, SourceIndex As Integer

    SourceIndex = Index - 1
    If SourceIndex = hilight.GasDiffusivity.PreviousIndex Then Exit Sub
    lblSource(SourceIndex).BackColor = &H800000
    lblSource(SourceIndex).ForeColor = &H80000005
    i = hilight.GasDiffusivity.PreviousIndex
    hilight.GasDiffusivity.PreviousIndex = SourceIndex
    If i = -1 Then Exit Sub
    If Option1(i + 1).Enabled = False Then Exit Sub
    lblSource(i).BackColor = &H80000005
    lblSource(i).ForeColor = &H80000008

End Sub

Private Sub txtGasDiffTemperature_GotFocus(Index As Integer)
  Call gotfocus_handle(Me, txtgasdifftemperature(Index), Temp_Text)
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.GasDiffusivity.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If
End Sub

Private Sub txtGasDiffTemperature_KeyPress(Index As Integer, keyascii As Integer)
    
    If keyascii = 13 Then
       keyascii = 0
       cmdok.SetFocus
       Exit Sub
    End If
    Call NumberCheck(keyascii)
    
End Sub

Private Sub txtGasDiffTemperature_LostFocus(Index As Integer)
    Dim msg As String, response As Integer
    Dim Answer As Integer
    Dim IsError As Integer
    Dim ValueChanged As Integer
    Dim EnglishValue As Double, SIValue As Double
Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtgasdifftemperature(Index))) Then
     Exit Sub
   End If

   flag_ok = True

    If txtgasdifftemperature(Index).Text = "" Then
       phprop.GasDiffusivity.input.temperature = -1E+25
       Call LostFocus_Handle(Me, txtgasdifftemperature(Index), flag_ok)
       Exit Sub
    End If

    Call TextHandleError(IsError, txtgasdifftemperature(Index), Temp_Text)
    If Not IsError Then
       If Not HaveNumber(CDbl(txtgasdifftemperature(Index).Text)) Then
          txtgasdifftemperature(Index).Text = Temp_Text
          txtgasdifftemperature(Index).SetFocus
         Call LostFocus_Handle(Me, txtgasdifftemperature(Index), flag_ok)
          Exit Sub
       End If

       Call TextNumberChanged(ValueChanged, txtgasdifftemperature(Index), Temp_Text)

       If ValueChanged Then
          If CurrentUnits = SIUnits Then
             phprop.GasDiffusivity.input.temperature = CDbl(txtgasdifftemperature(Index).Text)
          Else
             EnglishValue = CDbl(txtgasdifftemperature(Index).Text)
             Call TEMPENSI(SIValue, EnglishValue)
             phprop.GasDiffusivity.input.temperature = SIValue
          End If

       Else
         Call LostFocus_Handle(Me, txtgasdifftemperature(Index), flag_ok)
          Exit Sub
       End If

    End If
  Call LostFocus_Handle(Me, txtgasdifftemperature(Index), flag_ok)

End Sub

Private Sub txtGasDiffusivityValue_GotFocus(Index As Integer)
  Call gotfocus_handle(Me, txtgasdiffusivityvalue(Index), Temp_Text)
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.GasDiffusivity.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

End Sub

Private Sub txtGasDiffusivityValue_KeyPress(Index As Integer, keyascii As Integer)
    
    If keyascii = 13 Then
       keyascii = 0
       txtgasdifftemperature(Index).SetFocus
       Exit Sub
    End If
    Call NumberCheck(keyascii)
   
End Sub

Private Sub txtGasDiffusivityValue_LostFocus(Index As Integer)
    Dim msg As String, response As Integer
    Dim Answer As Integer
    Dim IsError As Integer
    Dim ValueChanged As Integer
    Dim EnglishValue As Double, SIValue As Double
Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtgasdiffusivityvalue(Index))) Then
     Exit Sub
   End If

   flag_ok = True
    If txtgasdiffusivityvalue(Index).Text = "" Then
       phprop.GasDiffusivity.input.Value = -1#
       PROPAVAILABLE(GAS_DIFFUSIVITY_INPUT) = False
       Call LostFocus_Handle(Me, txtgasdiffusivityvalue(Index), flag_ok)
       Exit Sub
    End If

    Call TextHandleError(IsError, txtgasdiffusivityvalue(Index), Temp_Text)
    If Not IsError Then
       If Not HaveNumber(CDbl(txtgasdiffusivityvalue(Index).Text)) Then
          txtgasdiffusivityvalue(Index).Text = Temp_Text
          txtgasdiffusivityvalue(Index).SetFocus
          Call LostFocus_Handle(Me, txtgasdiffusivityvalue(Index), flag_ok)
          Exit Sub
       End If

       Call TextNumberChanged(ValueChanged, txtgasdiffusivityvalue(Index), Temp_Text)

       If ValueChanged Then
          If CurrentUnits = SIUnits Then
             phprop.GasDiffusivity.input.Value = CDbl(txtgasdiffusivityvalue(Index).Text)
          Else
             EnglishValue = CDbl(txtgasdiffusivityvalue(Index).Text)
             Call GDIFENSI(SIValue, EnglishValue)
             phprop.GasDiffusivity.input.Value = SIValue
          End If

          PROPAVAILABLE(GAS_DIFFUSIVITY_INPUT) = True
       Else
          Call LostFocus_Handle(Me, txtgasdiffusivityvalue(Index), flag_ok)
          Exit Sub
       End If
    End If
  Call LostFocus_Handle(Me, txtgasdiffusivityvalue(Index), flag_ok)


End Sub

