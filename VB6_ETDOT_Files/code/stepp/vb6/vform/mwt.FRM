VERSION 5.00
Begin VB.Form mwt_form 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Molecular Weight"
   ClientHeight    =   4035
   ClientLeft      =   600
   ClientTop       =   2115
   ClientWidth     =   8445
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
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4035
   ScaleWidth      =   8445
   Begin VB.TextBox txtMolecularWeightValue 
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   2
      Left            =   4920
      TabIndex        =   7
      Top             =   3360
      Width           =   2412
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   615
      Left            =   4800
      TabIndex        =   1
      Top             =   840
      Width           =   3495
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      Caption         =   "Accept Selected Molecular Weight"
      Height          =   615
      Left            =   4800
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   3
      Left            =   1440
      TabIndex        =   6
      Top             =   3360
      Width           =   285
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   2
      Left            =   1440
      TabIndex        =   4
      Top             =   2880
      Width           =   285
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   1440
      TabIndex        =   2
      Top             =   2400
      Width           =   285
   End
   Begin VB.Label lblMolecularWeightValue 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   4920
      TabIndex        =   5
      Top             =   2880
      Width           =   2412
   End
   Begin VB.Label lblMolecularWeightValue 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   4920
      TabIndex        =   3
      Top             =   2400
      Width           =   2412
   End
   Begin VB.Line Line2 
      X1              =   1320
      X2              =   7560
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   2175
      Left            =   1320
      Top             =   1740
      Width           =   6255
   End
   Begin VB.Label lblSourceLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "User Input"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   1800
      TabIndex        =   17
      Top             =   3360
      Width           =   2895
   End
   Begin VB.Label lblSourceLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Group Contribution Method"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   16
      Top             =   2880
      Width           =   2895
   End
   Begin VB.Label lblSourceLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Database"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   1800
      TabIndex        =   15
      Top             =   2400
      Width           =   2895
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Source"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1800
      TabIndex        =   14
      Top             =   1800
      Width           =   2895
   End
   Begin VB.Label lblMWTLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Molecular Weight"
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   4920
      TabIndex        =   13
      Top             =   1800
      Width           =   2412
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4560
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   2
      Height          =   1332
      Left            =   120
      Top             =   120
      Width           =   4452
   End
   Begin VB.Label lblCurrentValues 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   2160
      TabIndex        =   12
      Top             =   720
      Width           =   2292
   End
   Begin VB.Label lblCurrentValues 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   2160
      TabIndex        =   11
      Top             =   1080
      Width           =   2292
   End
   Begin VB.Label lblCurrentInformation 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Source"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   240
      TabIndex        =   10
      Top             =   1080
      Width           =   1692
   End
   Begin VB.Label lblCurrentInformation 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Value"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   720
      Width           =   1692
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Current Molecular Weight Information"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "mwt_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PastMWTInputValue As String

Private Sub cmdCancel_Click()
    Dim SelectedOption As Integer   'Option selected permanently by the user (the option corresponding to the number on the main screen)

    Select Case phprop.MolecularWeight.CurrentSelection.choice
       Case MOLECULAR_WEIGHT_DATABASE
          SelectedOption = 1
       Case MOLECULAR_WEIGHT_UNIFAC
          SelectedOption = 2
       Case MOLECULAR_WEIGHT_INPUT
          SelectedOption = 3
          txtMolecularWeightValue(2).Text = PastMWTInputValue
       Case Else
          mwt_form.Hide
          Exit Sub
    End Select

    If Not Option1(SelectedOption).Value Then Option1(SelectedOption).Value = True

    mwt_form.Hide
End Sub

Private Sub cmdok_Click()
    Dim OptionSelected As Integer
    Dim ValueToDisplayIndex As Integer
    Dim i As Integer
    Dim NumContaminantInList As Integer

'*** Pass new selected value back to main screen
    For i = 1 To 3
        If Option1(i).Value Then
           OptionSelected = i
           Exit For
        End If
    Next i

    Select Case OptionSelected
       Case 1
          ValueToDisplayIndex = MOLECULAR_WEIGHT_DATABASE
       Case 2
          ValueToDisplayIndex = MOLECULAR_WEIGHT_UNIFAC
       Case 3
          If Not PROPAVAILABLE(MOLECULAR_WEIGHT_INPUT) Then
             MsgBox "User Input can not be selected without first entering a value", MB_ICONSTOP, "Error"
             txtMolecularWeightValue(2).SetFocus
             Exit Sub
          End If
          ValueToDisplayIndex = MOLECULAR_WEIGHT_INPUT
    End Select

    If ValueToDisplayIndex <> phprop.MolecularWeight.CurrentSelection.choice Then
       phprop.MolecularWeight.CurrentSelection.choice = ValueToDisplayIndex
       Call DisplayMolecularWeightMainScreen(ValueToDisplayIndex)
    ElseIf ValueToDisplayIndex = HENRYS_CONSTANT_INPUT Then
       Call DisplayMolecularWeightMainScreen(ValueToDisplayIndex)
    End If

    mwt_form.Hide

'Recalculate Properties using selected Molecular Weight

'          frmWaitForCalculations.Show
'          frmWaitForCalculations.Refresh

          contam_prop_form!lblContaminantProperties(5).Caption = ""
          contam_prop_form!lblContaminantProperties(6).Caption = ""
          contam_prop_form!lblContaminantProperties(11).Caption = ""
          contam_prop_form!lblContaminantProperties(12).Caption = ""

          Screen.MousePointer = 11   'Hourglass

          Call CalculateLiquidDensity
          contam_prop_form.Refresh
          Call CalculateMolarVolumeOpT
          contam_prop_form.Refresh
          Call CalculateLiquidDiffusivity
          contam_prop_form.Refresh
          Call CalculateGasDiffusivity
          contam_prop_form.Refresh

          Screen.MousePointer = 0    'Arrow

'          frmWaitForCalculations.Hide

          NumContaminantInList = contam_prop_form!cboSelectContaminant.ListIndex + 1
          PropContaminant(NumContaminantInList) = phprop

End Sub

Private Sub Form_Activate()
    
  Call centerform_relative(contam_prop_form, Me)
    
    PastMWTInputValue = txtMolecularWeightValue(2).Text
End Sub

Private Sub Form_Load()

  Call centerform_relative(contam_prop_form, Me)
    
    If (DemoMode) Then cmdok.Enabled = False


End Sub

Private Sub lblMolecularWeightValue_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ErrorFlag As Integer
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.MolecularWeight.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

    If Button <> 2 Then Exit Sub

    Select Case Index
       Case 0   'Database
          ErrorFlag = phprop.MolecularWeight.database.error
       Case 1   'Group Contribution Method
          ErrorFlag = phprop.MolecularWeight.UNIFAC.error
    End Select

    If ErrorFlag = 0 Then Exit Sub

    If ErrorFlag < 0 Then
       MsgBox ErrorMsg(ErrorFlag), MB_ICONINFORMATION, Trim$(phprop.Name) & " - Data Unavailable"
    End If

    If ErrorFlag > 0 Then
       MsgBox ErrorMsg(ErrorFlag), MB_ICONINFORMATION, Trim$(phprop.Name) & " - Warning"
    End If

End Sub

Private Sub lblSourceLabel_Click(Index As Integer)
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.MolecularWeight.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

End Sub

Private Sub Option1_Click(Index As Integer)
    Dim i As Integer, SourceIndex As Integer

    SourceIndex = Index - 1
    If SourceIndex = hilight.MolecularWeight.PreviousIndex Then Exit Sub
    lblSourceLabel(SourceIndex).BackColor = &H800000
    lblSourceLabel(SourceIndex).ForeColor = &H80000005
    i = hilight.MolecularWeight.PreviousIndex
    hilight.MolecularWeight.PreviousIndex = SourceIndex
    If i = -1 Then Exit Sub
    If Option1(i + 1).Enabled = False Then Exit Sub
    lblSourceLabel(i).BackColor = &H80000005
    lblSourceLabel(i).ForeColor = &H80000008

End Sub

Private Sub txtMolecularWeightValue_GotFocus(Index As Integer)
  Call gotfocus_handle(Me, txtMolecularWeightValue(Index), Temp_Text)
    
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.MolecularWeight.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If
End Sub

Private Sub txtMolecularWeightValue_KeyPress(Index As Integer, keyascii As Integer)
    
    If keyascii = 13 Then
       keyascii = 0
       cmdok.SetFocus
       Exit Sub
    End If
    Call NumberCheck(keyascii)
    
End Sub

Private Sub txtMolecularWeightValue_LostFocus(Index As Integer)
    Dim msg As String, response As Integer
    Dim Answer As Integer
    Dim IsError As Integer
    Dim ValueChanged As Integer
    Dim EnglishValue As Double, SIValue As Double
    Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtMolecularWeightValue(Index))) Then
     Exit Sub
   End If

   flag_ok = True

    If txtMolecularWeightValue(Index).Text = "" Then
       phprop.MolecularWeight.input.Value = -1#
       PROPAVAILABLE(MOLECULAR_WEIGHT_INPUT) = False
         Call LostFocus_Handle(Me, txtMolecularWeightValue(Index), flag_ok)
       Exit Sub
    End If

    Call TextHandleError(IsError, txtMolecularWeightValue(Index), Temp_Text)
    If Not IsError Then
       If Not HaveNumber(CDbl(txtMolecularWeightValue(Index).Text)) Then
          txtMolecularWeightValue(Index).Text = Temp_Text
          txtMolecularWeightValue(Index).SetFocus
         Call LostFocus_Handle(Me, txtMolecularWeightValue(Index), flag_ok)
          Exit Sub
       End If

       Call TextNumberChanged(ValueChanged, txtMolecularWeightValue(Index), Temp_Text)

       If ValueChanged Then
          If CurrentUnits = SIUnits Then
             phprop.MolecularWeight.input.Value = CDbl(txtMolecularWeightValue(Index).Text)
          Else
             EnglishValue = CDbl(txtMolecularWeightValue(Index).Text)
             Call MWENSI(SIValue, EnglishValue)
             phprop.MolecularWeight.input.Value = SIValue
          End If

          PROPAVAILABLE(MOLECULAR_WEIGHT_INPUT) = True
       Else
         Call LostFocus_Handle(Me, txtMolecularWeightValue(Index), flag_ok)
          Exit Sub
       End If
    End If
         Call LostFocus_Handle(Me, txtMolecularWeightValue(Index), flag_ok)

End Sub

