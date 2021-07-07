VERSION 5.00
Begin VB.Form nbp_form 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Normal Boiling Point"
   ClientHeight    =   3585
   ClientLeft      =   615
   ClientTop       =   1725
   ClientWidth     =   8400
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
   ScaleHeight     =   3585
   ScaleWidth      =   8400
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      Caption         =   "Accept Selected Normal Boiling Point"
      Height          =   615
      Left            =   4800
      TabIndex        =   0
      Top             =   120
      Width           =   3495
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
   Begin VB.TextBox txtNormalBPValue 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   4920
      TabIndex        =   5
      Top             =   2880
      Width           =   2055
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
      Width           =   270
   End
   Begin VB.Label lblNormalBPValue 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   4920
      TabIndex        =   3
      Top             =   2400
      Width           =   2052
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Current Normal Boiling Point Information"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   7
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
      TabIndex        =   14
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
      TabIndex        =   13
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
      TabIndex        =   12
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
      TabIndex        =   11
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
   Begin VB.Label lblNBPLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Normal Boiling Point"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4920
      TabIndex        =   10
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Source"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1800
      TabIndex        =   9
      Top             =   1800
      Width           =   2895
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
      Left            =   1800
      TabIndex        =   8
      Top             =   2400
      Width           =   2895
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
      Left            =   1800
      TabIndex        =   6
      Top             =   2880
      Width           =   2895
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   1695
      Left            =   1320
      Top             =   1680
      Width           =   5895
   End
   Begin VB.Line Line2 
      X1              =   1320
      X2              =   7200
      Y1              =   2160
      Y2              =   2160
   End
End
Attribute VB_Name = "nbp_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PastBPInputValue As String

Private Sub cmdCancel_Click()
    Dim SelectedOption As Integer   'Option selected permanently by the user (the option corresponding to the number on the main screen)

    Select Case phprop.BoilingPoint.CurrentSelection.choice
       Case BOILING_POINT_DATABASE
          SelectedOption = 1
       Case BOILING_POINT_INPUT
          SelectedOption = 2
          txtNormalBPValue(1).Text = PastBPInputValue
       Case Else
          nbp_form.Hide
          Exit Sub
    End Select

    If Not Option1(SelectedOption).Value Then Option1(SelectedOption).Value = True

    nbp_form.Hide
End Sub

Private Sub cmdok_Click()
    Dim OptionSelected As Integer
    Dim ValueToDisplayIndex As Integer
    Dim i As Integer
    Dim NumContaminantInList As Integer

'*** Pass new selected value back to main screen
    For i = 1 To 2
        If Option1(i).Value Then
           OptionSelected = i
           Exit For
        End If
    Next i

    Select Case OptionSelected
       Case 1
          ValueToDisplayIndex = BOILING_POINT_DATABASE
       Case 2
          If Not PROPAVAILABLE(BOILING_POINT_INPUT) Then
             MsgBox "User Input can not be selected without first entering a value", MB_ICONSTOP, "Error"
             txtNormalBPValue(1).SetFocus
             Exit Sub
          End If
          ValueToDisplayIndex = BOILING_POINT_INPUT
    End Select

    If ValueToDisplayIndex <> phprop.BoilingPoint.CurrentSelection.choice Then
       phprop.BoilingPoint.CurrentSelection.choice = ValueToDisplayIndex
       Call DisplayBoilingPointMainScreen(ValueToDisplayIndex)
    ElseIf ValueToDisplayIndex = BOILING_POINT_INPUT Then
       Call DisplayBoilingPointMainScreen(ValueToDisplayIndex)
    End If

    nbp_form.Hide

'Recalculate Gas Diffusivity using selected Normal Boiling Point

'          frmWaitForCalculations.Show
'          frmWaitForCalculations.Refresh

          contam_prop_form!lblContaminantProperties(12).Caption = ""

          Screen.MousePointer = 11   'Hourglass

          Call CalculateGasDiffusivity
          contam_prop_form.Refresh

          Screen.MousePointer = 0    'Arrow

'          frmWaitForCalculations.Hide

          NumContaminantInList = contam_prop_form!cboSelectContaminant.ListIndex + 1
          PropContaminant(NumContaminantInList) = phprop

End Sub

Private Sub Form_Activate()
    
  Call centerform_relative(contam_prop_form, Me)
    
    PastBPInputValue = txtNormalBPValue(1).Text
End Sub

Private Sub Form_Load()

  Call centerform_relative(contam_prop_form, Me)
    
    If (DemoMode) Then cmdok.Enabled = False


End Sub

Private Sub lblNormalBPValue_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ErrorFlag As Integer
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.BoilingPoint.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

    If Button <> 2 Then Exit Sub

    Select Case Index
       Case 0   'Database
          ErrorFlag = phprop.BoilingPoint.database.error
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
       If Index = hilight.BoilingPoint.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

End Sub

Private Sub Option1_Click(Index As Integer)
    Dim i As Integer, SourceIndex As Integer

    SourceIndex = Index - 1
    If SourceIndex = hilight.BoilingPoint.PreviousIndex Then Exit Sub
    lblSource(SourceIndex).BackColor = &H800000
    lblSource(SourceIndex).ForeColor = &H80000005
    i = hilight.BoilingPoint.PreviousIndex
    hilight.BoilingPoint.PreviousIndex = SourceIndex
    If i = -1 Then Exit Sub
    If Option1(i + 1).Enabled = False Then Exit Sub
    lblSource(i).BackColor = &H80000005
    lblSource(i).ForeColor = &H80000008

End Sub

Private Sub txtNormalBPValue_GotFocus(Index As Integer)
  Call gotfocus_handle(Me, txtNormalBPValue(Index), Temp_Text)
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.BoilingPoint.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

End Sub

Private Sub txtNormalBPValue_KeyPress(Index As Integer, keyascii As Integer)
    
    If keyascii = 13 Then
       keyascii = 0
       cmdok.SetFocus
       Exit Sub
    End If
    Call NumberCheck(keyascii)
    
End Sub

Private Sub txtNormalBPValue_LostFocus(Index As Integer)
    Dim msg As String, response As Integer
    Dim Answer As Integer
    Dim IsError As Integer
    Dim ValueChanged As Integer
    Dim EnglishValue As Double, SIValue As Double
    Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtNormalBPValue(Index))) Then
     Exit Sub
   End If

   flag_ok = True

    If txtNormalBPValue(Index).Text = "" Then
       phprop.BoilingPoint.input.Value = -1E+25
       PROPAVAILABLE(BOILING_POINT_INPUT) = False
        Call LostFocus_Handle(Me, txtNormalBPValue(Index), flag_ok)
       Exit Sub
    End If

    Call TextHandleError(IsError, txtNormalBPValue(Index), Temp_Text)
    If Not IsError Then
       If Not HaveNumber2(CDbl(txtNormalBPValue(Index).Text)) Then
          txtNormalBPValue(Index).Text = Temp_Text
          txtNormalBPValue(Index).SetFocus
        Call LostFocus_Handle(Me, txtNormalBPValue(Index), flag_ok)
          Exit Sub
       End If

       Call TextNumberChanged(ValueChanged, txtNormalBPValue(Index), Temp_Text)

       If ValueChanged Then
          If CurrentUnits = SIUnits Then
             phprop.BoilingPoint.input.Value = CDbl(txtNormalBPValue(Index).Text)
          Else
             EnglishValue = CDbl(txtNormalBPValue(Index).Text)
             Call NBPENSI(SIValue, EnglishValue)
             phprop.BoilingPoint.input.Value = SIValue
          End If
          
          PROPAVAILABLE(BOILING_POINT_INPUT) = True
       Else
        Call LostFocus_Handle(Me, txtNormalBPValue(Index), flag_ok)
          Exit Sub
       End If
    End If

        Call LostFocus_Handle(Me, txtNormalBPValue(Index), flag_ok)
End Sub

