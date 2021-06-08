VERSION 5.00
Begin VB.Form mv_nbp_form 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Molar Volume at Normal Boiling Point"
   ClientHeight    =   3510
   ClientLeft      =   1230
   ClientTop       =   4620
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
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3510
   ScaleWidth      =   8505
   Begin VB.TextBox txtMVNBPTemperature 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   6240
      TabIndex        =   9
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox txtMolarVolumeNBPValue 
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   1
      Left            =   4080
      TabIndex        =   8
      Top             =   2880
      Width           =   1932
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
      Caption         =   "Accept Selected Molar Volume"
      Height          =   615
      Left            =   4920
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   2
      Left            =   1080
      TabIndex        =   6
      Top             =   2880
      Width           =   225
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   1080
      TabIndex        =   2
      Top             =   2400
      Width           =   225
   End
   Begin VB.Label lblMVNBPTemperature 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   6240
      TabIndex        =   5
      Top             =   2400
      Width           =   972
   End
   Begin VB.Label lblMolarVolumeNBPValue 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   4080
      TabIndex        =   4
      Top             =   2400
      Width           =   1932
   End
   Begin VB.Line Line2 
      X1              =   960
      X2              =   7440
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   1692
      Left            =   960
      Top             =   1680
      Width           =   6492
   End
   Begin VB.Label lblSource 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Schroeder's Method"
      ForeColor       =   &H00000000&
      Height          =   252
      Index           =   0
      Left            =   1440
      TabIndex        =   3
      Top             =   2400
      Width           =   2412
   End
   Begin VB.Label lblSource 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "User Input"
      ForeColor       =   &H00000000&
      Height          =   252
      Index           =   1
      Left            =   1440
      TabIndex        =   7
      Top             =   2880
      Width           =   2412
   End
   Begin VB.Label lblMVNBPTempLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Temp."
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   6240
      TabIndex        =   17
      Top             =   1800
      Width           =   972
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Source"
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   1440
      TabIndex        =   16
      Top             =   1800
      Width           =   2412
   End
   Begin VB.Label lblMVNBPLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Molar Vol."
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   4080
      TabIndex        =   15
      Top             =   1800
      Width           =   1932
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
      TabIndex        =   14
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
      TabIndex        =   13
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
      TabIndex        =   12
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
      TabIndex        =   11
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Current Molar Volume (NBP)  Information"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "mv_nbp_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PastMVNBPInputValue As String
Dim PastMVNBPInputTemp As String

Private Sub cmdCancel_Click()
    Dim SelectedOption As Integer   'Option selected permanently by the user (the option corresponding to the number on the main screen)

    Select Case phprop.MolarVolume.BoilingPoint.CurrentSelection.choice
       Case MOLAR_VOLUME_NBP_UNIFAC
          SelectedOption = 1
       Case MOLAR_VOLUME_NBP_INPUT
          SelectedOption = 2
          txtMolarVolumeNBPValue(1).Text = PastMVNBPInputValue
          txtMVNBPTemperature(1).Text = PastMVNBPInputTemp
       Case Else
          mv_nbp_form.Hide
          Exit Sub
    End Select

    If Not Option1(SelectedOption).Value Then Option1(SelectedOption).Value = True

    mv_nbp_form.Hide

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
          ValueToDisplayIndex = MOLAR_VOLUME_NBP_UNIFAC
       Case 2
          If Not PROPAVAILABLE(MOLAR_VOLUME_NBP_INPUT) Then
             MsgBox "User Input can not be selected without first entering a value", MB_ICONSTOP, "Error"
             txtMolarVolumeNBPValue(1).SetFocus
             Exit Sub
          End If
          ValueToDisplayIndex = MOLAR_VOLUME_NBP_INPUT
    End Select

    If ValueToDisplayIndex <> phprop.MolarVolume.BoilingPoint.CurrentSelection.choice Then
       phprop.MolarVolume.BoilingPoint.CurrentSelection.choice = ValueToDisplayIndex
       Call DisplayMolarVolumeNBPMainScreen(ValueToDisplayIndex)
    ElseIf ValueToDisplayIndex = MOLAR_VOLUME_NBP_INPUT Then
       Call DisplayMolarVolumeNBPMainScreen(ValueToDisplayIndex)
    End If

    mv_nbp_form.Hide

'Recalculate Properties using selected Molar Volume at NBP

'          frmWaitForCalculations.Show
'          frmWaitForCalculations.Refresh

          contam_prop_form!lblContaminantProperties(5).Caption = ""
          contam_prop_form!lblContaminantProperties(11).Caption = ""
          contam_prop_form!lblContaminantProperties(12).Caption = ""

          Screen.MousePointer = 11   'Hourglass

          Call CalculateLiquidDensity
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
    
    PastMVNBPInputValue = txtMolarVolumeNBPValue(1).Text
    PastMVNBPInputTemp = txtMVNBPTemperature(1).Text
End Sub

Private Sub Form_Load()

  Call centerform_relative(contam_prop_form, Me)

    If (DemoMode) Then cmdok.Enabled = False


End Sub

Private Sub lblMolarVolumeNBPValue_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
0    Dim ErrorFlag As Integer
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.MolarVolumeBoilingPoint.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

    If Button <> 2 Then Exit Sub

    Select Case Index
       Case 0  'Source = Schroeder's Group Contribution Method
          ErrorFlag = phprop.MolarVolume.BoilingPoint.UNIFAC.error
    End Select

    If ErrorFlag = 0 Then Exit Sub

    If ErrorFlag < 0 Then
       MsgBox ErrorMsg(ErrorFlag), MB_ICONINFORMATION, Trim$(phprop.Name) & " - Data Unavailable"
    End If

    If ErrorFlag > 0 Then
       MsgBox ErrorMsg(ErrorFlag), MB_ICONINFORMATION, Trim$(phprop.Name) & " - Warning"
    End If

End Sub

Private Sub lblMVNBPTemperature_Click(Index As Integer)
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.MolarVolumeBoilingPoint.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

End Sub

Private Sub lblSource_Click(Index As Integer)
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.MolarVolumeBoilingPoint.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

End Sub

Private Sub Option1_Click(Index As Integer)
    Dim i As Integer, SourceIndex As Integer

    SourceIndex = Index - 1
    If SourceIndex = hilight.MolarVolumeBoilingPoint.PreviousIndex Then Exit Sub
    lblSource(SourceIndex).BackColor = &H800000
    lblSource(SourceIndex).ForeColor = &H80000005
    i = hilight.MolarVolumeBoilingPoint.PreviousIndex
    hilight.MolarVolumeBoilingPoint.PreviousIndex = SourceIndex
    If i = -1 Then Exit Sub
    If Option1(i + 1).Enabled = False Then Exit Sub
    lblSource(i).BackColor = &H80000005
    lblSource(i).ForeColor = &H80000008

End Sub

Private Sub txtMolarVolumeNBPValue_GotFocus(Index As Integer)
  Call gotfocus_handle(Me, txtMolarVolumeNBPValue(Index), Temp_Text)
    
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.MolarVolumeBoilingPoint.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If


End Sub

Private Sub txtMolarVolumeNBPValue_KeyPress(Index As Integer, keyascii As Integer)
    
    If keyascii = 13 Then
       keyascii = 0
       txtMVNBPTemperature(Index).SetFocus
       Exit Sub
    End If
    Call NumberCheck(keyascii)
  
End Sub

Private Sub txtMolarVolumeNBPValue_LostFocus(Index As Integer)
    Dim msg As String, response As Integer
    Dim Answer As Integer
    Dim IsError As Integer
    Dim ValueChanged As Integer
    Dim EnglishValue As Double, SIValue As Double
     Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtMolarVolumeNBPValue(Index))) Then
     Exit Sub
   End If

   flag_ok = True

    If txtMolarVolumeNBPValue(Index).Text = "" Then
       phprop.MolarVolume.BoilingPoint.input.Value = -1#
       PROPAVAILABLE(MOLAR_VOLUME_NBP_INPUT) = False
       Call LostFocus_Handle(Me, txtMolarVolumeNBPValue(Index), flag_ok)
       Exit Sub
    End If

    Call TextHandleError(IsError, txtMolarVolumeNBPValue(Index), Temp_Text)
    If Not IsError Then
       If Not HaveNumber(CDbl(txtMolarVolumeNBPValue(Index).Text)) Then
          txtMolarVolumeNBPValue(Index).Text = Temp_Text
          txtMolarVolumeNBPValue(Index).SetFocus
          Call LostFocus_Handle(Me, txtMolarVolumeNBPValue(Index), flag_ok)
          Exit Sub
       End If

       Call TextNumberChanged(ValueChanged, txtMolarVolumeNBPValue(Index), Temp_Text)

       If ValueChanged Then
          If CurrentUnits = SIUnits Then
             phprop.MolarVolume.BoilingPoint.input.Value = CDbl(txtMolarVolumeNBPValue(Index).Text)
          Else
             EnglishValue = CDbl(txtMolarVolumeNBPValue(Index).Text)
             Call MVBPENSI(SIValue, EnglishValue)
             phprop.MolarVolume.BoilingPoint.input.Value = SIValue
          End If
          
          PROPAVAILABLE(MOLAR_VOLUME_NBP_INPUT) = True
         Else
          Call LostFocus_Handle(Me, txtMolarVolumeNBPValue(Index), flag_ok)
          Exit Sub
       End If
    End If

  Call LostFocus_Handle(Me, txtMolarVolumeNBPValue(Index), flag_ok)
End Sub

Private Sub txtMVNBPTemperature_GotFocus(Index As Integer)
  Call gotfocus_handle(Me, txtMVNBPTemperature(Index), Temp_Text)

    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.MolarVolumeBoilingPoint.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

End Sub

Private Sub txtMVNBPTemperature_KeyPress(Index As Integer, keyascii As Integer)
    
    If keyascii = 13 Then
       keyascii = 0
       cmdok.SetFocus
       Exit Sub
    End If
    Call NumberCheck(keyascii)
   
End Sub

Private Sub txtMVNBPTemperature_LostFocus(Index As Integer)
    Dim msg As String, response As Integer
    Dim Answer As Integer
    Dim IsError As Integer
    Dim ValueChanged As Integer
    Dim EnglishValue As Double, SIValue As Double
    Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtMVNBPTemperature(Index))) Then
     Exit Sub
   End If

   flag_ok = True

    If txtMVNBPTemperature(Index).Text = "" Then
       phprop.MolarVolume.BoilingPoint.input.temperature = -1E+25
       Call LostFocus_Handle(Me, txtMVNBPTemperature(Index), flag_ok)
       Exit Sub
    End If

    Call TextHandleError(IsError, txtMVNBPTemperature(Index), Temp_Text)
    If Not IsError Then
       If Not HaveNumber(CDbl(txtMVNBPTemperature(Index).Text)) Then
          txtMVNBPTemperature(Index).Text = Temp_Text
          txtMVNBPTemperature(Index).SetFocus
       Call LostFocus_Handle(Me, txtMVNBPTemperature(Index), flag_ok)
          Exit Sub
       End If

       Call TextNumberChanged(ValueChanged, txtMVNBPTemperature(Index), Temp_Text)

       If ValueChanged Then
          If CurrentUnits = SIUnits Then
             phprop.MolarVolume.BoilingPoint.input.temperature = CDbl(txtMVNBPTemperature(Index).Text)
          Else
             EnglishValue = CDbl(txtMVNBPTemperature(Index).Text)
             Call TEMPENSI(SIValue, EnglishValue)
             phprop.MolarVolume.BoilingPoint.input.temperature = SIValue
          End If
          
       Else
       Call LostFocus_Handle(Me, txtMVNBPTemperature(Index), flag_ok)
          Exit Sub
       End If

    End If
       Call LostFocus_Handle(Me, txtMVNBPTemperature(Index), flag_ok)

End Sub

