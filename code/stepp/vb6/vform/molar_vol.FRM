VERSION 5.00
Begin VB.Form molar_vol_form 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Molar Volume at Operating Temperature"
   ClientHeight    =   4005
   ClientLeft      =   675
   ClientTop       =   1995
   ClientWidth     =   8535
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
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4005
   ScaleWidth      =   8535
   Begin VB.TextBox txtMVOpTmaximumT 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   7380
      TabIndex        =   8
      Text            =   "not visible"
      Top             =   3360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtMVOpTminimumT 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   6240
      TabIndex        =   7
      Text            =   "not visible"
      Top             =   3360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtMVOpTTemperature 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   5160
      TabIndex        =   6
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox txtMolarVolumeOpTValue 
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   2
      Left            =   3240
      TabIndex        =   5
      Top             =   3360
      Width           =   1812
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
      Left            =   240
      TabIndex        =   2
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
      TabIndex        =   3
      Top             =   3360
      Width           =   285
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   2400
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   2172
      Left            =   120
      Top             =   1680
      Width           =   8292
   End
   Begin VB.Label lblMVOpTmaximumT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "not visible"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   7320
      TabIndex        =   29
      Top             =   2880
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label lblMVOpTmaximumT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   7320
      TabIndex        =   28
      Top             =   2400
      Width           =   972
   End
   Begin VB.Label lblMVOpTminimumT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "not visible"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   6240
      TabIndex        =   27
      Top             =   2880
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label lblMVOpTminimumT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   6240
      TabIndex        =   26
      Top             =   2400
      Width           =   972
   End
   Begin VB.Label lblMVOpTTemperature 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   5160
      TabIndex        =   25
      Top             =   2880
      Width           =   972
   End
   Begin VB.Label lblMVOpTTemperature 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   5160
      TabIndex        =   24
      Top             =   2400
      Width           =   972
   End
   Begin VB.Label lblMolarVolumeOpTValue 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   3240
      TabIndex        =   23
      Top             =   2880
      Width           =   1812
   End
   Begin VB.Label lblMolarVolumeOpTValue 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   3240
      TabIndex        =   22
      Top             =   2400
      Width           =   1812
   End
   Begin VB.Label lblMVOpTminTLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tmin"
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   6240
      TabIndex        =   21
      Top             =   1800
      Width           =   972
   End
   Begin VB.Label lblMVOpTMaxTLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tmax"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7320
      TabIndex        =   20
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label lblMVOpTTempLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Temp."
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   5160
      TabIndex        =   19
      Top             =   1800
      Width           =   972
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   8400
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
      Index           =   2
      Left            =   600
      TabIndex        =   4
      Top             =   3360
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
      TabIndex        =   18
      Top             =   2880
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
      TabIndex        =   17
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Source"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   600
      TabIndex        =   16
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label lblMVOpTLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Molar Vol."
      ForeColor       =   &H00000000&
      Height          =   400
      Left            =   3240
      TabIndex        =   15
      Top             =   1750
      Width           =   1812
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
      Caption         =   "Current Molar Volume (Operating T)  Information"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "molar_vol_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PastMVopTInputValue As String
Dim PastMVopTInputTemp As String

Private Sub cmdCancel_Click()
    Dim SelectedOption As Integer   'Option selected permanently by the user (the option corresponding to the number on the main screen)

    Select Case phprop.MolarVolume.operatingT.CurrentSelection.choice
       Case MOLAR_VOLUME_OPT_DATABASE
          SelectedOption = 1
       Case MOLAR_VOLUME_OPT_UNIFAC
          SelectedOption = 2
       Case MOLAR_VOLUME_OPT_INPUT
          SelectedOption = 3
          txtMolarVolumeOPTvalue(2).Text = PastMVopTInputValue
          txtmvOPTTemperature(2).Text = PastMVopTInputTemp
       Case Else
          molar_vol_form.Hide
          Exit Sub
    End Select

    If Not Option1(SelectedOption).Value Then Option1(SelectedOption).Value = True

    molar_vol_form.Hide
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
          ValueToDisplayIndex = MOLAR_VOLUME_OPT_DATABASE
       Case 2
          ValueToDisplayIndex = MOLAR_VOLUME_OPT_UNIFAC
       Case 3
          If Not PROPAVAILABLE(MOLAR_VOLUME_OPT_INPUT) Then
             MsgBox "User Input can not be selected without first entering a value", MB_ICONSTOP, "Error"
             txtMolarVolumeOPTvalue(2).SetFocus
             Exit Sub
          End If
          ValueToDisplayIndex = MOLAR_VOLUME_OPT_INPUT
    End Select

    If ValueToDisplayIndex <> phprop.MolarVolume.operatingT.CurrentSelection.choice Then
       phprop.MolarVolume.operatingT.CurrentSelection.choice = ValueToDisplayIndex
       Call DisplayMolarVolumeOpTMainScreen(ValueToDisplayIndex)
    ElseIf ValueToDisplayIndex = MOLAR_VOLUME_OPT_INPUT Then
       Call DisplayMolarVolumeOpTMainScreen(ValueToDisplayIndex)
    End If

    molar_vol_form.Hide

End Sub

Private Sub Form_Activate()

  Call centerform_relative(contam_prop_form, Me)

    PastMVopTInputValue = txtMolarVolumeOPTvalue(2).Text
    PastMVopTInputTemp = txtmvOPTTemperature(2).Text
End Sub

Private Sub Form_Load()

  Call centerform_relative(contam_prop_form, Me)
    
    If (DemoMode) Then cmdok.Enabled = False

End Sub

Private Sub lblMolarVolumeOpTValue_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ErrorFlag As Integer
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.MolarVolumeOperatingT.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

    If Button <> 2 Then Exit Sub

    Select Case Index
       Case 0  'Source = Database
          ErrorFlag = phprop.MolarVolume.operatingT.database.error
       Case 1  'Source = Group Contribution Method
          ErrorFlag = phprop.MolarVolume.operatingT.UNIFAC.error
    End Select

    If ErrorFlag = 0 Then Exit Sub

    If ErrorFlag < 0 Then
       MsgBox ErrorMsg(ErrorFlag), MB_ICONINFORMATION, Trim$(phprop.Name) & " - Data Unavailable"
    End If

    If ErrorFlag > 0 Then
       MsgBox ErrorMsg(ErrorFlag), MB_ICONINFORMATION, Trim$(phprop.Name) & " - Warning"
    End If

End Sub

Private Sub lblMVOpTmaximumT_Click(Index As Integer)
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.MolarVolumeOperatingT.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

End Sub

Private Sub lblMVOpTminimumT_Click(Index As Integer)
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.MolarVolumeOperatingT.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

End Sub

Private Sub lblMVOpTTemperature_Click(Index As Integer)
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.MolarVolumeOperatingT.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

End Sub

Private Sub lblSource_Click(Index As Integer)
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.MolarVolumeOperatingT.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

End Sub

Private Sub Option1_Click(Index As Integer)
    Dim i As Integer, SourceIndex As Integer

    SourceIndex = Index - 1
    If SourceIndex = hilight.MolarVolumeOperatingT.PreviousIndex Then Exit Sub
    lblSource(SourceIndex).BackColor = &H800000
    lblSource(SourceIndex).ForeColor = &H80000005
    i = hilight.MolarVolumeOperatingT.PreviousIndex
    hilight.MolarVolumeOperatingT.PreviousIndex = SourceIndex
    If i = -1 Then Exit Sub
    If Option1(i + 1).Enabled = False Then Exit Sub
    lblSource(i).BackColor = &H80000005
    lblSource(i).ForeColor = &H80000008

End Sub

Private Sub txtMolarVolumeOpTValue_GotFocus(Index As Integer)
  Call gotfocus_handle(Me, txtMolarVolumeOPTvalue(Index), Temp_Text)
    
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.MolarVolumeOperatingT.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

End Sub

Private Sub txtMolarVolumeOpTValue_KeyPress(Index As Integer, keyascii As Integer)
    
    If keyascii = 13 Then
       keyascii = 0
       txtmvOPTTemperature(Index).SetFocus
       Exit Sub
    End If
    Call NumberCheck(keyascii)
    
End Sub

Private Sub txtMolarVolumeOpTValue_LostFocus(Index As Integer)
    Dim msg As String, response As Integer
    Dim Answer As Integer
    Dim IsError As Integer
    Dim ValueChanged As Integer
    Dim EnglishValue As Double, SIValue As Double
    Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtMolarVolumeOPTvalue(Index))) Then
     Exit Sub
   End If

   flag_ok = True

    If txtMolarVolumeOPTvalue(Index).Text = "" Then
       phprop.MolarVolume.operatingT.input.Value = -1#
       PROPAVAILABLE(MOLAR_VOLUME_OPT_INPUT) = False
       Call LostFocus_Handle(Me, txtMolarVolumeOPTvalue(Index), flag_ok)
       Exit Sub
    End If

    Call TextHandleError(IsError, txtMolarVolumeOPTvalue(Index), Temp_Text)
    If Not IsError Then
       If Not HaveNumber(CDbl(txtMolarVolumeOPTvalue(Index).Text)) Then
          txtMolarVolumeOPTvalue(Index).Text = Temp_Text
          txtMolarVolumeOPTvalue(Index).SetFocus
          Call LostFocus_Handle(Me, txtMolarVolumeOPTvalue(Index), flag_ok)
          Exit Sub
       End If

       Call TextNumberChanged(ValueChanged, txtMolarVolumeOPTvalue(Index), Temp_Text)

       If ValueChanged Then
          If CurrentUnits = SIUnits Then
             phprop.MolarVolume.operatingT.input.Value = CDbl(txtMolarVolumeOPTvalue(Index).Text)
          Else
             EnglishValue = CDbl(txtMolarVolumeOPTvalue(Index).Text)
             Call MVOTENSI(SIValue, EnglishValue)
             phprop.MolarVolume.operatingT.input.Value = SIValue
          End If
          
          PROPAVAILABLE(MOLAR_VOLUME_OPT_INPUT) = True
       Else
          Call LostFocus_Handle(Me, txtMolarVolumeOPTvalue(Index), flag_ok)
          Exit Sub
       End If
    End If
    
    Call LostFocus_Handle(Me, txtMolarVolumeOPTvalue(Index), flag_ok)
End Sub

Private Sub txtMVOpTemperature_KeyPress(Index As Integer, keyascii As Integer)
    If Index <> 2 Then keyascii = 0
End Sub

Private Sub txtMVOpTmaximumT_GotFocus(Index As Integer)
  Call gotfocus_handle(Me, txtmvOPTmaximumT(Index), Temp_Text)

End Sub

Private Sub txtMVOpTmaximumT_KeyPress(Index As Integer, keyascii As Integer)
    If Index <> 2 Then keyascii = 0
End Sub

Private Sub txtMVOpTmaximumT_LostFocus(Index As Integer)
Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtmvOPTmaximumT(Index))) Then
     Exit Sub
   End If

   flag_ok = True
  Call LostFocus_Handle(Me, txtmvOPTmaximumT(Index), flag_ok)

End Sub

Private Sub txtMVOpTminimimumT_KeyPress(Index As Integer, keyascii As Integer)
    If Index <> 2 Then keyascii = 0
End Sub

Private Sub txtMVOpTminimumT_GotFocus(Index As Integer)
  Call gotfocus_handle(Me, txtmvOPTminimumt(Index), Temp_Text)

End Sub

Private Sub txtMVOpTminimumT_LostFocus(Index As Integer)
Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtmvOPTminimumt(Index))) Then
     Exit Sub
   End If

   flag_ok = True
  Call LostFocus_Handle(Me, txtmvOPTminimumt(Index), flag_ok)

End Sub

Private Sub txtMVOpTTemperature_GotFocus(Index As Integer)

  Call gotfocus_handle(Me, txtmvOPTTemperature(Index), Temp_Text)
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.MolarVolumeOperatingT.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If
End Sub

Private Sub txtMVOpTTemperature_KeyPress(Index As Integer, keyascii As Integer)
    
    If keyascii = 13 Then
       keyascii = 0
       cmdok.SetFocus
       Exit Sub
    End If
    Call NumberCheck(keyascii)

End Sub

Private Sub txtMVOpTTemperature_LostFocus(Index As Integer)
    Dim msg As String, response As Integer
    Dim Answer As Integer
    Dim IsError As Integer
    Dim ValueChanged As Integer
    Dim EnglishValue As Double, SIValue As Double
    Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtmvOPTTemperature(Index))) Then
     Exit Sub
   End If

   flag_ok = True

    If txtmvOPTTemperature(Index).Text = "" Then
        phprop.MolarVolume.operatingT.input.temperature = -1E+25
        Call LostFocus_Handle(Me, txtmvOPTTemperature(Index), flag_ok)
       Exit Sub
    End If

    Call TextHandleError(IsError, txtmvOPTTemperature(Index), Temp_Text)
    If Not IsError Then
       If Not HaveNumber(CDbl(txtmvOPTTemperature(Index).Text)) Then
          txtmvOPTTemperature(Index).Text = Temp_Text
          txtmvOPTTemperature(Index).SetFocus
          Call LostFocus_Handle(Me, txtmvOPTTemperature(Index), flag_ok)
          Exit Sub
       End If

       Call TextNumberChanged(ValueChanged, txtmvOPTTemperature(Index), Temp_Text)

       If ValueChanged Then
          If CurrentUnits = SIUnits Then
             phprop.MolarVolume.operatingT.input.temperature = CDbl(txtmvOPTTemperature(Index).Text)
          Else
             EnglishValue = CDbl(txtmvOPTTemperature(Index).Text)
             Call TEMPENSI(SIValue, EnglishValue)
             phprop.MolarVolume.operatingT.input.temperature = SIValue
          End If

       Else
        Call LostFocus_Handle(Me, txtmvOPTTemperature(Index), flag_ok)
          Exit Sub
       End If

    End If
        Call LostFocus_Handle(Me, txtmvOPTTemperature(Index), flag_ok)

End Sub

