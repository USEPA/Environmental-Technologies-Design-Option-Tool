VERSION 5.00
Begin VB.Form aqsol_form 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aqueous Solubility"
   ClientHeight    =   5670
   ClientLeft      =   960
   ClientTop       =   1470
   ClientWidth     =   8850
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
   ScaleHeight     =   5670
   ScaleWidth      =   8850
   Begin VB.ComboBox cboUNIFACParameterSet 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   3720
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   5160
      Width           =   4932
   End
   Begin VB.TextBox txtAqSolTemperature 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   4
      Left            =   6480
      TabIndex        =   9
      Top             =   4320
      Width           =   975
   End
   Begin VB.TextBox txtAqueousSolubilityValue 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   4
      Left            =   4440
      TabIndex        =   8
      Top             =   4320
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   1560
      TabIndex        =   6
      Top             =   4320
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   1560
      TabIndex        =   5
      Top             =   3840
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   1560
      TabIndex        =   4
      Top             =   3360
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   1560
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
      Left            =   1560
      TabIndex        =   2
      Top             =   2400
      Width           =   255
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   615
      Left            =   5040
      TabIndex        =   1
      Top             =   840
      Width           =   3495
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      Caption         =   "Accept Selected Aqueous Solubility"
      Height          =   615
      Left            =   5040
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      Height          =   492
      Left            =   120
      Top             =   5040
      Width           =   8652
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "UNIFAC Group Interaction Database:"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   240
      TabIndex        =   31
      Top             =   5160
      Width           =   3252
   End
   Begin VB.Label lblAqSolTemperature 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   3
      Left            =   6480
      TabIndex        =   30
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label lblAqSolTemperature 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   2
      Left            =   6480
      TabIndex        =   29
      Top             =   3360
      Width           =   972
   End
   Begin VB.Label lblAqSolTemperature 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   6480
      TabIndex        =   28
      Top             =   2880
      Width           =   972
   End
   Begin VB.Label lblAqSolTemperature 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   6480
      TabIndex        =   27
      Top             =   2400
      Width           =   972
   End
   Begin VB.Label lblAqueousSolubilityValue 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   3
      Left            =   4440
      TabIndex        =   26
      Top             =   3840
      Width           =   1812
   End
   Begin VB.Label lblAqueousSolubilityValue 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   2
      Left            =   4440
      TabIndex        =   25
      Top             =   3360
      Width           =   1812
   End
   Begin VB.Label lblAqueousSolubilityValue 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   4440
      TabIndex        =   24
      Top             =   2880
      Width           =   1812
   End
   Begin VB.Label lblAqueousSolubilityValue 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   4440
      TabIndex        =   23
      Top             =   2400
      Width           =   1812
   End
   Begin VB.Line Line2 
      X1              =   1440
      X2              =   7560
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   3132
      Left            =   1440
      Top             =   1680
      Width           =   6132
   End
   Begin VB.Label lblSource 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "User Input"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   4
      Left            =   1920
      TabIndex        =   22
      Top             =   4320
      Width           =   2292
   End
   Begin VB.Label lblSource 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "UNIFAC at Database T"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   3
      Left            =   1920
      TabIndex        =   21
      Top             =   3840
      Width           =   2292
   End
   Begin VB.Label lblSource 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Database"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   2
      Left            =   1920
      TabIndex        =   20
      Top             =   3360
      Width           =   2292
   End
   Begin VB.Label lblSource 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "UNIFAC Fit with Data Pt."
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   1920
      TabIndex        =   19
      Top             =   2400
      Width           =   2292
   End
   Begin VB.Label lblSource 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "UNIFAC at Operating T"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   1920
      TabIndex        =   18
      Top             =   2880
      Width           =   2292
   End
   Begin VB.Label lblAqSolTempLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Temp."
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   6480
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
      Left            =   1920
      TabIndex        =   16
      Top             =   1800
      Width           =   2292
   End
   Begin VB.Label lblAqSolLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Aqueous Sol."
      ForeColor       =   &H00000000&
      Height          =   400
      Left            =   4440
      TabIndex        =   15
      Top             =   1750
      Width           =   1812
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   4680
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   2
      Height          =   1332
      Left            =   360
      Top             =   120
      Width           =   4332
   End
   Begin VB.Label lblCurrentValues 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   2280
      TabIndex        =   14
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
      Left            =   2280
      TabIndex        =   13
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
      Left            =   480
      TabIndex        =   12
      Top             =   1080
      Width           =   1572
   End
   Begin VB.Label lblCurrentInformation 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Value"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   480
      TabIndex        =   11
      Top             =   720
      Width           =   1572
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Current Aqueous Solubility Information"
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   480
      TabIndex        =   10
      Top             =   240
      Width           =   4092
   End
End
Attribute VB_Name = "aqsol_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PastAqSolInputValue As String
Dim PastAqSolInputTemp As String
Dim EFlag As Integer

Private Sub cboUNIFACParameterSet_Click()

    If aqsol_form!cboUNIFACParameterSet.ListCount = 1 Then Exit Sub

    Select Case aqsol_form!cboUNIFACParameterSet.ListIndex
       Case 0   'Original UNIFAC VLE
            If phprop.AqueousSolubility.BinaryInteractionParameterDatabase <> 1 Then
               If phprop.AqueousSolubility.BinaryInteractionParameterDBAvailable(1) = False Then
                  MsgBox "Original UNIFAC VLE not available to calculate aqueous solubility for this compound.", MB_ICONSTOP, "Data Not Available"
                  aqsol_form!cboUNIFACParameterSet.ListIndex = phprop.AqueousSolubility.PreviousBinaryInteractionParameterDB - 1
                  Exit Sub
               End If
               phprop.AqueousSolubility.PreviousBinaryInteractionParameterDB = phprop.AqueousSolubility.BinaryInteractionParameterDatabase
               phprop.AqueousSolubility.BinaryInteractionParameterDatabase = 1
               UserSelectedTheUnifacBIPDBAqSol = True
               Screen.MousePointer = 11   'Hourglass
               Call CalculateAqueousSolubility
               Screen.MousePointer = 0    'Arrow
            End If
       Case 1   'UNIFAC LLE
               If phprop.AqueousSolubility.BinaryInteractionParameterDBAvailable(2) = False Then
                  MsgBox "UNIFAC LLE not available to calculate aqueous solubility for this compound.", MB_ICONSTOP, "Data Not Available"
                  aqsol_form!cboUNIFACParameterSet.ListIndex = phprop.AqueousSolubility.PreviousBinaryInteractionParameterDB - 1
                  Exit Sub
               End If
            If phprop.AqueousSolubility.BinaryInteractionParameterDatabase <> 2 Then
               phprop.AqueousSolubility.PreviousBinaryInteractionParameterDB = phprop.AqueousSolubility.BinaryInteractionParameterDatabase
               phprop.AqueousSolubility.BinaryInteractionParameterDatabase = 2
               UserSelectedTheUnifacBIPDBAqSol = True
               Screen.MousePointer = 11  'hourglass
               Call CalculateAqueousSolubility
               Screen.MousePointer = 0   'arrow
            End If
       Case 2   'Environmental VLE
               If phprop.AqueousSolubility.BinaryInteractionParameterDBAvailable(3) = False Then
                  MsgBox "Environmental VLE not available to calculate aqueous solubility for this compound.", MB_ICONSTOP, "Data Not Available"
                  aqsol_form!cboUNIFACParameterSet.ListIndex = phprop.AqueousSolubility.PreviousBinaryInteractionParameterDB - 1
                  Exit Sub
               End If
            If phprop.AqueousSolubility.BinaryInteractionParameterDatabase <> 3 Then
               phprop.AqueousSolubility.PreviousBinaryInteractionParameterDB = phprop.AqueousSolubility.BinaryInteractionParameterDatabase
               phprop.AqueousSolubility.BinaryInteractionParameterDatabase = 3
               UserSelectedTheUnifacBIPDBAqSol = True
               Screen.MousePointer = 11    'Hourglass
               Call CalculateAqueousSolubility
               Screen.MousePointer = 0     'Arrow
            End If
    End Select

End Sub

Private Sub cmdCancel_Click()
    Dim SelectedOption As Integer   'Option selected permanently by the user (the option corresponding to the number on the main screen)

    Select Case phprop.AqueousSolubility.CurrentSelection.choice
       Case AQUEOUS_SOLUBILITY_FIT
          SelectedOption = 1
       Case AQUEOUS_SOLUBILITY_OPT_UNIFAC
          SelectedOption = 2
       Case AQUEOUS_SOLUBILITY_DATABASE
          SelectedOption = 3
       Case AQUEOUS_SOLUBILITY_DBT_UNIFAC
          SelectedOption = 4
       Case AQUEOUS_SOLUBILITY_INPUT
          SelectedOption = 5
          txtaqueoussolubilityValue(4).Text = PastAqSolInputValue
          txtaqSoltemperature(4).Text = PastAqSolInputTemp
       Case Else
          aqsol_form.Hide
          Exit Sub
    End Select

    If Not Option1(SelectedOption).Value Then Option1(SelectedOption).Value = True

    aqsol_form.Hide
End Sub

Private Sub cmdok_Click()
    Dim OptionSelected As Integer
    Dim ValueToDisplayIndex As Integer
    Dim i As Integer

'*** Pass new selected value back to main screen
    For i = 1 To 5
        If Option1(i).Value Then
           OptionSelected = i
           Exit For
        End If
    Next i

    Select Case OptionSelected
       Case 1
          ValueToDisplayIndex = AQUEOUS_SOLUBILITY_FIT
       Case 2
          ValueToDisplayIndex = AQUEOUS_SOLUBILITY_OPT_UNIFAC
       Case 3
          ValueToDisplayIndex = AQUEOUS_SOLUBILITY_DATABASE
       Case 4
          ValueToDisplayIndex = AQUEOUS_SOLUBILITY_DBT_UNIFAC
       Case 5
          If Not PROPAVAILABLE(AQUEOUS_SOLUBILITY_INPUT) Then
             MsgBox "User Input can not be selected without first entering a value", MB_ICONSTOP, "Error"
             txtaqueoussolubilityValue(4).SetFocus
             Exit Sub
          End If
          ValueToDisplayIndex = AQUEOUS_SOLUBILITY_INPUT
    End Select

    If ValueToDisplayIndex <> phprop.AqueousSolubility.CurrentSelection.choice Then
       phprop.AqueousSolubility.CurrentSelection.choice = ValueToDisplayIndex
       Call DisplayAqueousSolubilityMainScreen(ValueToDisplayIndex)
    ElseIf ValueToDisplayIndex = AQUEOUS_SOLUBILITY_INPUT Then
       Call DisplayAqueousSolubilityMainScreen(ValueToDisplayIndex)
    End If

    aqsol_form.Hide
End Sub

Private Sub Form_Activate()
    
  Call centerform_relative(contam_prop_form, Me)
    
    PastAqSolInputValue = txtaqueoussolubilityValue(4).Text
    PastAqSolInputTemp = txtaqSoltemperature(4).Text

       aqsol_form!cboUNIFACParameterSet.Clear
       aqsol_form!cboUNIFACParameterSet.AddItem "Original UNIFAC VLE"
       aqsol_form!cboUNIFACParameterSet.AddItem "UNIFAC LLE"
       aqsol_form!cboUNIFACParameterSet.AddItem "Environmental VLE"
       If Not aqsol_form!cboUNIFACParameterSet.Enabled Then aqsol_form!cboUNIFACParameterSet.Enabled = True
    
    Select Case phprop.AqueousSolubility.BinaryInteractionParameterDatabase
       Case 1
            aqsol_form!cboUNIFACParameterSet.ListIndex = 0
       Case 2
            aqsol_form!cboUNIFACParameterSet.ListIndex = 1
       Case 3
            aqsol_form!cboUNIFACParameterSet.ListIndex = 2
       Case 0
            aqsol_form!cboUNIFACParameterSet.Clear
            aqsol_form!cboUNIFACParameterSet.AddItem "UNIFAC Solubility Calculation Not Possible"
            aqsol_form!cboUNIFACParameterSet.ListIndex = 0
            aqsol_form!cboUNIFACParameterSet.Enabled = False
    End Select
    
End Sub

Private Sub Form_Load()

  Call centerform_relative(contam_prop_form, Me)

    If (DemoMode) Then cmdok.Enabled = False

    EFlag = False


End Sub

Private Sub lblAqSolTemperature_Click(Index As Integer)
    
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.AqueousSolubility.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

End Sub

Private Sub lblAqueousSolubilityValue_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim ErrorFlag As Integer
    
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.AqueousSolubility.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

    If Button <> 2 Then Exit Sub

    Select Case Index
       Case 0  'Source = UNIFAC Fit with a Data Point
          ErrorFlag = phprop.AqueousSolubility.fit.UNIFAC.error
       Case 1  'Source = UNIFAC at Operating T
          ErrorFlag = phprop.AqueousSolubility.operatingT.UNIFAC.error
       Case 2  'Source = Database
          ErrorFlag = phprop.AqueousSolubility.database.error
       Case 3  'Source = UNIFAC at Database T
          ErrorFlag = phprop.AqueousSolubility.UNIFAC.error
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
       If Index = hilight.AqueousSolubility.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If
    
End Sub

Private Sub Option1_Click(Index As Integer)
    Dim i As Integer, SourceIndex As Integer
    
    SourceIndex = Index - 1
    If SourceIndex = hilight.AqueousSolubility.PreviousIndex Then Exit Sub
    lblSource(SourceIndex).BackColor = &H800000
    lblSource(SourceIndex).ForeColor = &H80000005
    i = hilight.AqueousSolubility.PreviousIndex
    hilight.AqueousSolubility.PreviousIndex = SourceIndex
    If i = -1 Then Exit Sub
    If Option1(i + 1).Enabled = False Then Exit Sub
    lblSource(i).BackColor = &H80000005
    lblSource(i).ForeColor = &H80000008

End Sub

Private Sub txtAqSolTemperature_GotFocus(Index As Integer)
  Call gotfocus_handle(Me, txtaqSoltemperature(Index), Temp_Text)

    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.AqueousSolubility.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

End Sub

Private Sub txtAqSolTemperature_KeyPress(Index As Integer, keyascii As Integer)
    
    If keyascii = 13 Then
       keyascii = 0
       cmdok.SetFocus
       For i = 1 To 4
          Option1(i).Enabled = True
       Next i
       cboUNIFACParameterSet.Enabled = True
       EFlag = False
       Exit Sub
    End If
    Call NumberCheck(keyascii)
   
End Sub

Private Sub txtAqSolTemperature_LostFocus(Index As Integer)
    Dim msg As String, response As Integer
    Dim Answer As Integer
    Dim IsError As Integer
    Dim ValueChanged As Integer
    Dim EnglishValue As Double, SIValue As Double
    Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtaqSoltemperature(Index))) Then
     Exit Sub
   End If

   flag_ok = True

    If txtaqSoltemperature(Index).Text = "" Then
       phprop.AqueousSolubility.input.temperature = -1E+25
       Call LostFocus_Handle(Me, txtaqSoltemperature(Index), flag_ok)
       Exit Sub
    End If

    Call TextHandleError(IsError, txtaqSoltemperature(Index), Temp_Text)
    If Not IsError Then
       If Not HaveNumber(CDbl(txtaqSoltemperature(Index).Text)) Then
          txtaqSoltemperature(Index).Text = Temp_Text
          txtaqSoltemperature(Index).SetFocus
          Call LostFocus_Handle(Me, txtaqSoltemperature(Index), flag_ok)
          Exit Sub
       End If

       Call TextNumberChanged(ValueChanged, txtaqSoltemperature(Index), Temp_Text)

       If ValueChanged Then
          If CurrentUnits = SIUnits Then
             phprop.AqueousSolubility.input.temperature = CDbl(txtaqSoltemperature(Index).Text)
          Else
             EnglishValue = CDbl(txtaqSoltemperature(Index).Text)
             Call TEMPENSI(SIValue, EnglishValue)
             phprop.AqueousSolubility.input.temperature = SIValue
          End If

          
       Else
          Call LostFocus_Handle(Me, txtaqSoltemperature(Index), flag_ok)
          Exit Sub
       End If

    End If
  Call LostFocus_Handle(Me, txtaqSoltemperature(Index), flag_ok)

End Sub

Private Sub txtAqueousSolubilityValue_GotFocus(Index As Integer)
  Call gotfocus_handle(Me, txtaqueoussolubilityValue(Index), Temp_Text)

    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.AqueousSolubility.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If
End Sub

Private Sub txtAqueousSolubilityValue_KeyPress(Index As Integer, keyascii As Integer)
    
    If keyascii = 13 Then
       keyascii = 0
       txtaqSoltemperature(Index).SetFocus
       Exit Sub
    End If
    Call NumberCheck(keyascii)

End Sub

Private Sub txtAqueousSolubilityValue_LostFocus(Index As Integer)
    Dim msg As String, response As Integer
    Dim Answer As Integer
    Dim IsError As Integer
    Dim ValueChanged As Integer
    Dim EnglishValue As Double, SIValue As Double
    Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtaqueoussolubilityValue(Index))) Then
     Exit Sub
   End If

   flag_ok = True
    If txtaqueoussolubilityValue(Index).Text = "" Then
       phprop.AqueousSolubility.input.Value = -1#
       PROPAVAILABLE(AQUEOUS_SOLUBILITY_INPUT) = False
       Call LostFocus_Handle(Me, txtaqueoussolubilityValue(Index), flag_ok)
       Exit Sub
    End If

    Call TextHandleError(IsError, txtaqueoussolubilityValue(Index), Temp_Text)
    If Not IsError Then
       If Not HaveNumber(CDbl(txtaqueoussolubilityValue(Index).Text)) Then
          txtaqueoussolubilityValue(Index).Text = Temp_Text
          txtaqueoussolubilityValue(Index).SetFocus
          Call LostFocus_Handle(Me, txtaqueoussolubilityValue(Index), flag_ok)
          Exit Sub
       End If

       Call TextNumberChanged(ValueChanged, txtaqueoussolubilityValue(Index), Temp_Text)

       If ValueChanged Then
          If CurrentUnits = SIUnits Then
             phprop.AqueousSolubility.input.Value = CDbl(txtaqueoussolubilityValue(Index).Text)
          Else
             EnglishValue = CDbl(txtaqueoussolubilityValue(Index).Text)
             Call AQSENSI(SIValue, EnglishValue)
             phprop.AqueousSolubility.input.Value = SIValue
          End If

          
          PROPAVAILABLE(AQUEOUS_SOLUBILITY_INPUT) = True
       Else
          Call LostFocus_Handle(Me, txtaqueoussolubilityValue(Index), flag_ok)
          Exit Sub
       End If
    End If
  Call LostFocus_Handle(Me, txtaqueoussolubilityValue(Index), flag_ok)


End Sub

