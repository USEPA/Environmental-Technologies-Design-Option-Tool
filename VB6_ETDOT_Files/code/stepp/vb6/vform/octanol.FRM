VERSION 5.00
Begin VB.Form octanol_form 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Octanol Water Partition Coefficient (Kow)"
   ClientHeight    =   4950
   ClientLeft      =   1350
   ClientTop       =   1590
   ClientWidth     =   8880
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
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4950
   ScaleWidth      =   8880
   Begin VB.ComboBox cboUNIFACParameterSet 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   3720
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   4440
      Width           =   4932
   End
   Begin VB.TextBox txtOWPCTemperature 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   6360
      TabIndex        =   13
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox txtOctWatPartCoeffValue 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   4320
      TabIndex        =   12
      Top             =   3840
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   1320
      TabIndex        =   11
      Top             =   3840
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   1320
      TabIndex        =   8
      Top             =   3360
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   1320
      TabIndex        =   5
      Top             =   2880
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   1320
      TabIndex        =   2
      Top             =   2400
      Width           =   255
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      Caption         =   "Accept Selected Partition Coefficient"
      Height          =   615
      Left            =   5160
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   615
      Left            =   5160
      TabIndex        =   1
      Top             =   840
      Width           =   3495
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
      TabIndex        =   27
      Top             =   4440
      Width           =   3252
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      Height          =   492
      Left            =   120
      Top             =   4320
      Width           =   8652
   End
   Begin VB.Label lblOWPCTemperature 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   2
      Left            =   6360
      TabIndex        =   10
      Top             =   3360
      Width           =   1212
   End
   Begin VB.Label lblOWPCTemperature 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   6360
      TabIndex        =   7
      Top             =   2880
      Width           =   1212
   End
   Begin VB.Label lblOWPCTemperature 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   6360
      TabIndex        =   4
      Top             =   2400
      Width           =   1212
   End
   Begin VB.Label lblOctWatPartCoeffValue 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   2
      Left            =   4320
      TabIndex        =   9
      Top             =   3360
      Width           =   1812
   End
   Begin VB.Label lblOctWatPartCoeffValue 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   4320
      TabIndex        =   6
      Top             =   2880
      Width           =   1812
   End
   Begin VB.Label lblOctWatPartCoeffValue 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   4320
      TabIndex        =   3
      Top             =   2400
      Width           =   1812
   End
   Begin VB.Label lblSource 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "UNIFAC @ Database T"
      ForeColor       =   &H00000000&
      Height          =   252
      Index           =   2
      Left            =   1680
      TabIndex        =   26
      Top             =   3360
      Width           =   2412
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Current Oct. Water Part. Coeff. Information"
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   480
      TabIndex        =   25
      Top             =   240
      Width           =   4212
   End
   Begin VB.Label lblCurrentInformation 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Value of log Kow"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   480
      TabIndex        =   24
      Top             =   720
      Width           =   1692
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
      TabIndex        =   23
      Top             =   1080
      Width           =   1692
   End
   Begin VB.Label lblCurrentValues 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   2400
      TabIndex        =   22
      Top             =   1080
      Width           =   2292
   End
   Begin VB.Label lblCurrentValues 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   2400
      TabIndex        =   21
      Top             =   720
      Width           =   2292
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   2
      Height          =   1332
      Left            =   360
      Top             =   120
      Width           =   4452
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   4800
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label lblKowLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "log Kow"
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   4320
      TabIndex        =   20
      Top             =   1800
      Width           =   1812
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Source"
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   1680
      TabIndex        =   19
      Top             =   1800
      Width           =   2412
   End
   Begin VB.Label lblSource 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "UNIFAC @ Operating T"
      ForeColor       =   &H00000000&
      Height          =   252
      Index           =   0
      Left            =   1680
      TabIndex        =   18
      Top             =   2400
      Width           =   2412
   End
   Begin VB.Label lblSource 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Database"
      ForeColor       =   &H00000000&
      Height          =   252
      Index           =   1
      Left            =   1680
      TabIndex        =   17
      Top             =   2880
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
      Index           =   3
      Left            =   1680
      TabIndex        =   16
      Top             =   3840
      Width           =   2412
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   2532
      Left            =   1200
      Top             =   1680
      Width           =   6492
   End
   Begin VB.Line Line2 
      X1              =   1200
      X2              =   7680
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label lblKowTempLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Temp."
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   6360
      TabIndex        =   15
      Top             =   1800
      Width           =   1212
   End
End
Attribute VB_Name = "octanol_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PastKowInputValue As String
Dim PastKowInputTemp As String

Private Sub cboUNIFACParameterSet_Click()

    If octanol_form!cboUNIFACParameterSet.ListCount = 1 Then Exit Sub

    Select Case octanol_form!cboUNIFACParameterSet.ListIndex
       Case 0   'Original UNIFAC VLE
            If phprop.OctWaterPartCoeff.BinaryInteractionParameterDatabase <> 1 Then
               If phprop.OctWaterPartCoeff.BinaryInteractionParameterDBAvailable(1) = False Then
                  MsgBox "Original UNIFAC VLE not available to calculate octanol water partition coefficient for this compound.", MB_ICONSTOP, "Data Not Available"
                  octanol_form!cboUNIFACParameterSet.ListIndex = phprop.OctWaterPartCoeff.PreviousBinaryInteractionParameterDB - 1
                  Exit Sub
               End If
               phprop.OctWaterPartCoeff.PreviousBinaryInteractionParameterDB = phprop.OctWaterPartCoeff.BinaryInteractionParameterDatabase
               phprop.OctWaterPartCoeff.BinaryInteractionParameterDatabase = 1
               UserSelectedTheUnifacBIPDBKow = True
               Screen.MousePointer = 11   'Hourglass
               Call CalculateOctWaterPartCoeff
               Screen.MousePointer = 0    'Arrow
            End If
       Case 1   'UNIFAC LLE
               If phprop.OctWaterPartCoeff.BinaryInteractionParameterDBAvailable(2) = False Then
                  MsgBox "UNIFAC LLE not available to calculate octanol water partition coefficient for this compound.", MB_ICONSTOP, "Data Not Available"
                  octanol_form!cboUNIFACParameterSet.ListIndex = phprop.OctWaterPartCoeff.PreviousBinaryInteractionParameterDB - 1
                  Exit Sub
               End If
            If phprop.OctWaterPartCoeff.BinaryInteractionParameterDatabase <> 2 Then
               phprop.OctWaterPartCoeff.PreviousBinaryInteractionParameterDB = phprop.OctWaterPartCoeff.BinaryInteractionParameterDatabase
               phprop.OctWaterPartCoeff.BinaryInteractionParameterDatabase = 2
               UserSelectedTheUnifacBIPDBKow = True
               Screen.MousePointer = 11  'hourglass
               Call CalculateOctWaterPartCoeff
               Screen.MousePointer = 0   'arrow
            End If
    End Select

End Sub

Private Sub cmdCancel_Click()
    Dim SelectedOption As Integer   'Option selected permanently by the user (the option corresponding to the number on the main screen)

    Select Case phprop.OctWaterPartCoeff.CurrentSelection.choice
       Case OCT_WATER_PART_COEFF_OPT_UNIFAC
          SelectedOption = 1
       Case OCT_WATER_PART_COEFF_DB
          SelectedOption = 2
       Case OCT_WATER_PART_COEFF_DBT_UNIFAC
          SelectedOption = 3
       Case OCT_WATER_PART_COEFF_INPUT
          SelectedOption = 4
          txtoctwatpartcoeffvalue(3).Text = PastKowInputValue
          txtOWPCTemperature(3).Text = PastKowInputTemp
       Case Else
          octanol_form.Hide
          Exit Sub
    End Select

    If Not Option1(SelectedOption).Value Then Option1(SelectedOption).Value = True

    octanol_form.Hide
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
          ValueToDisplayIndex = OCT_WATER_PART_COEFF_OPT_UNIFAC
       Case 2
          ValueToDisplayIndex = OCT_WATER_PART_COEFF_DB
       Case 3
          ValueToDisplayIndex = OCT_WATER_PART_COEFF_DBT_UNIFAC
       Case 4
          If Not PROPAVAILABLE(OCT_WATER_PART_COEFF_INPUT) Then
             MsgBox "User Input can not be selected without first entering a value", MB_ICONSTOP, "Error"
             txtoctwatpartcoeffvalue(3).SetFocus
             Exit Sub
          End If
          ValueToDisplayIndex = OCT_WATER_PART_COEFF_INPUT
    End Select

    If ValueToDisplayIndex <> phprop.OctWaterPartCoeff.CurrentSelection.choice Then
       phprop.OctWaterPartCoeff.CurrentSelection.choice = ValueToDisplayIndex
       Call DisplayOctWaterPartCoeffMainScreen(ValueToDisplayIndex)
    ElseIf ValueToDisplayIndex = OCT_WATER_PART_COEFF_INPUT Then
       Call DisplayOctWaterPartCoeffMainScreen(ValueToDisplayIndex)
    End If

    octanol_form.Hide
End Sub

Private Sub Form_Activate()
    
  Call centerform_relative(contam_prop_form, Me)
    
    PastKowInputValue = txtoctwatpartcoeffvalue(3).Text
    PastKowInputTemp = txtOWPCTemperature(3).Text

       octanol_form!cboUNIFACParameterSet.Clear
       octanol_form!cboUNIFACParameterSet.AddItem "Original UNIFAC VLE"
       octanol_form!cboUNIFACParameterSet.AddItem "UNIFAC LLE"

       If Not octanol_form!cboUNIFACParameterSet.Enabled Then octanol_form!cboUNIFACParameterSet.Enabled = True
    
    Select Case phprop.OctWaterPartCoeff.BinaryInteractionParameterDatabase
       Case 1
            octanol_form!cboUNIFACParameterSet.ListIndex = 0
       Case 2
            octanol_form!cboUNIFACParameterSet.ListIndex = 1
       Case 3
            octanol_form!cboUNIFACParameterSet.ListIndex = 2
       Case 0
            octanol_form!cboUNIFACParameterSet.Clear
            octanol_form!cboUNIFACParameterSet.AddItem "UNIFAC Kow Calculation Not Possible"
            octanol_form!cboUNIFACParameterSet.ListIndex = 0
            octanol_form!cboUNIFACParameterSet.Enabled = False
    End Select

End Sub

Private Sub Form_Load()

  Call centerform_relative(contam_prop_form, Me)

    If (DemoMode) Then cmdok.Enabled = False


End Sub

Private Sub lblOctWatPartCoeffValue_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ErrorFlag As Integer
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.OctWaterPartCoeff.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

    If Button <> 2 Then Exit Sub

    Select Case Index
       Case 0  'Source = UNIFAC at Operating T
          ErrorFlag = phprop.OctWaterPartCoeff.operatingT.UNIFAC.error
       Case 1  'Source = Database
          ErrorFlag = phprop.OctWaterPartCoeff.database.error
       Case 2  'Source = UNIFAC at Database T
          ErrorFlag = phprop.OctWaterPartCoeff.databaseT.UNIFAC.error
    End Select

    If ErrorFlag = 0 Then Exit Sub

    If ErrorFlag < 0 Then
       MsgBox ErrorMsg(ErrorFlag), MB_ICONINFORMATION, Trim$(phprop.Name) & " - Data Unavailable"
    End If

    If ErrorFlag > 0 Then
       MsgBox ErrorMsg(ErrorFlag), MB_ICONINFORMATION, Trim$(phprop.Name) & " - Warning"
    End If

End Sub

Private Sub lblOWPCTemperature_Click(Index As Integer)
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.OctWaterPartCoeff.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

End Sub

Private Sub lblSource_Click(Index As Integer)
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.OctWaterPartCoeff.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

End Sub

Private Sub Option1_Click(Index As Integer)
    Dim i As Integer, SourceIndex As Integer

    SourceIndex = Index - 1
    If SourceIndex = hilight.OctWaterPartCoeff.PreviousIndex Then Exit Sub
    lblSource(SourceIndex).BackColor = &H800000
    lblSource(SourceIndex).ForeColor = &H80000005
    i = hilight.OctWaterPartCoeff.PreviousIndex
    hilight.OctWaterPartCoeff.PreviousIndex = SourceIndex
    If i = -1 Then Exit Sub
    If Option1(i + 1).Enabled = False Then Exit Sub
    lblSource(i).BackColor = &H80000005
    lblSource(i).ForeColor = &H80000008

End Sub

Private Sub txtOctWatPartCoeffValue_GotFocus(Index As Integer)
  Call gotfocus_handle(Me, txtoctwatpartcoeffvalue(Index), Temp_Text)
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.OctWaterPartCoeff.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

End Sub

Private Sub txtOctWatPartCoeffValue_KeyPress(Index As Integer, keyascii As Integer)
    
    If keyascii = 13 Then
       keyascii = 0
       txtOWPCTemperature(Index).SetFocus
       Exit Sub
    End If
    Call NumberCheck(keyascii)
    
End Sub

Private Sub txtOctWatPartCoeffValue_LostFocus(Index As Integer)
    Dim msg As String, response As Integer
    Dim Answer As Integer
    Dim IsError As Integer
    Dim ValueChanged As Integer
    Dim EnglishValue As Double, SIValue As Double
    Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtoctwatpartcoeffvalue(Index))) Then
     Exit Sub
   End If

   flag_ok = True

    If txtoctwatpartcoeffvalue(Index).Text = "" Then
       phprop.OctWaterPartCoeff.input.Value = -1#
       PROPAVAILABLE(OCT_WATER_PART_COEFF_INPUT) = False
         Call LostFocus_Handle(Me, txtoctwatpartcoeffvalue(Index), flag_ok)
       Exit Sub
    End If

    Call TextHandleError(IsError, txtoctwatpartcoeffvalue(Index), Temp_Text)
    If Not IsError Then
       If Not HaveNumber2(CDbl(txtoctwatpartcoeffvalue(Index).Text)) Then
          txtoctwatpartcoeffvalue(Index).Text = Temp_Text
          txtoctwatpartcoeffvalue(Index).SetFocus
         Call LostFocus_Handle(Me, txtoctwatpartcoeffvalue(Index), flag_ok)
          Exit Sub
       End If

       Call TextNumberChanged(ValueChanged, txtoctwatpartcoeffvalue(Index), Temp_Text)

       If ValueChanged Then
          If CurrentUnits = SIUnits Then
             phprop.OctWaterPartCoeff.input.Value = CDbl(txtoctwatpartcoeffvalue(Index).Text)
          Else
             EnglishValue = CDbl(txtoctwatpartcoeffvalue(Index).Text)
             Call KOWENSI(SIValue, EnglishValue)
             phprop.OctWaterPartCoeff.input.Value = SIValue
          End If
          
          PROPAVAILABLE(OCT_WATER_PART_COEFF_INPUT) = True
       Else
         Call LostFocus_Handle(Me, txtoctwatpartcoeffvalue(Index), flag_ok)
          Exit Sub
       End If
    End If
         Call LostFocus_Handle(Me, txtoctwatpartcoeffvalue(Index), flag_ok)

End Sub

Private Sub txtOWPCTemperature_GotFocus(Index As Integer)

  Call gotfocus_handle(Me, txtOWPCTemperature(Index), Temp_Text)
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.OctWaterPartCoeff.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If
End Sub

Private Sub txtOWPCTemperature_KeyPress(Index As Integer, keyascii As Integer)
    
    If keyascii = 13 Then
       keyascii = 0
       cmdok.SetFocus
       Exit Sub
    End If
    Call NumberCheck(keyascii)
  
End Sub

Private Sub txtOWPCTemperature_LostFocus(Index As Integer)
    Dim msg As String, response As Integer
    Dim Answer As Integer
    Dim IsError As Integer
    Dim ValueChanged As Integer
    Dim EnglishValue As Double, SIValue As Double
    Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtOWPCTemperature(Index))) Then
     Exit Sub
   End If

   flag_ok = True

    If txtOWPCTemperature(Index).Text = "" Then
       phprop.OctWaterPartCoeff.input.temperature = -1E+25
       Call LostFocus_Handle(Me, txtOWPCTemperature(Index), flag_ok)
       Exit Sub
    End If

    Call TextHandleError(IsError, txtOWPCTemperature(Index), Temp_Text)
    If Not IsError Then
       If Not HaveNumber(CDbl(txtOWPCTemperature(Index).Text)) Then
          txtOWPCTemperature(Index).Text = Temp_Text
          txtOWPCTemperature(Index).SetFocus
       Call LostFocus_Handle(Me, txtOWPCTemperature(Index), flag_ok)
          Exit Sub
       End If

       Call TextNumberChanged(ValueChanged, txtOWPCTemperature(Index), Temp_Text)

       If ValueChanged Then
          If CurrentUnits = SIUnits Then
             phprop.OctWaterPartCoeff.input.temperature = CDbl(txtOWPCTemperature(Index).Text)
          Else
             EnglishValue = CDbl(txtOWPCTemperature(Index).Text)
             Call TEMPENSI(SIValue, EnglishValue)
             phprop.OctWaterPartCoeff.input.temperature = SIValue
          End If

       Else
       Call LostFocus_Handle(Me, txtOWPCTemperature(Index), flag_ok)
          Exit Sub
       End If

    End If
       Call LostFocus_Handle(Me, txtOWPCTemperature(Index), flag_ok)

End Sub

