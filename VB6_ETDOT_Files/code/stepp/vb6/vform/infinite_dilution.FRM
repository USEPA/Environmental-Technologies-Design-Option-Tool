VERSION 5.00
Begin VB.Form Infinite_dilution_form 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Infinite Dilution Activity Coefficient"
   ClientHeight    =   3780
   ClientLeft      =   1530
   ClientTop       =   3900
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
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3780
   ScaleWidth      =   8850
   Begin VB.ComboBox cboUNIFACParameterSet 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   3720
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3240
      Width           =   4932
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      Caption         =   "Accept Selected Activity Coefficient"
      Height          =   615
      Left            =   5040
      TabIndex        =   0
      Top             =   120
      Width           =   3495
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
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   1
      Left            =   1560
      TabIndex        =   2
      Top             =   2400
      Width           =   240
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   1212
      Left            =   1440
      Top             =   1680
      Width           =   5892
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
      TabIndex        =   15
      Top             =   3240
      Width           =   3252
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      Height          =   492
      Left            =   120
      Top             =   3120
      Width           =   8652
   End
   Begin VB.Label lblACTemperature 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   6120
      TabIndex        =   5
      Top             =   2400
      Width           =   972
   End
   Begin VB.Label lblActivityCoefficientValue 
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
      X1              =   1440
      X2              =   7320
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label lblSourceLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "UNIFAC"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   1920
      TabIndex        =   3
      Top             =   2400
      Width           =   1932
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Current Activity Coefficient Information"
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   480
      TabIndex        =   14
      Top             =   240
      Width           =   4092
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
      TabIndex        =   13
      Top             =   720
      Width           =   1932
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
      Width           =   1932
   End
   Begin VB.Label lblCurrentValues 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   2640
      TabIndex        =   11
      Top             =   1080
      Width           =   1932
   End
   Begin VB.Label lblCurrentValues 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   2640
      TabIndex        =   7
      Top             =   720
      Width           =   1932
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   2
      Height          =   1332
      Left            =   360
      Top             =   120
      Width           =   4332
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   4680
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label lblACLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Activity Coefficient"
      ForeColor       =   &H00000000&
      Height          =   400
      Left            =   4080
      TabIndex        =   10
      Top             =   1750
      Width           =   1812
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Source"
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   1920
      TabIndex        =   9
      Top             =   1800
      Width           =   1932
   End
   Begin VB.Label lblACTempLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Temp."
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   6120
      TabIndex        =   8
      Top             =   1800
      Width           =   972
   End
End
Attribute VB_Name = "Infinite_dilution_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BIP_DB_When_ActCoeff_Window_Opened As Integer

Private Sub cboUNIFACParameterSet_Click()
    If Infinite_dilution_form!cboUNIFACParameterSet.ListCount = 1 Then Exit Sub

    Select Case Infinite_dilution_form!cboUNIFACParameterSet.ListIndex
       Case 0   'Original UNIFAC VLE
            If phprop.ActivityCoefficient.BinaryInteractionParameterDatabase <> 1 Then
               If phprop.ActivityCoefficient.BinaryInteractionParameterDBAvailable(1) = False Then
                  MsgBox "Original UNIFAC VLE not available to calculate activity coefficient for this compound.", MB_ICONSTOP, "Data Not Available"
                  Infinite_dilution_form!cboUNIFACParameterSet.ListIndex = phprop.ActivityCoefficient.PreviousBinaryInteractionParameterDB - 1
                  Exit Sub
               End If
               phprop.ActivityCoefficient.PreviousBinaryInteractionParameterDB = phprop.ActivityCoefficient.BinaryInteractionParameterDatabase
               phprop.ActivityCoefficient.BinaryInteractionParameterDatabase = 1
               UserSelectedTheUnifacBIPDBActCoeff = True
               Screen.MousePointer = 11   'Hourglass
               Call CalculateActivityCoefficient
               Screen.MousePointer = 0    'Arrow
            End If
       Case 1   'UNIFAC LLE
               If phprop.ActivityCoefficient.BinaryInteractionParameterDBAvailable(2) = False Then
                  MsgBox "UNIFAC LLE not available to calculate activity coefficient for this compound.", MB_ICONSTOP, "Data Not Available"
                  Infinite_dilution_form!cboUNIFACParameterSet.ListIndex = phprop.ActivityCoefficient.PreviousBinaryInteractionParameterDB - 1
                  Exit Sub
               End If
            If phprop.ActivityCoefficient.BinaryInteractionParameterDatabase <> 2 Then
               phprop.ActivityCoefficient.PreviousBinaryInteractionParameterDB = phprop.ActivityCoefficient.BinaryInteractionParameterDatabase
               phprop.ActivityCoefficient.BinaryInteractionParameterDatabase = 2
               UserSelectedTheUnifacBIPDBActCoeff = True
               Screen.MousePointer = 11  'Hourglass
               Call CalculateActivityCoefficient
               Screen.MousePointer = 0   'Arrow
            End If
       Case 2   'Environmental VLE
               If phprop.ActivityCoefficient.BinaryInteractionParameterDBAvailable(3) = False Then
                  MsgBox "Environmental VLE not available to calculate activity coefficient for this compound.", MB_ICONSTOP, "Data Not Available"
                  Infinite_dilution_form!cboUNIFACParameterSet.ListIndex = phprop.ActivityCoefficient.PreviousBinaryInteractionParameterDB - 1
                  Exit Sub
               End If
            If phprop.ActivityCoefficient.BinaryInteractionParameterDatabase <> 3 Then
               phprop.ActivityCoefficient.PreviousBinaryInteractionParameterDB = phprop.ActivityCoefficient.BinaryInteractionParameterDatabase
               phprop.ActivityCoefficient.BinaryInteractionParameterDatabase = 3
               UserSelectedTheUnifacBIPDBActCoeff = True
               Screen.MousePointer = 11    'Hourglass
               Call CalculateActivityCoefficient
               Screen.MousePointer = 0     'Arrow
            End If
    End Select

End Sub

Private Sub cmdCancel_Click()
    Dim SelectedOption As Integer   'Option selected permanently by the user (the option corresponding to the number on the main screen)

    Select Case phprop.ActivityCoefficient.CurrentSelection.choice
       Case ACTIVITY_COEFFICIENT_UNIFAC
          SelectedOption = 1
       Case Else
          Infinite_dilution_form.Hide
          Exit Sub
    End Select

    If Not Option1(SelectedOption).Value Then Option1(SelectedOption).Value = True

    Infinite_dilution_form.Hide
End Sub

Private Sub cmdok_Click()
    Dim OptionSelected As Integer
    Dim ValueToDisplayIndex As Integer
    Dim i As Integer

'*** Pass new selected value back to main screen
    For i = 1 To 1
        If Option1(i).Value Then
           OptionSelected = i
           Exit For
        End If
    Next i

    Select Case OptionSelected
       Case 1
          ValueToDisplayIndex = ACTIVITY_COEFFICIENT_UNIFAC
    End Select

    If ValueToDisplayIndex <> phprop.ActivityCoefficient.CurrentSelection.choice Then
       phprop.ActivityCoefficient.CurrentSelection.choice = ValueToDisplayIndex
       Call DisplayActivityCoefficientMainScreen(ValueToDisplayIndex)
    End If

    Infinite_dilution_form.Hide

'Recalculate Henry's Constant using new activity coefficient 'if UNIFAC parameter set has changed

          If phprop.ActivityCoefficient.BinaryInteractionParameterDatabase = BIP_DB_When_ActCoeff_Window_Opened Then Exit Sub

'          frmWaitForCalculations.Show
'          frmWaitForCalculations.Refresh

          contam_prop_form!lblContaminantProperties(2).Caption = ""

          Screen.MousePointer = 11   'Hourglass

          Call CalculateHenrysConstant
          contam_prop_form.Refresh

          Screen.MousePointer = 0    'Arrow

'          frmWaitForCalculations.Hide

          NumContaminantInList = contam_prop_form!cboSelectContaminant.ListIndex + 1
          PropContaminant(NumContaminantInList) = phprop

End Sub

Private Sub Form_Activate()

  Call centerform_relative(contam_prop_form, Me)

       BIP_DB_When_ActCoeff_Window_Opened = phprop.ActivityCoefficient.BinaryInteractionParameterDatabase

       Infinite_dilution_form!cboUNIFACParameterSet.Clear
       Infinite_dilution_form!cboUNIFACParameterSet.AddItem "Original UNIFAC VLE"
       Infinite_dilution_form!cboUNIFACParameterSet.AddItem "UNIFAC LLE"
       Infinite_dilution_form!cboUNIFACParameterSet.AddItem "Environmental VLE"
       If Not Infinite_dilution_form!cboUNIFACParameterSet.Enabled Then Infinite_dilution_form!cboUNIFACParameterSet.Enabled = True
    
    Select Case phprop.ActivityCoefficient.BinaryInteractionParameterDatabase
       Case 1
            Infinite_dilution_form!cboUNIFACParameterSet.ListIndex = 0
       Case 2
            Infinite_dilution_form!cboUNIFACParameterSet.ListIndex = 1
       Case 3
            Infinite_dilution_form!cboUNIFACParameterSet.ListIndex = 2
       Case 0
            Infinite_dilution_form!cboUNIFACParameterSet.Clear
            Infinite_dilution_form!cboUNIFACParameterSet.AddItem "UNIFAC Activity Coefficient Calculation Not Possible"
            Infinite_dilution_form!cboUNIFACParameterSet.ListIndex = 0
            Infinite_dilution_form!cboUNIFACParameterSet.Enabled = False
    End Select

End Sub

Private Sub Form_Load()

  Call centerform_relative(contam_prop_form, Me)
    
    If (DemoMode) Then cmdok.Enabled = False


End Sub

Private Sub lblActivityCoefficientValue_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ErrorFlag As Integer

    If Button <> 2 Then Exit Sub

    Select Case Index
       Case 0   'UNIFAC
          ErrorFlag = phprop.ActivityCoefficient.UNIFAC.error
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
    Dim i As Integer

    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.ActivityCoefficient.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

End Sub

Private Sub Option1_Click(Index As Integer)
    Dim i As Integer, SourceIndex As Integer

    SourceIndex = Index - 1
    If SourceIndex = hilight.ActivityCoefficient.PreviousIndex Then Exit Sub
    lblSourceLabel(SourceIndex).BackColor = &H800000
    lblSourceLabel(SourceIndex).ForeColor = &H80000005
    i = hilight.ActivityCoefficient.PreviousIndex
    hilight.ActivityCoefficient.PreviousIndex = SourceIndex
    If i = -1 Then Exit Sub
    If Option1(i + 1).Enabled = False Then Exit Sub
    lblSourceLabel(i).BackColor = &H80000005
    lblSourceLabel(i).ForeColor = &H80000008

End Sub

