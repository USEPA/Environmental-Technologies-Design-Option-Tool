VERSION 5.00
Begin VB.Form rindex_form 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Refractive Index"
   ClientHeight    =   3480
   ClientLeft      =   645
   ClientTop       =   1695
   ClientWidth     =   8460
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
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3480
   ScaleWidth      =   8460
   Begin VB.TextBox txtRefractiveIndexValue 
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   1
      Left            =   4800
      TabIndex        =   5
      Top             =   2880
      Width           =   1692
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      Caption         =   "Accept Selected Refractive Index"
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
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   2
      Left            =   1800
      TabIndex        =   4
      Top             =   2880
      Width           =   225
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   1800
      TabIndex        =   2
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label lblRefractiveIndexValue 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   4800
      TabIndex        =   3
      Top             =   2400
      Width           =   1692
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Current Refractive Index Information"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   14
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
      TabIndex        =   13
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
      TabIndex        =   12
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
      TabIndex        =   11
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
      TabIndex        =   10
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
   Begin VB.Label lblRefIndexLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Refractive Index @25 C"
      ForeColor       =   &H00000000&
      Height          =   470
      Left            =   4800
      TabIndex        =   9
      Top             =   1800
      Width           =   1692
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Source"
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   2160
      TabIndex        =   8
      Top             =   1800
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
      Left            =   2160
      TabIndex        =   7
      Top             =   2880
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
      Index           =   0
      Left            =   2160
      TabIndex        =   6
      Top             =   2400
      Width           =   2412
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   1692
      Left            =   1680
      Top             =   1680
      Width           =   5052
   End
   Begin VB.Line Line2 
      X1              =   1680
      X2              =   6720
      Y1              =   2280
      Y2              =   2280
   End
End
Attribute VB_Name = "rindex_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PastRefIndInputValue As String

Private Sub cmdCancel_Click()
    Dim SelectedOption As Integer   'Option selected permanently by the user (the option corresponding to the number on the main screen)

    Select Case phprop.RefractiveIndex.CurrentSelection.choice
       Case REFRACTIVE_INDEX_DATABASE
          SelectedOption = 1
       Case REFRACTIVE_INDEX_INPUT
          SelectedOption = 2
          txtRefractiveIndexValue(1).Text = PastRefIndInputValue
       Case Else
          rindex_form.Hide
          Exit Sub
    End Select

    If Not Option1(SelectedOption).Value Then Option1(SelectedOption).Value = True

    rindex_form.Hide
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
          ValueToDisplayIndex = REFRACTIVE_INDEX_DATABASE
       Case 2
          If Not PROPAVAILABLE(REFRACTIVE_INDEX_INPUT) Then
             MsgBox "User Input can not be selected without first entering a value", MB_ICONSTOP, "Error"
             txtRefractiveIndexValue(1).SetFocus
             Exit Sub
          End If
          ValueToDisplayIndex = REFRACTIVE_INDEX_INPUT
    End Select

    If ValueToDisplayIndex <> phprop.RefractiveIndex.CurrentSelection.choice Then
       phprop.RefractiveIndex.CurrentSelection.choice = ValueToDisplayIndex
       Call DisplayRefractiveIndexMainScreen(ValueToDisplayIndex)
    ElseIf ValueToDisplayIndex = REFRACTIVE_INDEX_INPUT Then
       Call DisplayRefractiveIndexMainScreen(ValueToDisplayIndex)
    End If

    rindex_form.Hide

End Sub

Private Sub Form_Activate()
    
  Call centerform_relative(contam_prop_form, Me)
    
    PastRefIndInputValue = txtRefractiveIndexValue(1).Text
End Sub

Private Sub Form_Load()

  Call centerform_relative(contam_prop_form, Me)
    
  If (DemoMode) Then cmdok.Enabled = False


End Sub

Private Sub lblRefractiveIndexValue_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ErrorFlag As Integer
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.RefractiveIndex.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

    If Button <> 2 Then Exit Sub

    Select Case Index
       Case 0   'Database
          ErrorFlag = phprop.RefractiveIndex.database.error
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
       If Index = hilight.RefractiveIndex.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

End Sub

Private Sub Option1_Click(Index As Integer)
    Dim i As Integer, SourceIndex As Integer

    SourceIndex = Index - 1
    If SourceIndex = hilight.RefractiveIndex.PreviousIndex Then Exit Sub
    lblSource(SourceIndex).BackColor = &H800000
    lblSource(SourceIndex).ForeColor = &H80000005
    i = hilight.RefractiveIndex.PreviousIndex
    hilight.RefractiveIndex.PreviousIndex = SourceIndex
    If i = -1 Then Exit Sub
    If Option1(i + 1).Enabled = False Then Exit Sub
    lblSource(i).BackColor = &H80000005
    lblSource(i).ForeColor = &H80000008

End Sub

Private Sub txtRefractiveIndexValue_GotFocus(Index As Integer)
  Call gotfocus_handle(Me, txtRefractiveIndexValue(Index), Temp_Text)
    
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.RefractiveIndex.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If
End Sub

Private Sub txtRefractiveIndexValue_KeyPress(Index As Integer, keyascii As Integer)
    
    If keyascii = 13 Then
       keyascii = 0
       cmdok.SetFocus
       Exit Sub
    End If
    Call NumberCheck(keyascii)
   
End Sub

Private Sub txtRefractiveIndexValue_LostFocus(Index As Integer)
    Dim msg As String, response As Integer
    Dim Answer As Integer
    Dim IsError As Integer
    Dim ValueChanged As Integer
    Dim EnglishValue As Double, SIValue As Double
    Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtRefractiveIndexValue(Index))) Then
     Exit Sub
   End If

   flag_ok = True

    If txtRefractiveIndexValue(Index).Text = "" Then
       phprop.RefractiveIndex.input.Value = -1E+25
       PROPAVAILABLE(REFRACTIVE_INDEX_INPUT) = False
       Call LostFocus_Handle(Me, txtRefractiveIndexValue(Index), flag_ok)
       Exit Sub
    End If

    Call TextHandleError(IsError, txtRefractiveIndexValue(Index), Temp_Text)
    If Not IsError Then
       If Not HaveNumber(CDbl(txtRefractiveIndexValue(Index).Text)) Then
          txtRefractiveIndexValue(Index).Text = Temp_Text
          txtRefractiveIndexValue(Index).SetFocus
       Call LostFocus_Handle(Me, txtRefractiveIndexValue(Index), flag_ok)
          Exit Sub
       End If

       Call TextNumberChanged(ValueChanged, txtRefractiveIndexValue(Index), Temp_Text)

       If ValueChanged Then
          If CurrentUnits = SIUnits Then
             phprop.RefractiveIndex.input.Value = CDbl(txtRefractiveIndexValue(Index).Text)
          Else
             EnglishValue = CDbl(txtRefractiveIndexValue(Index).Text)
             Call RIENSI(SIValue, EnglishValue)
             phprop.RefractiveIndex.input.Value = SIValue
          End If

          
          PROPAVAILABLE(REFRACTIVE_INDEX_INPUT) = True
       Else
       Call LostFocus_Handle(Me, txtRefractiveIndexValue(Index), flag_ok)
          Exit Sub
       End If
    End If

       Call LostFocus_Handle(Me, txtRefractiveIndexValue(Index), flag_ok)
End Sub

