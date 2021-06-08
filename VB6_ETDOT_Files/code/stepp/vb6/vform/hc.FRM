VERSION 5.00
Begin VB.Form hc_form 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Henry's Constant"
   ClientHeight    =   6210
   ClientLeft      =   2940
   ClientTop       =   975
   ClientWidth     =   8745
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6210
   ScaleWidth      =   8745
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   2
      Left            =   240
      TabIndex        =   8
      Top             =   2640
      Width           =   225
   End
   Begin VB.TextBox txtHCmaximumT 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   5
      Left            =   7560
      TabIndex        =   41
      Text            =   "No Visible"
      Top             =   5160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtHCminimumT 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   5
      Left            =   6360
      TabIndex        =   40
      Text            =   "No Visible"
      Top             =   5160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtHCTemperature 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   5
      Left            =   5160
      TabIndex        =   39
      Top             =   5160
      Width           =   975
   End
   Begin VB.TextBox txtHenrysConstantValue 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   5
      Left            =   3120
      TabIndex        =   38
      Top             =   5160
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   6
      Left            =   240
      TabIndex        =   36
      Top             =   5160
      Width           =   225
   End
   Begin VB.ComboBox cboUNIFAC 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   3120
      Style           =   2  'Dropdown List
      TabIndex        =   35
      Top             =   4680
      Width           =   3492
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   5
      Left            =   240
      TabIndex        =   28
      Top             =   4320
      Width           =   225
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      Caption         =   "Accept Selected Henry's Constant"
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
      Index           =   4
      Left            =   240
      TabIndex        =   20
      Top             =   3480
      Width           =   225
   End
   Begin VB.ComboBox hc_list 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   3120
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   3840
      Width           =   3492
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   3
      Left            =   240
      TabIndex        =   14
      Top             =   3000
      Width           =   225
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   2280
      Width           =   225
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   3852
      Left            =   120
      Top             =   1680
      Width           =   8532
   End
   Begin VB.Label lblUNIFACParameterSet 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   3720
      TabIndex        =   53
      Top             =   5760
      Width           =   4812
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      Height          =   492
      Left            =   120
      Top             =   5640
      Width           =   8532
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "UNIFAC Group Interaction Database:"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   240
      TabIndex        =   52
      Top             =   5760
      Width           =   3252
   End
   Begin VB.Label lblHCmaximumT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "not visible"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   4
      Left            =   7560
      TabIndex        =   33
      Top             =   4320
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label lblHCmaximumT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "not visible"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   3
      Left            =   7560
      TabIndex        =   25
      Top             =   3480
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label lblHCmaximumT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "not visible"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   2
      Left            =   7560
      TabIndex        =   19
      Top             =   3000
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label lblHCmaximumT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "not visible"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   7560
      TabIndex        =   13
      Top             =   2640
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label lblHCmaximumT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   7560
      TabIndex        =   7
      Top             =   2280
      Width           =   972
   End
   Begin VB.Label lblHCminimumT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "not visible"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   4
      Left            =   6360
      TabIndex        =   32
      Top             =   4320
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label lblHCminimumT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "not visible"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   3
      Left            =   6360
      TabIndex        =   24
      Top             =   3480
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label lblHCminimumT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "not visible"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   2
      Left            =   6360
      TabIndex        =   18
      Top             =   3000
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label lblHCminimumT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "not visible"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   6360
      TabIndex        =   12
      Top             =   2640
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label lblHCminimumT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   6360
      TabIndex        =   6
      Top             =   2280
      Width           =   972
   End
   Begin VB.Label lblHCTemperature 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   4
      Left            =   5160
      TabIndex        =   31
      Top             =   4320
      Width           =   972
   End
   Begin VB.Label lblHCTemperature 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   3
      Left            =   5160
      TabIndex        =   23
      Top             =   3480
      Width           =   972
   End
   Begin VB.Label lblHCTemperature 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   2
      Left            =   5160
      TabIndex        =   17
      Top             =   3000
      Width           =   972
   End
   Begin VB.Label lblHCTemperature 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   5160
      TabIndex        =   11
      Top             =   2640
      Width           =   972
   End
   Begin VB.Label lblHCTemperature 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   5160
      TabIndex        =   5
      Top             =   2280
      Width           =   972
   End
   Begin VB.Label lblHenrysConstantValue 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   4
      Left            =   3120
      TabIndex        =   30
      Top             =   4320
      Width           =   1812
   End
   Begin VB.Label lblHenrysConstantValue 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   3
      Left            =   3120
      TabIndex        =   22
      Top             =   3480
      Width           =   1812
   End
   Begin VB.Label lblHenrysConstantValue 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   2
      Left            =   3120
      TabIndex        =   16
      Top             =   3000
      Width           =   1812
   End
   Begin VB.Label lblHenrysConstantValue 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   3120
      TabIndex        =   10
      Top             =   2640
      Width           =   1812
   End
   Begin VB.Label lblHenrysConstantValue 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   3120
      TabIndex        =   4
      Top             =   2280
      Width           =   1812
   End
   Begin VB.Label lblUNIFAC 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   840
      TabIndex        =   34
      Top             =   4680
      Width           =   1812
   End
   Begin VB.Label lblDatabase 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   840
      TabIndex        =   26
      Top             =   3840
      Width           =   1812
   End
   Begin VB.Label lblSource 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "User Input"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   600
      TabIndex        =   37
      Top             =   5160
      Width           =   2295
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   8640
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label lblSource 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "UNIFAC @ Database T's"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   600
      TabIndex        =   29
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Label lblSource 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Database"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   600
      TabIndex        =   21
      Top             =   3480
      Width           =   2295
   End
   Begin VB.Label lblSource 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "UNIFAC at Operating T"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   600
      TabIndex        =   15
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label lblSource 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "UNIFAC Fit with Data Pt."
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   9
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label lblSource 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Regression of Data Points"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   3
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label lblHCLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Henry's Constant"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3120
      TabIndex        =   51
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label lblHCminTLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tmin"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6360
      TabIndex        =   50
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label lblHCmaxTLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tmax"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7560
      TabIndex        =   49
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Source"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   600
      TabIndex        =   48
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label lblHCTempLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Temp."
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   5160
      TabIndex        =   47
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Current Henry's Constant Information"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   46
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
      Left            =   360
      TabIndex        =   45
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
      Left            =   360
      TabIndex        =   44
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
      Left            =   2160
      TabIndex        =   43
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
      Left            =   2160
      TabIndex        =   42
      Top             =   720
      Width           =   2295
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   2
      Height          =   1335
      Left            =   240
      Top             =   120
      Width           =   4335
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   4560
      Y1              =   600
      Y2              =   600
   End
End
Attribute VB_Name = "hc_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PastHCInputValue As String
Dim PastHCInputTemp As String

Private Sub cboUNIFAC_Click()
    Dim i As Integer, J As Integer
    Dim msg As String
    Dim SIValue As Double, EnglishValue As Double
    Dim ValueToDisplay As Double

    i = phprop.HenrysConstant.chosenUNIFACIndex
    J = cboUNIFAC.ListIndex + 1

    If i = J Then Exit Sub

    If Left$(cboUNIFAC.List(J - 1), 3) = "N/A" Then
       msg = "This UNIFAC item can not be selected"
       msg = msg + " because it is not available"
       MsgBox msg, MB_ICONSTOP, "Warning"
       cboUNIFAC.ListIndex = i - 1
    Else
       phprop.HenrysConstant.chosenUNIFACIndex = J

       SIValue = phprop.HenrysConstant.UNIFAC(phprop.HenrysConstant.chosenUNIFACIndex).Value
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call HCCONV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       hc_form!lblHenrysConstantValue(4).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       SIValue = phprop.HenrysConstant.UNIFAC(phprop.HenrysConstant.chosenUNIFACIndex).temperature
       If CurrentUnits = SIUnits Then
          ValueToDisplay = SIValue
       ElseIf CurrentUnits = EnglishUnits Then
          Call TEMPCNV(EnglishValue, SIValue)
          ValueToDisplay = EnglishValue
       End If
       hc_form!lblHCTemperature(4).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

    End If

End Sub

Private Sub cboUNIFAC_KeyPress(KeyAscii As Integer)
    KeyPress = 0
End Sub

Private Sub cmdCancel_Click()
    Dim SelectedOption As Integer   'Option selected permanently by the user (the option corresponding to the number on the main screen)

    Select Case phprop.HenrysConstant.CurrentSelection.choice
       Case HENRYS_CONSTANT_REGRESS
          SelectedOption = 1
       Case HENRYS_CONSTANT_FIT
          SelectedOption = 2
       Case HENRYS_CONSTANT_OPT_UNIFAC
          SelectedOption = 3
       Case HENRYS_CONSTANT_DATABASE
          SelectedOption = 4
       Case HENRYS_CONSTANT_UNIFAC
          SelectedOption = 5
       Case HENRYS_CONSTANT_INPUT
          SelectedOption = 6
          txtHenrysConstantValue(5).Text = PastHCInputValue
          txtHCTemperature(5).Text = PastHCInputTemp
       Case Else
          hc_form.Hide
          Exit Sub
    End Select

    If Not Option1(SelectedOption).Value Then Option1(SelectedOption).Value = True

    hc_form.Hide
End Sub

Private Sub cmdOK_Click()
    Dim OptionSelected As Integer
    Dim ValueToDisplayIndex As Integer
    Dim i As Integer

'*** Pass new selected value back to main screen
    For i = 1 To 6
        If Option1(i).Value Then
           OptionSelected = i
           Exit For
        End If
    Next i

    Select Case OptionSelected
       Case 1
          ValueToDisplayIndex = HENRYS_CONSTANT_REGRESS
       Case 2
          ValueToDisplayIndex = HENRYS_CONSTANT_FIT
       Case 3
          ValueToDisplayIndex = HENRYS_CONSTANT_OPT_UNIFAC
       Case 4
          ValueToDisplayIndex = HENRYS_CONSTANT_DATABASE
       Case 5
          ValueToDisplayIndex = HENRYS_CONSTANT_UNIFAC
       Case 6
          If Not PROPAVAILABLE(HENRYS_CONSTANT_INPUT) Then
             MsgBox "User Input can not be selected without first entering a value", MB_ICONSTOP, "Error"
             txtHenrysConstantValue(5).SetFocus
             Exit Sub
          End If
          ValueToDisplayIndex = HENRYS_CONSTANT_INPUT
    End Select

    If ValueToDisplayIndex <> phprop.HenrysConstant.CurrentSelection.choice Then
       phprop.HenrysConstant.CurrentSelection.choice = ValueToDisplayIndex
       Call DisplayHenrysConstantMainScreen(ValueToDisplayIndex)
    ElseIf ValueToDisplayIndex = HENRYS_CONSTANT_INPUT Then
       Call DisplayHenrysConstantMainScreen(ValueToDisplayIndex)
    End If

    hc_form.Hide
End Sub

Private Sub Form_Activate()
    Dim ActCoeffIndex As Integer

  Call centerform_relative(contam_prop_form, Me)
    
    PastHCInputValue = txtHenrysConstantValue(5).Text
    PastHCInputTemp = txtHCTemperature(5).Text

    If HaveProperty(ACTIVITY_COEFFICIENT) And HaveProperty(VAPOR_PRESSURE) Then
       Select Case phprop.ActivityCoefficient.BinaryInteractionParameterDatabase
          Case 1   'Original UNIFAC VLE
             hc_form!lblUNIFACParameterSet.Caption = "Original UNIFAC VLE"
          Case 2   'UNIFAC LLE
             hc_form!lblUNIFACParameterSet.Caption = "UNIFAC LLE"
          Case 3   'Environmental VLE'
             hc_form!lblUNIFACParameterSet.Caption = "Environmental VLE"
       End Select
    Else
       hc_form!lblUNIFACParameterSet.Caption = "UNIFAC Henry's Constant Not Available"
    End If

End Sub

Private Sub Form_Load()

  Call centerform_relative(contam_prop_form, Me)

    If (DemoMode) Then cmdok.Enabled = False


End Sub

Private Sub hc_list_Click()
    Dim i As Integer, J As Integer
    Dim SIValue As Double, EnglishValue As Double
    Dim ValueToDisplay As Double

    i = phprop.HenrysConstant.chosenDatabaseIndex
    J = hc_list.ListIndex + 1

    If i = J Then Exit Sub

    phprop.HenrysConstant.chosenDatabaseIndex = J

    SIValue = phprop.HenrysConstant.database(phprop.HenrysConstant.chosenDatabaseIndex).Value
    If CurrentUnits = SIUnits Then
       ValueToDisplay = SIValue
    ElseIf CurrentUnits = EnglishUnits Then
       Call HCCONV(EnglishValue, SIValue)
       ValueToDisplay = EnglishValue
    End If
    hc_form!lblHenrysConstantValue(3).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

    SIValue = phprop.HenrysConstant.database(phprop.HenrysConstant.chosenDatabaseIndex).temperature
    If CurrentUnits = SIUnits Then
       ValueToDisplay = SIValue
    ElseIf CurrentUnits = EnglishUnits Then
       Call TEMPCNV(EnglishValue, SIValue)
       ValueToDisplay = EnglishValue
    End If
    hc_form!lblHCTemperature(3).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

End Sub

Private Sub hc_list_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub lblHCmaximumT_Click(Index As Integer)
0    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.HenrysConstant.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

End Sub

Private Sub lblHCminimumT_Click(Index As Integer)
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.HenrysConstant.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

End Sub

Private Sub lblHCTemperature_Click(Index As Integer)
    
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.HenrysConstant.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

End Sub

Private Sub lblHenrysConstantValue_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ErrorFlag As Integer
    Dim msg As String
    
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.HenrysConstant.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

    If Button <> 2 Then Exit Sub

    Select Case Index
       Case 0   'Henry's constant regression
            ErrorFlag = phprop.HenrysConstant.regress.error
       Case 1   'UNIFAC Fit with Data Point
            ErrorFlag = phprop.HenrysConstant.fit.UNIFAC.error
       Case 2   'UNIFAC at Operating T
            ErrorFlag = phprop.HenrysConstant.operatingT.UNIFAC.error
       Case 3   'Database
            If hc_form!lblHenrysConstantValue(3).Caption = "Not Available" Then
               If phprop.HenrysConstant.NumberOfDatabaseHenrysConstants = 0 Then
                  msg = "There are no Henry's constants in the database."
                  MsgBox msg, MB_ICONINFORMATION, Trim$(phprop.Name) & " - Data Unavailable"
                  Exit Sub
               Else
                  ErrorFlag = phprop.HenrysConstant.database(1).error
               End If
            Else
               ErrorFlag = phprop.HenrysConstant.database(phprop.HenrysConstant.chosenDatabaseIndex).error
            End If
       Case 4   'UNIFAC at Database Ts
            If hc_form!lblHenrysConstantValue(4).Caption = "Not Available" Then
               If phprop.HenrysConstant.NumberOfDatabaseHenrysConstants = 0 Then
                  msg = "There are no Henry's constants in the database.  Consequently, the calculation of Henry's Constants from UNIFAC at database temperatures is not possible."
                  MsgBox msg, MB_ICONINFORMATION, Trim$(phprop.Name) & " - Data Unavailable"
                  Exit Sub
               Else
                  ErrorFlag = phprop.HenrysConstant.UNIFAC(1).error
               End If
            Else
               ErrorFlag = phprop.HenrysConstant.UNIFAC(phprop.HenrysConstant.chosenUNIFACIndex).error
            End If

    End Select


    If ErrorFlag = 0 Then Exit Sub   'No errors or warnings

    If ErrorFlag < 0 Then
       MsgBox ErrorMsg(ErrorFlag), MB_ICONINFORMATION, Trim$(phprop.Name) & " - Data Unavailable"
    End If

    If ErrorFlag > 0 Then
       MsgBox ErrorMsg(ErrorFlag), MB_ICONINFORMATION, Trim$(phprop.Name) & " - Warning"
    End If

End Sub

Private Sub lblSource_Click(Index As Integer)
    Dim i As Integer

    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.HenrysConstant.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

End Sub

Private Sub Option1_Click(Index As Integer)
    Dim i As Integer, SourceIndex As Integer

    SourceIndex = Index - 1
    If SourceIndex = hilight.HenrysConstant.PreviousIndex Then Exit Sub
    lblSource(SourceIndex).BackColor = &H800000
    lblSource(SourceIndex).ForeColor = &H80000005
    i = hilight.HenrysConstant.PreviousIndex
    hilight.HenrysConstant.PreviousIndex = SourceIndex
    If i = -1 Then Exit Sub
    If Option1(i + 1).Enabled = False Then Exit Sub
    lblSource(i).BackColor = &H80000005
    lblSource(i).ForeColor = &H80000008

End Sub

Private Sub txtHCmaximumT_GotFocus(Index As Integer)
  Call GotFocus_Handle(Me, txtHCmaximumT(Index), Temp_Text)

End Sub

Private Sub txtHCmaximumT_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index <> 5 Then KeyAscii = 0
End Sub

Private Sub txtHCmaximumT_LostFocus(Index As Integer)
Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtHCmaximumT(Index))) Then
     Exit Sub
   End If

   flag_ok = True
  Call LostFocus_Handle(Me, txtHCmaximumT(Index), flag_ok)

End Sub

Private Sub txtHCminimumT_GotFocus(Index As Integer)
  Call GotFocus_Handle(Me, txtHCminimumT(Index), Temp_Text)

End Sub

Private Sub txtHCminimumT_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index <> 5 Then KeyAscii = 0
End Sub

Private Sub txtHCminimumT_LostFocus(Index As Integer)
Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtHCminimumT(Index))) Then
     Exit Sub
   End If

   flag_ok = True
  Call LostFocus_Handle(Me, txtHCminimumT(Index), flag_ok)

End Sub

Private Sub txtHCTemperature_GotFocus(Index As Integer)
  Call GotFocus_Handle(Me, txtHCTemperature(Index), Temp_Text)

    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.HenrysConstant.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If

End Sub

Private Sub txtHCTemperature_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = 13 Then
       KeyAscii = 0
       cmdok.SetFocus
       Exit Sub
    End If
    Call NumberCheck(KeyAscii)
   
End Sub

Private Sub txtHCTemperature_LostFocus(Index As Integer)
    Dim msg As String, Response As Integer
    Dim Answer As Integer
    Dim IsError As Integer
    Dim ValueChanged As Integer
    Dim EnglishValue As Double, SIValue As Double
    Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtHCTemperature(Index))) Then
     Exit Sub
   End If

   flag_ok = True

    If txtHCTemperature(Index).Text = "" Then
       phprop.HenrysConstant.input.temperature = -1E+25
       Call LostFocus_Handle(Me, txtHCTemperature(Index), flag_ok)
       Exit Sub
    End If

    Call TextHandleError(IsError, txtHCTemperature(Index), Temp_Text)
    If Not IsError Then
       If Not HaveNumber(CDbl(txtHCTemperature(Index).Text)) Then
          txtHCTemperature(Index).Text = Temp_Text
          txtHCTemperature(Index).SetFocus
          Call LostFocus_Handle(Me, txtHCTemperature(Index), flag_ok)
          Exit Sub
       End If

       Call TextNumberChanged(ValueChanged, txtHCTemperature(Index), Temp_Text)

       If ValueChanged Then
          If CurrentUnits = SIUnits Then
             phprop.HenrysConstant.input.temperature = CDbl(txtHCTemperature(Index).Text)
          Else
             EnglishValue = CDbl(txtHCTemperature(Index).Text)
             Call TEMPENSI(SIValue, EnglishValue)
             phprop.HenrysConstant.input.temperature = SIValue
          End If

       Else
          Call LostFocus_Handle(Me, txtHCTemperature(Index), flag_ok)
          Exit Sub
       End If

    End If
          Call LostFocus_Handle(Me, txtHCTemperature(Index), flag_ok)
End Sub

Private Sub txtHenrysConstantValue_GotFocus(Index As Integer)
  Call GotFocus_Handle(Me, txtHenrysConstantValue(Index), Temp_Text)
    
    If Option1(Index + 1).Enabled = True Then
       If Index = hilight.HenrysConstant.PreviousIndex Then Exit Sub
       Option1(Index + 1).Value = True
    End If


End Sub

Private Sub txtHenrysConstantValue_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = 13 Then
       KeyAscii = 0
       txtHCTemperature(Index).SetFocus
       Exit Sub
    End If
    Call NumberCheck(KeyAscii)
    
End Sub

Private Sub txtHenrysConstantValue_LostFocus(Index As Integer)
    Dim msg As String, Response As Integer
    Dim Answer As Integer
    Dim IsError As Integer
    Dim ValueChanged As Integer
    Dim EnglishValue As Double, SIValue As Double
    Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtHenrysConstantValue(Index))) Then
     Exit Sub
   End If

   flag_ok = True

    If txtHenrysConstantValue(Index).Text = "" Then
       phprop.HenrysConstant.input.Value = -1#
       PROPAVAILABLE(HENRYS_CONSTANT_INPUT) = False
       Call LostFocus_Handle(Me, txtHenrysConstantValue(Index), flag_ok)
       Exit Sub
    End If

    Call TextHandleError(IsError, txtHenrysConstantValue(Index), Temp_Text)
    If Not IsError Then
       If Not HaveNumber(CDbl(txtHenrysConstantValue(Index).Text)) Then
          txtHenrysConstantValue(Index).Text = Temp_Text
          txtHenrysConstantValue(Index).SetFocus
       Call LostFocus_Handle(Me, txtHenrysConstantValue(Index), flag_ok)
          Exit Sub
       End If

       Call TextNumberChanged(ValueChanged, txtHenrysConstantValue(Index), Temp_Text)

       If ValueChanged Then
          If CurrentUnits = SIUnits Then
             phprop.HenrysConstant.input.Value = CDbl(txtHenrysConstantValue(Index).Text)
          Else
             EnglishValue = CDbl(txtHenrysConstantValue(Index).Text)
             Call HCENSI(SIValue, EnglishValue)
             phprop.HenrysConstant.input.Value = SIValue
          End If

          PROPAVAILABLE(HENRYS_CONSTANT_INPUT) = True
       Else
       Call LostFocus_Handle(Me, txtHenrysConstantValue(Index), flag_ok)
          Exit Sub
       End If
    End If
       Call LostFocus_Handle(Me, txtHenrysConstantValue(Index), flag_ok)
End Sub

