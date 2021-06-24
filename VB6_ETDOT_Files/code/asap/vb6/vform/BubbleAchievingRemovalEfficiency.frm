VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmBubbleAchievingRemovalEfficiency 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bubble Aeration - Achieving Desired Removal Efficiency"
   ClientHeight    =   6075
   ClientLeft      =   2865
   ClientTop       =   3165
   ClientWidth     =   8970
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   8970
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4080
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5280
      Width           =   852
   End
   Begin Threed.SSFrame Frame3D1 
      Height          =   1305
      Index           =   0
      Left            =   1980
      TabIndex        =   4
      Top             =   3630
      Width           =   4515
      _Version        =   65536
      _ExtentX        =   7964
      _ExtentY        =   2302
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtAchieving 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Index           =   1
         Left            =   3060
         TabIndex        =   0
         Top             =   540
         Width           =   1272
      End
      Begin VB.TextBox txtAchieving 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Index           =   2
         Left            =   3060
         TabIndex        =   1
         Top             =   900
         Width           =   1272
      End
      Begin VB.Label lblAchievingLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Minimum Air To Water Ratio"
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
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   8
         Top             =   180
         Width           =   2715
      End
      Begin VB.Label lblAchievingLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Air To Water Ratio"
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
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   7
         Top             =   540
         Width           =   2715
      End
      Begin VB.Label lblAchievingLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Tanks"
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
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   6
         Top             =   900
         Width           =   2715
      End
      Begin VB.Label lblAchieving 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   255
         Index           =   0
         Left            =   3060
         TabIndex        =   5
         Top             =   180
         Width           =   1275
      End
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   3015
      Left            =   300
      TabIndex        =   3
      Top             =   240
      Width           =   8355
   End
End
Attribute VB_Name = "frmBubbleAchievingRemovalEfficiency"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MinimumAirToWaterRatio As Double
Dim AirToWaterRatio As Double
Dim NumberOfTanks As Long

Private Sub cmdOK_Click()

    If Abs(bub.MinimumAirToWaterRatio.Value - MinimumAirToWaterRatio) > TOLERANCE Then
       bub.MinimumAirToWaterRatio.Value = MinimumAirToWaterRatio
       frmBubble!lblFlowParameters(1).Caption = Format$(bub.MinimumAirToWaterRatio.Value, GetTheFormat(bub.MinimumAirToWaterRatio.Value))
    End If

    If Abs(bub.AirToWaterRatio.Value - AirToWaterRatio) > TOLERANCE Then
       bub.AirToWaterRatio.Value = AirToWaterRatio
       bub.AirToWaterRatio.UserInput = True
       frmBubble!txtFlowParameters(2).Text = Format$(bub.AirToWaterRatio.Value, GetTheFormat(bub.AirToWaterRatio.Value))
       Call CalculateAirFlowRate
    End If

    If Abs(bub.NumberOfTanks.Value - NumberOfTanks) > TOLERANCE Then
       bub.NumberOfTanks.Value = NumberOfTanks
       frmBubble!txtTankParameters(0).Text = Format$(bub.NumberOfTanks.Value, "0")
    End If

    frmBubbleAchievingRemovalEfficiency.Hide
End Sub

Private Sub Form_Activate()
    Dim msg As String
    
    'Center the form on the screen
    If WindowState = 0 Then
       'don't attempt if screen Minimized or Maximized
      If frmPTADScreen1.Visible = True Then Move frmPTADScreen1.Left + (frmPTADScreen1.Width / 2) - (frmBubbleAchievingRemovalEfficiency.Width / 2), frmPTADScreen1.Top + (frmPTADScreen1.Height / 2) - (frmBubbleAchievingRemovalEfficiency.Height / 2)
      If frmPTADScreen2.Visible = True Then Move frmPTADScreen2.Left + (frmPTADScreen2.Width / 2) - (frmBubbleAchievingRemovalEfficiency.Width / 2), frmPTADScreen2.Top + (frmPTADScreen2.Height / 2) - (frmBubbleAchievingRemovalEfficiency.Height / 2)
      If frmBubble.Visible = True Then Move frmBubble.Left + (frmBubble.Width / 2) - (frmBubbleAchievingRemovalEfficiency.Width / 2), frmBubble.Top + (frmBubble.Height / 2) - (frmBubbleAchievingRemovalEfficiency.Height / 2)
      If frmSurface.Visible = True Then Move frmSurface.Left + (frmSurface.Width / 2) - (frmBubbleAchievingRemovalEfficiency.Width / 2), frmSurface.Top + (frmSurface.Height / 2) - (frmBubbleAchievingRemovalEfficiency.Height / 2)
    End If

    msg = "It is not possible to achieve the desired removal efficiency with an air to water ratio "
    msg = msg & "that is less than the minimum air to water ratio.  Currently, the air to water "
    msg = msg & "ratio is less than the minimum air to water ratio, as shown below."
    msg = msg & "." & Chr$(13) & Chr$(13)
    msg = msg & "There are three options to allow the system to achieve the desired removal efficiency:"
    msg = msg & Chr$(13)
    msg = msg & "    (1)  Increase the air to water ratio to a number > the minimum air to water ratio" & Chr$(13)
    msg = msg & "    (2)  Increase the number of tanks, thereby reducing the minimum air to water ratio" & Chr$(13)
    msg = msg & "    (3)  Do a combination of 1 and 2" & Chr$(13) & Chr$(13)
    msg = msg & "In design mode, the goal is to achieve the desired removal efficiency.  Therefore, it "
    msg = msg & "is neccessary to arrive at a situation where air to water ratio is greater than "
    msg = msg & "minimum air to water ratio before continuing.  Do this by performing operations "
    msg = msg & "(1), (2), or (3) above.  When the air to water ratio is greater than the minimum "
    msg = msg & "air to water ratio, it will be possible to return to the main Bubble Aeration - "
    msg = msg & "Design Mode screen (i.e. the OK button will become enabled)."
    
    Label1.Caption = msg

    MinimumAirToWaterRatio = bub.MinimumAirToWaterRatio.Value
    AirToWaterRatio = bub.AirToWaterRatio.Value
    NumberOfTanks = bub.NumberOfTanks.Value

    cmdOK.Enabled = False

End Sub
Private Sub Form_Load()
  'REPLACE WHITE BACKGROUND WITH TRANSPARENT BACKGROUND.
  Label1.BackStyle = 0
End Sub


Private Sub txtAchieving_GotFocus(Index As Integer)
  Call GotFocus_Handle(Me, txtAchieving(Index), Temp_Text)
End Sub

Private Sub txtAchieving_KeyPress(Index As Integer, KeyAscii As Integer)
  Call TextBoxNumber_KeyPress(KeyAscii)
End Sub

Private Sub txtAchieving_LostFocus(Index As Integer)
    Dim ValueChanged As Integer, Dummy As Double
    Dim msg As String
Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtAchieving(Index))) Then
     Exit Sub
   End If
   
   flag_ok = True

    Call TextHandleError(IsError, txtAchieving(Index), Temp_Text)

    If Not IsError Then
       If Index = 2 Then
          Dummy = CInt(txtAchieving(Index).Text)
       Else
          Dummy = CDbl(txtAchieving(Index).Text)
       End If

       Select Case Index
          Case 1    'Air To Water Ratio
             Call TextNumberChanged(ValueChanged, txtAchieving(1), Temp_Text)
             If ValueChanged Then
                If HaveValue(Dummy) Then
                   AirToWaterRatio = Dummy
                Else
                   txtAchieving(1).Text = Temp_Text
                    txtAchieving(1).SetFocus
                    Call LostFocus_Handle(Me, txtAchieving(Index), flag_ok)
                   Exit Sub
                End If
             End If

          Case 2    'Number of Tanks in Series
             Call TextNumberChanged(ValueChanged, txtAchieving(2), Temp_Text)
             If ValueChanged Then
                If HaveValue(Dummy) Then
                   NumberOfTanks = Dummy
                   Call VQMINBUB(MinimumAirToWaterRatio, bub.DesignContaminant.Influent.Value, bub.DesignContaminant.TreatmentObjective.Value, bub.DesignContaminant.HenrysConstant.Value, NumberOfTanks)
                   lblAchieving(0).Caption = Format$(MinimumAirToWaterRatio, GetTheFormat(MinimumAirToWaterRatio))
                Else
                   txtAchieving(2).Text = Temp_Text
                   txtAchieving(2).SetFocus
                    Call LostFocus_Handle(Me, txtAchieving(Index), flag_ok)
                   Exit Sub
                End If
             End If
       End Select
       If AirToWaterRatio > MinimumAirToWaterRatio Then
          cmdOK.Enabled = True
          cmdOK.SetFocus
       Else
          cmdOK.Enabled = False
       End If
    End If
  Call LostFocus_Handle(Me, txtAchieving(Index), flag_ok)


End Sub


