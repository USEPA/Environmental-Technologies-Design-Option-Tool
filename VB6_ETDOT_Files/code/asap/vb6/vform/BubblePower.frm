VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmBubblePower 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Power Calculations"
   ClientHeight    =   4845
   ClientLeft      =   6105
   ClientTop       =   3360
   ClientWidth     =   4710
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   4710
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
      Left            =   1770
      TabIndex        =   16
      Top             =   4080
      Width           =   1212
   End
   Begin Threed.SSFrame fraBlowerBrakePower 
      Height          =   1935
      Left            =   240
      TabIndex        =   4
      Top             =   210
      Width           =   4215
      _Version        =   65536
      _ExtentX        =   7429
      _ExtentY        =   3408
      _StockProps     =   14
      Caption         =   "Blower Brake Power:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtPower 
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
         Left            =   2490
         TabIndex        =   1
         Top             =   720
         Width           =   1572
      End
      Begin VB.TextBox txtPower 
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
         Index           =   0
         Left            =   2490
         TabIndex        =   0
         Top             =   360
         Width           =   1572
      End
      Begin VB.TextBox txtPower 
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
         Left            =   2490
         TabIndex        =   2
         Top             =   1080
         Width           =   1572
      End
      Begin VB.Label lblPower 
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
         Index           =   3
         Left            =   2490
         TabIndex        =   10
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label lblPowerLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Blower Brake Power"
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
         Index           =   3
         Left            =   90
         TabIndex        =   9
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label lblPowerLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Blower Efficiency"
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
         Left            =   90
         TabIndex        =   8
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label lblPowerLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Inlet Air Temperature"
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
         Left            =   90
         TabIndex        =   7
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label lblPowerLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Tank Water Depth"
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
         Left            =   90
         TabIndex        =   6
         Top             =   1080
         Width           =   2295
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   1455
      Left            =   240
      TabIndex        =   5
      Top             =   2310
      Width           =   4215
      _Version        =   65536
      _ExtentX        =   7435
      _ExtentY        =   2561
      _StockProps     =   14
      Caption         =   "Total Brake Power:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.23
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtPower 
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
         Index           =   5
         Left            =   2490
         TabIndex        =   3
         Top             =   600
         Width           =   1572
      End
      Begin VB.Label lblPower 
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
         Index           =   6
         Left            =   2490
         TabIndex        =   15
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lblPowerLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Brake Power"
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
         Index           =   6
         Left            =   90
         TabIndex        =   14
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label lblPowerLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "No. of Tanks (in Series)"
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
         Index           =   4
         Left            =   90
         TabIndex        =   13
         Top             =   300
         Width           =   2295
      End
      Begin VB.Label lblPower 
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
         Index           =   4
         Left            =   2490
         TabIndex        =   12
         Top             =   300
         Width           =   1575
      End
      Begin VB.Label lblPowerLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "No. of Blowers per Tank"
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
         Index           =   5
         Left            =   90
         TabIndex        =   11
         Top             =   600
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmBubblePower"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    frmBubblePower.Hide
End Sub

Private Sub Form_Activate()
  Call CenterThisForm(Me)
End Sub

Private Sub Form_Load()

    Call CenterThisForm(Me)
    Call LabelsBubblePowerSI

End Sub

Private Sub txtPower_GotFocus(Index As Integer)
  Call GotFocus_Handle(Me, txtPower(Index), Temp_Text)
End Sub

Private Sub txtPower_KeyPress(Index As Integer, KeyAscii As Integer)
  Call TextBoxNumber_KeyPress(KeyAscii)
End Sub

Private Sub txtPower_LostFocus(Index As Integer)
    Dim ValueChanged As Integer, Dummy As Double
    Dim msg As String
    Dim CalculatedPower As Integer
Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtPower(Index))) Then
     Exit Sub
   End If
   
   flag_ok = True

    Call TextHandleError(IsError, txtPower(Index), Temp_Text)

    If Not IsError Then
       If Index = 5 Then
          Dummy = CInt(txtPower(Index).Text)
       Else
          Dummy = CDbl(txtPower(Index).Text)
       End If
       Select Case Index
          Case 0    'InletAirTemperature
             Call TextNumberChanged(ValueChanged, txtPower(0), Temp_Text)
             If ValueChanged Then
                If HaveValue(Dummy) Then
                   bub.Power.InletAirTemperature = Dummy
                Else
                   txtPower(0).Text = Temp_Text
                   txtPower(0).SetFocus
                   Exit Sub
                End If
             End If
          Case 1    'Blower Efficiency
             Call TextNumberChanged(ValueChanged, txtPower(1), Temp_Text)
             If ValueChanged Then
                If HaveValue(Dummy) Then
                   bub.Power.BlowerEfficiency = Dummy
                Else
                   txtPower(1).Text = Temp_Text
                   txtPower(1).SetFocus
                   Exit Sub
                End If
             End If
          Case 2    'Tank Water Depth
             Call TextNumberChanged(ValueChanged, txtPower(2), Temp_Text)
             If ValueChanged Then
                If HaveValue(Dummy) Then
                   bub.Power.TankWaterDepth = Dummy
                Else
                   txtPower(2).Text = Temp_Text
                   txtPower(2).SetFocus
                   Exit Sub
                End If
             End If

          Case 5    'Number of Blowers per Tank
             Call TextNumberChanged(ValueChanged, txtPower(5), Temp_Text)
             If ValueChanged Then
                If HaveValue(Dummy) Then
                   bub.Power.NumberOfBlowersinEachTank = Dummy
                Else
                   txtPower(5).Text = Temp_Text
                   txtPower(5).SetFocus
                   Exit Sub
                End If
             End If

       End Select

       If ValueChanged Then
          Call CalculatePowerBubble
          
             frmBubblePower!txtPower(0).Text = Format$(bub.Power.InletAirTemperature, GetTheFormat(bub.Power.InletAirTemperature))
             frmBubblePower!txtPower(1).Text = Format$(bub.Power.BlowerEfficiency, GetTheFormat(bub.Power.BlowerEfficiency))
             frmBubblePower!txtPower(2).Text = Format$(bub.Power.TankWaterDepth, GetTheFormat(bub.Power.TankWaterDepth))
             frmBubblePower!lblPower(3).Caption = Format$(bub.Power.BlowerBrakePower, GetTheFormat(bub.Power.BlowerBrakePower))
             frmBubblePower!lblPower(4).Caption = Format$(bub.NumberOfTanks.Value, "0")
             frmBubblePower!txtPower(5).Text = Format$(bub.Power.NumberOfBlowersinEachTank, "0")
             frmBubblePower!lblPower(6).Caption = Format$(bub.Power.TotalBrakePower, GetTheFormat(bub.Power.TotalBrakePower))
          
       End If

    End If
  Call LostFocus_Handle(Me, txtPower(Index), flag_ok)


End Sub


