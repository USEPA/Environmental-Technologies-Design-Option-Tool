VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmPowerScreen2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Power Calculations"
   ClientHeight    =   5475
   ClientLeft      =   450
   ClientTop       =   2280
   ClientWidth     =   4710
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
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
      Left            =   1740
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4710
      Width           =   1212
   End
   Begin Threed.SSFrame fraPumpBrakePower 
      Height          =   1215
      Left            =   240
      TabIndex        =   4
      Top             =   2070
      Width           =   4215
      _Version        =   65536
      _ExtentX        =   7429
      _ExtentY        =   2138
      _StockProps     =   14
      Caption         =   "Pump Brake Power:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
         Index           =   3
         Left            =   2520
         TabIndex        =   2
         Top             =   330
         Width           =   1572
      End
      Begin VB.Label lblPowerLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Pump Efficiency"
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
         Left            =   120
         TabIndex        =   7
         Top             =   330
         Width           =   2295
      End
      Begin VB.Label lblPowerLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Pump Brake Power"
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
         Left            =   120
         TabIndex        =   6
         Top             =   690
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
         Left            =   2520
         TabIndex        =   5
         Top             =   690
         Width           =   1575
      End
   End
   Begin Threed.SSFrame fraTotalBrakePower 
      Height          =   855
      Left            =   240
      TabIndex        =   8
      Top             =   3540
      Width           =   4215
      _Version        =   65536
      _ExtentX        =   7429
      _ExtentY        =   1503
      _StockProps     =   14
      Caption         =   "Total Brake Power:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
         Index           =   5
         Left            =   120
         TabIndex        =   10
         Top             =   360
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
         Index           =   5
         Left            =   2520
         TabIndex        =   9
         Top             =   360
         Width           =   1575
      End
   End
   Begin Threed.SSFrame fraBlowerBrakePower 
      Height          =   1575
      Left            =   240
      TabIndex        =   11
      Top             =   240
      Width           =   4215
      _Version        =   65536
      _ExtentX        =   7429
      _ExtentY        =   2773
      _StockProps     =   14
      Caption         =   "Blower Brake Power:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
         Index           =   0
         Left            =   2520
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
         Index           =   1
         Left            =   2520
         TabIndex        =   1
         Top             =   720
         Width           =   1572
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
         Left            =   120
         TabIndex        =   15
         Top             =   360
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
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   2295
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
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   1080
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
         Index           =   2
         Left            =   2520
         TabIndex        =   12
         Top             =   1080
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmPowerScreen2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
  frmPowerScreen2.Hide
End Sub

Private Sub Form_Activate()
    frmPowerScreen2.WindowState = 0

    'Position the form on the screen (Centered)
    If WindowState = 0 Then
       'don't attempt if screen Minimized or Maximized
       Move (Screen.Width - frmPowerScreen2.Width) / 2, (Screen.Height - frmPowerScreen2.Height) / 2
    
       If frmPTADScreen1.Visible = True Then Move frmPTADScreen1.Left + (frmPTADScreen1.Width / 2) - (frmPowerScreen2.Width / 2), frmPTADScreen1.Top + (frmPTADScreen1.Height / 2) - (frmPowerScreen2.Height / 2)
       If frmPTADScreen2.Visible = True Then Move frmPTADScreen2.Left + (frmPTADScreen2.Width / 2) - (frmPowerScreen2.Width / 2), frmPTADScreen2.Top + (frmPTADScreen2.Height / 2) - (frmPowerScreen2.Height / 2)
       If frmBubble.Visible = True Then Move frmBubble.Left + (frmBubble.Width / 2) - (frmPowerScreen2.Width / 2), frmBubble.Top + (frmBubble.Height / 2) - (frmPowerScreen2.Height / 2)
       If frmSurface.Visible = True Then Move frmSurface.Left + (frmSurface.Width / 2) - (frmPowerScreen2.Width / 2), frmSurface.Top + (frmSurface.Height / 2) - (frmPowerScreen2.Height / 2)
       
    End If


End Sub

Private Sub Form_Load()

    Call LabelsPowerScreen2SI

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
       Dummy = CDbl(txtPower(Index).Text)
       Select Case Index
          Case 0    'InletAirTemperature
             Call TextNumberChanged(ValueChanged, txtPower(0), Temp_Text)
             If ValueChanged Then
                If HaveValue(Dummy) Then
                   Scr2.Power.InletAirTemperature = Dummy
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
                   Scr2.Power.BlowerEfficiency = Dummy
                Else
                   txtPower(1).Text = Temp_Text
                   txtPower(1).SetFocus
                   Exit Sub
                End If
             End If
          Case 3    'Pump Efficiency
             Call TextNumberChanged(ValueChanged, txtPower(3), Temp_Text)
             If ValueChanged Then
                If HaveValue(Dummy) Then
                   Scr2.Power.PumpEfficiency = Dummy
                Else
                   txtPower(3).Text = Temp_Text
                   txtPower(3).SetFocus
                   Exit Sub
                End If
             End If

       End Select

       If ValueChanged Then
          Call CalculatePowerScreen2(CalculatedPower)
          If CalculatedPower Then
             frmPowerScreen2!txtPower(0).Text = Format$(Scr2.Power.InletAirTemperature, "0.0")
             frmPowerScreen2!txtPower(1).Text = Format$(Scr2.Power.BlowerEfficiency, "0.0")
             frmPowerScreen2!lblPower(2).Caption = Format$(Scr2.Power.BlowerBrakePower, "0.000")
             frmPowerScreen2!txtPower(3).Text = Format$(Scr2.Power.PumpEfficiency, "0.0")
             frmPowerScreen2!lblPower(4).Caption = Format$(Scr2.Power.PumpBrakePower, "0.000")
             frmPowerScreen2!lblPower(5).Caption = Format$(Scr2.Power.TotalBrakePower, "0.000")
          End If
       End If

    End If
  Call LostFocus_Handle(Me, txtPower(Index), flag_ok)


End Sub


