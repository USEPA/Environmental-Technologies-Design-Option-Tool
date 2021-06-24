VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmOxygenMassTransferCoeff 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find KLa, O2 from Clean Water Oxygen Transfer Test Data"
   ClientHeight    =   6120
   ClientLeft      =   1320
   ClientTop       =   1515
   ClientWidth     =   8925
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   8925
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
      Height          =   555
      Left            =   6510
      TabIndex        =   8
      Top             =   1560
      Width           =   1275
   End
   Begin Threed.SSFrame fraDataAvailable 
      Height          =   2265
      Left            =   120
      TabIndex        =   6
      Top             =   270
      Width           =   4305
      _Version        =   65536
      _ExtentX        =   7594
      _ExtentY        =   3995
      _StockProps     =   14
      Caption         =   "Data Available:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSOption optDataAvailable 
         Height          =   315
         Index           =   2
         Left            =   480
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1620
         Width           =   3075
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "No Data Available"
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   19.64
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption optDataAvailable 
         Height          =   315
         Index           =   1
         Left            =   480
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1020
         Width           =   3075
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "SOTE vs. Qair Data Available"
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   19.64
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption optDataAvailable 
         Height          =   315
         Index           =   0
         Left            =   480
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   360
         Width           =   3075
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "SOTR vs. Qair Data Available"
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   19.64
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin Threed.SSFrame fraDataParameters 
      Height          =   3240
      Left            =   120
      TabIndex        =   7
      Top             =   2700
      Width           =   8565
      _Version        =   65536
      _ExtentX        =   15108
      _ExtentY        =   5715
      _StockProps     =   14
      Caption         =   "Clean Water Oxygen Transfer Test Data Parameters:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtDataParameters 
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
         Height          =   285
         Index           =   0
         Left            =   2400
         TabIndex        =   0
         Top             =   540
         Width           =   1455
      End
      Begin VB.TextBox txtDataParameters 
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
         Height          =   285
         Index           =   1
         Left            =   2400
         TabIndex        =   1
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtDataParameters 
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
         Height          =   285
         Index           =   2
         Left            =   2400
         TabIndex        =   2
         Top             =   1380
         Width           =   1455
      End
      Begin VB.TextBox txtDataParameters 
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
         Height          =   285
         Index           =   3
         Left            =   2400
         TabIndex        =   3
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox txtDataParameters 
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
         Height          =   285
         Index           =   4
         Left            =   2400
         TabIndex        =   4
         Top             =   2220
         Width           =   1455
      End
      Begin VB.TextBox txtDataParameters 
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
         Height          =   285
         Index           =   5
         Left            =   2400
         TabIndex        =   5
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label lblDataParametersLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SOTE"
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
         Left            =   300
         TabIndex        =   29
         Top             =   540
         Width           =   1815
      End
      Begin VB.Label lblDataParametersLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SOTR"
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
         Left            =   300
         TabIndex        =   28
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label lblDataParametersLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Air Flow Rate"
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
         Left            =   300
         TabIndex        =   27
         Top             =   1380
         Width           =   1815
      End
      Begin VB.Label lblDataParametersLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Barometric Pressure"
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
         Left            =   300
         TabIndex        =   26
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label lblDataParametersLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Water Depth"
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
         Left            =   300
         TabIndex        =   25
         Top             =   2220
         Width           =   1815
      End
      Begin VB.Label lblDataParametersLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Water Volume"
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
         Left            =   300
         TabIndex        =   24
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label lblDataParametersLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "C*, 20"
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
         Left            =   4620
         TabIndex        =   23
         Top             =   540
         Width           =   1815
      End
      Begin VB.Label lblDataParametersLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Apparent KLa, 20"
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
         Index           =   7
         Left            =   3960
         TabIndex        =   22
         Top             =   1020
         Width           =   2475
      End
      Begin VB.Label lblDataParametersLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Phi"
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
         Index           =   8
         Left            =   4560
         TabIndex        =   21
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label lblDataParametersLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "True KLa, 20"
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
         Index           =   9
         Left            =   4560
         TabIndex        =   20
         Top             =   1860
         Width           =   1815
      End
      Begin VB.Label lblDataParametersLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Theta"
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
         Index           =   10
         Left            =   4560
         TabIndex        =   19
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label lblDataParametersLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "KLa, O2 at Operating T"
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
         Index           =   11
         Left            =   4320
         TabIndex        =   18
         Top             =   2700
         Width           =   2115
      End
      Begin VB.Label lblDataParameters 
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
         Left            =   6720
         TabIndex        =   17
         Top             =   540
         Width           =   1455
      End
      Begin VB.Label lblDataParameters 
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
         Index           =   7
         Left            =   6720
         TabIndex        =   16
         Top             =   1020
         Width           =   1455
      End
      Begin VB.Label lblDataParameters 
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
         Index           =   8
         Left            =   6720
         TabIndex        =   15
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label lblDataParameters 
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
         Index           =   9
         Left            =   6720
         TabIndex        =   14
         Top             =   1860
         Width           =   1455
      End
      Begin VB.Label lblDataParameters 
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
         Index           =   10
         Left            =   6720
         TabIndex        =   13
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label lblDataParameters 
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
         Index           =   11
         Left            =   6720
         TabIndex        =   12
         Top             =   2700
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmOxygenMassTransferCoeff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit





Const frmOxygenMassTransferCoeff_decl_end = True


Sub cmdOK_Click()
    
    If optDataAvailable(2).value Then
       frmBubble!cboOxygen.ListIndex = 1
       frmOxygenMassTransferCoeff.Hide
       Exit Sub
    Else
       bub.Oxygen.MassTransferCoefficient.value = bub.Oxygen.CWO2TestData.TrueOxygenMTCoeffOperatingT_KLAO2
       bub.Oxygen.KLaMethod = KLA_METHOD_CWO2_TRANSFER_TEST
       bub.Oxygen.MassTransferCoefficient.UserInput = False
       frmBubble!txtOxygen(2).Text = Format$(bub.Oxygen.MassTransferCoefficient.value, GetTheFormat(bub.Oxygen.MassTransferCoefficient.value))
    End If

    If bub.NumChemical > 0 Then
       Call CalculateContaminantMTCoeff

       If BubbleAerationMode = DESIGN_MODE Then
          Call CalculateTankVolumeBubble
          Call CalculateRetentionTimesAndTankVolumes
       End If

       Call CalculateStantonNo
       Call CalculateEffluentConcentrationsBubble
    End If

    frmOxygenMassTransferCoeff.Hide
End Sub

Sub Form_Activate()
    'Center the form on the screen
    If WindowState = 0 Then
       'don't attempt if screen Minimized or Maximized
      If frmPTADScreen1.Visible = True Then Move frmPTADScreen1.Left + (frmPTADScreen1.Width / 2) - (frmOxygenMassTransferCoeff.Width / 2), frmPTADScreen1.Top + (frmPTADScreen1.Height / 2) - (frmOxygenMassTransferCoeff.Height / 2)
      If frmPTADScreen2.Visible = True Then Move frmPTADScreen2.Left + (frmPTADScreen2.Width / 2) - (frmOxygenMassTransferCoeff.Width / 2), frmPTADScreen2.Top + (frmPTADScreen2.Height / 2) - (frmOxygenMassTransferCoeff.Height / 2)
      If frmBubble.Visible = True Then Move frmBubble.Left + (frmBubble.Width / 2) - (frmOxygenMassTransferCoeff.Width / 2), frmBubble.Top + (frmBubble.Height / 2) - (frmOxygenMassTransferCoeff.Height / 2)
      If frmSurface.Visible = True Then Move frmSurface.Left + (frmSurface.Width / 2) - (frmOxygenMassTransferCoeff.Width / 2), frmSurface.Top + (frmSurface.Height / 2) - (frmOxygenMassTransferCoeff.Height / 2)
    End If

    Call InitializeCWO2TestData

End Sub

Sub Form_Load()


    Call LabelsBubbleKLaO2SI

End Sub

Sub optDataAvailable_Click(index As Integer, value As Integer)
    Dim i As Integer

    Select Case index
       Case 0   'SOTR vs. QAIR Data Available
          txtDataParameters(0).Enabled = False
          txtDataParameters(1).Enabled = True
          For i = 2 To 5
              If Not txtDataParameters(i).Enabled Then txtDataParameters(i).Enabled = True
          Next i
          For i = 6 To 11
              If Not lblDataParameters(i).Enabled Then lblDataParameters(i).Enabled = True
          Next i

          
       Case 1   'SOTE vs. QAIR Data Available
          txtDataParameters(0).Enabled = True
          txtDataParameters(1).Enabled = False
          For i = 2 To 5
              If Not txtDataParameters(i).Enabled Then txtDataParameters(i).Enabled = True
          Next i
          For i = 6 To 11
              If Not lblDataParameters(i).Enabled Then lblDataParameters(i).Enabled = True
          Next i

          Call CalculateSOTR
          Call CalculateDOSaturationConc
          Call CalculateApparentKLa
          Call CalculateTrueKLa


       Case 2   'No data available
          
          For i = 0 To 5
              txtDataParameters(i).Enabled = False
          Next i
          For i = 6 To 11
              lblDataParameters(i).Enabled = False
          Next i
     
     End Select

End Sub

Sub txtDataParameters_GotFocus(index As Integer)
  Call GotFocus_Handle(Me, txtDataParameters(index), Temp_Text)

End Sub

Sub txtDataParameters_KeyPress(index As Integer, KeyAscii As Integer)
  Call TextBoxNumber_KeyPress(KeyAscii)
End Sub

Sub txtDataParameters_LostFocus(index As Integer)
    Dim Answer As Integer, Response As Integer
    Dim msg As String
    Dim Dummy As Double
    Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtDataParameters(index))) Then
     Exit Sub
   End If
   
   flag_ok = True

    Call TextHandleError(IsError, txtDataParameters(index), Temp_Text)
    If Not IsError Then
       If Not HaveValue(CDbl(txtDataParameters(index).Text)) Then
          txtDataParameters(index).Text = Temp_Text
          txtDataParameters(index).SetFocus
          Exit Sub
       End If
       
       Dummy = CDbl(txtDataParameters(index).Text)
       Select Case index
          Case 0   'SOTE
             Call TextNumberChanged(bub.Oxygen.CWO2TestData.SOTE.ValChanged, txtDataParameters(0), Temp_Text)

             If bub.Oxygen.CWO2TestData.SOTE.ValChanged Then
                bub.Oxygen.CWO2TestData.SOTE.value = Dummy
             Else
                 Call LostFocus_Handle(Me, txtDataParameters(index), flag_ok)
                Exit Sub
             End If

             Call CalculateSOTR
             Call CalculateApparentKLa
             Call CalculateTrueKLa

             If bub.NumChemical > 0 Then
                   'Update Variables on Screen
             End If
                                            

          Case 1   'SOTR
             Call TextNumberChanged(bub.Oxygen.CWO2TestData.SOTR.ValChanged, txtDataParameters(1), Temp_Text)

             If bub.Oxygen.CWO2TestData.SOTR.ValChanged Then
                bub.Oxygen.CWO2TestData.SOTR.value = Dummy
             Else
                 Call LostFocus_Handle(Me, txtDataParameters(index), flag_ok)
                Exit Sub
             End If

             Call CalculateSOTE
             Call CalculateApparentKLa
             Call CalculateTrueKLa

             If bub.NumChemical > 0 Then
                   'Update Variables on Screen
             End If


          Case 2   'Air Flow Rate
             Call TextNumberChanged(bub.Oxygen.CWO2TestData.AirFlowRate_QAIR.ValChanged, txtDataParameters(2), Temp_Text)

             If bub.Oxygen.CWO2TestData.AirFlowRate_QAIR.ValChanged Then
                bub.Oxygen.CWO2TestData.AirFlowRate_QAIR.value = Dummy
             Else
                 Call LostFocus_Handle(Me, txtDataParameters(index), flag_ok)
                Exit Sub
             End If

             If frmOxygenMassTransferCoeff!optDataAvailable(1).value = True Then
                Call CalculateSOTR
             End If

             Call CalculateTrueKLa

             If bub.NumChemical > 0 Then
                   'Update Variables on Screen
             End If
            

          Case 3   'Barometric Pressure
             Call TextNumberChanged(bub.Oxygen.CWO2TestData.BarometricPressure_PB.ValChanged, txtDataParameters(3), Temp_Text)

             If bub.Oxygen.CWO2TestData.BarometricPressure_PB.ValChanged Then
                bub.Oxygen.CWO2TestData.BarometricPressure_PB.value = Dummy * 1# / 101325#
             Else
                 Call LostFocus_Handle(Me, txtDataParameters(index), flag_ok)
                Exit Sub
             End If

             Call CalculateDOSaturationConc
             Call CalculateApparentKLa
             Call CalculateTrueKLa

             If bub.NumChemical > 0 Then
                   'Update Variables on Screen
             End If
            


          Case 4   'Water Depth
             Call TextNumberChanged(bub.Oxygen.CWO2TestData.WaterDepth_DEPTHW.ValChanged, txtDataParameters(4), Temp_Text)

             If bub.Oxygen.CWO2TestData.WaterDepth_DEPTHW.ValChanged Then
                bub.Oxygen.CWO2TestData.WaterDepth_DEPTHW.value = Dummy
             Else
                 Call LostFocus_Handle(Me, txtDataParameters(index), flag_ok)
                Exit Sub
             End If

             Call CalculateDOSaturationConc
             Call CalculateApparentKLa
             Call CalculateTrueKLa

             If bub.NumChemical > 0 Then
                   'Update Variables on Screen
             End If
          
          
          Case 5   'Water Volume
             Call TextNumberChanged(bub.Oxygen.CWO2TestData.WaterVolumePerTank_VM3.ValChanged, txtDataParameters(5), Temp_Text)

             If bub.Oxygen.CWO2TestData.WaterVolumePerTank_VM3.ValChanged Then
                bub.Oxygen.CWO2TestData.WaterVolumePerTank_VM3.value = Dummy
              Else
                Call LostFocus_Handle(Me, txtDataParameters(index), flag_ok)
                Exit Sub
             End If

             Call CalculateApparentKLa
             Call CalculateTrueKLa

             If bub.NumChemical > 0 Then
                   'Update Variables on Screen
             End If

        End Select
    End If

  Call LostFocus_Handle(Me, txtDataParameters(index), flag_ok)

End Sub


