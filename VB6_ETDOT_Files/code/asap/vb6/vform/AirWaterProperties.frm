VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmAirWaterProperties 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Properties of Air and Water"
   ClientHeight    =   3990
   ClientLeft      =   1830
   ClientTop       =   3675
   ClientWidth     =   7920
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtAirWaterProperties 
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
      Left            =   4680
      TabIndex        =   0
      Top             =   810
      Width           =   1215
   End
   Begin VB.TextBox txtAirWaterProperties 
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
      Left            =   4680
      TabIndex        =   1
      Top             =   1350
      Width           =   1215
   End
   Begin VB.TextBox txtAirWaterProperties 
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
      Left            =   4680
      TabIndex        =   2
      Top             =   1830
      Width           =   1215
   End
   Begin VB.TextBox txtAirWaterProperties 
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
      Left            =   4680
      TabIndex        =   3
      Top             =   2310
      Width           =   1215
   End
   Begin VB.TextBox txtAirWaterProperties 
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
      Left            =   4680
      TabIndex        =   4
      Top             =   2790
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
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
      Height          =   495
      Left            =   3720
      TabIndex        =   10
      Top             =   3390
      Width           =   615
   End
   Begin Threed.SSCheck chkUpdateValues 
      Height          =   255
      Index           =   4
      Left            =   600
      TabIndex        =   5
      Top             =   2790
      Width           =   255
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   -1  'True
   End
   Begin Threed.SSCheck chkUpdateValues 
      Height          =   255
      Index           =   3
      Left            =   600
      TabIndex        =   6
      Top             =   2310
      Width           =   255
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   -1  'True
   End
   Begin Threed.SSCheck chkUpdateValues 
      Height          =   255
      Index           =   2
      Left            =   600
      TabIndex        =   7
      Top             =   1830
      Width           =   255
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   -1  'True
   End
   Begin Threed.SSCheck chkUpdateValues 
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   8
      Top             =   1350
      Width           =   255
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   -1  'True
   End
   Begin Threed.SSCheck chkUpdateValues 
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   9
      Top             =   870
      Width           =   255
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   -1  'True
   End
   Begin VB.Shape Shape1 
      Height          =   3015
      Left            =   240
      Top             =   150
      Width           =   7455
   End
   Begin VB.Label lblAirWaterProperties 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Water Density"
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
      Left            =   1560
      TabIndex        =   24
      Top             =   870
      Width           =   2895
   End
   Begin VB.Label lblAirWaterProperties 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Water Viscosity"
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
      Left            =   1560
      TabIndex        =   23
      Top             =   1350
      Width           =   2895
   End
   Begin VB.Label lblAirWaterProperties 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Water Surface Tension"
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
      Left            =   1560
      TabIndex        =   22
      Top             =   1830
      Width           =   2895
   End
   Begin VB.Label lblAirWaterProperties 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Air Density"
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
      Left            =   1560
      TabIndex        =   21
      Top             =   2310
      Width           =   2895
   End
   Begin VB.Label lblAirWaterProperties 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Air Viscosity"
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
      Left            =   1560
      TabIndex        =   20
      Top             =   2790
      Width           =   2895
   End
   Begin VB.Label lblUpdateValues 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "label"
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
      Height          =   495
      Left            =   240
      TabIndex        =   19
      Top             =   270
      Width           =   975
   End
   Begin VB.Label lblSourceofValues 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "label"
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
      Height          =   375
      Left            =   6240
      TabIndex        =   18
      Top             =   270
      Width           =   1215
   End
   Begin VB.Label lblValueSource 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Left            =   6240
      TabIndex        =   17
      Top             =   870
      Width           =   1215
   End
   Begin VB.Label lblValueSource 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Left            =   6240
      TabIndex        =   16
      Top             =   1350
      Width           =   1215
   End
   Begin VB.Label lblValueSource 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Left            =   6240
      TabIndex        =   15
      Top             =   1830
      Width           =   1215
   End
   Begin VB.Label lblValueSource 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Left            =   6240
      TabIndex        =   14
      Top             =   2310
      Width           =   1215
   End
   Begin VB.Label lblValueSource 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Left            =   6240
      TabIndex        =   13
      Top             =   2790
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   7680
      Y1              =   750
      Y2              =   750
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   7680
      Y1              =   1230
      Y2              =   1230
   End
   Begin VB.Line Line3 
      X1              =   240
      X2              =   7680
      Y1              =   1710
      Y2              =   1710
   End
   Begin VB.Line Line4 
      X1              =   240
      X2              =   7680
      Y1              =   2190
      Y2              =   2190
   End
   Begin VB.Line Line5 
      X1              =   240
      X2              =   7680
      Y1              =   2670
      Y2              =   2670
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Property"
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
      Left            =   1560
      TabIndex        =   12
      Top             =   390
      Width           =   2895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Value"
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
      Left            =   4680
      TabIndex        =   11
      Top             =   390
      Width           =   1215
   End
End
Attribute VB_Name = "frmAirWaterProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim AtLeastOneValueChanged As Integer

Private Sub Command1_Click()

    frmAirWaterProperties.Hide

    If CurrentMode = 1 Then
       scr1 = CurrentScreen

      If scr1.NumChemical > 0 Then
        'Update Variables on Screen
        Call GetVQmultVQAndAirFlowRate
        Call GetLoadings
        Call GetTowerAreaAndDiameter
        Call GetOndaMassTransferCoefficient
        Call GetDesignKLaOrKLaSafetyFactor
        Call GetTowerHeightAndVolume
      End If

    Else
       Scr2 = CurrentScreen
       Call GetFlowsAndLoadingsScreen2

      If (Scr2.NumChemical > 0) Then
        'Update Variables on Screen
        Call GetContaminantConcentrationsScreen2
             
      End If

    End If

End Sub

Private Sub Form_Activate()
  Call CenterThisForm(Me)
End Sub

Private Sub Form_Load()

    Call CenterThisForm(Me)
    
    AtLeastOneValueChanged = False
    lblUpdateValues.Caption = "Update" & Chr$(13) & "Values"
    lblSourceofValues.Caption = "Source of" & Chr$(13) & "Values"
    Call LabelsAirWaterPropertiesSI
    
End Sub

Private Sub txtAirWaterProperties_GotFocus(Index As Integer)
    
  Call GotFocus_Handle(Me, txtAirWaterProperties(Index), Temp_Text)
End Sub

Private Sub txtAirWaterProperties_KeyPress(Index As Integer, KeyAscii As Integer)
  Call TextBoxNumber_KeyPress(KeyAscii)
End Sub

Private Sub txtAirWaterProperties_LostFocus(Index As Integer)

    Dim ValueChanged As Integer
    Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtAirWaterProperties(Index))) Then
     Exit Sub
   End If
   
   flag_ok = True

    Call TextHandleError(IsError, txtAirWaterProperties(Index), Temp_Text)

    If Not IsError Then
       'Make sure number in text box is > zero
       If Not HaveValue(CDbl(txtAirWaterProperties(Index).Text)) Then
          txtAirWaterProperties(Index).Text = Temp_Text
          txtAirWaterProperties(Index).SetFocus
          Exit Sub
       End If

       Call TextNumberChanged(ValueChanged, txtAirWaterProperties(Index), Temp_Text)

       If ValueChanged Then lblValueSource(Index).Caption = "User-Input"

       Select Case Index
          Case 0
               CurrentScreen.WaterDensity.ValChanged = ValueChanged
               CurrentScreen.WaterDensity.UserInput = ValueChanged
               If CurrentScreen.WaterDensity.UserInput Then CurrentScreen.WaterDensity.Value = CDbl(txtAirWaterProperties(0).Text)
          Case 1
               CurrentScreen.WaterViscosity.ValChanged = ValueChanged
               CurrentScreen.WaterViscosity.UserInput = ValueChanged
               If CurrentScreen.WaterViscosity.UserInput Then CurrentScreen.WaterViscosity.Value = CDbl(txtAirWaterProperties(1).Text)
          Case 2
               CurrentScreen.WaterSurfaceTension.ValChanged = ValueChanged
               CurrentScreen.WaterSurfaceTension.UserInput = ValueChanged
               If CurrentScreen.WaterSurfaceTension.UserInput Then CurrentScreen.WaterSurfaceTension.Value = CDbl(txtAirWaterProperties(2).Text)
          Case 3
               CurrentScreen.AirDensity.ValChanged = ValueChanged
               CurrentScreen.AirDensity.UserInput = ValueChanged
               If CurrentScreen.AirDensity.UserInput Then CurrentScreen.AirDensity.Value = CDbl(txtAirWaterProperties(3).Text)
          Case 4
               CurrentScreen.AirViscosity.ValChanged = ValueChanged
               CurrentScreen.AirViscosity.UserInput = ValueChanged
               If CurrentScreen.AirViscosity.UserInput Then CurrentScreen.AirViscosity.Value = CDbl(txtAirWaterProperties(4).Text)
       End Select
    End If

  Call LostFocus_Handle(Me, txtAirWaterProperties(Index), flag_ok)


End Sub


