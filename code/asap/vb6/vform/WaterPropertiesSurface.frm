VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmWaterPropertiesSurface 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Properties of Water"
   ClientHeight    =   2940
   ClientLeft      =   1830
   ClientTop       =   6090
   ClientWidth     =   7905
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   7905
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
      Top             =   780
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
      Top             =   1260
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
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2040
      Width           =   615
   End
   Begin Threed.SSCheck chkUpdateValues 
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   2
      Top             =   1260
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
      TabIndex        =   3
      Top             =   780
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
      TabIndex        =   15
      Top             =   780
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
      TabIndex        =   14
      Top             =   1260
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
      TabIndex        =   13
      Top             =   180
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
      TabIndex        =   12
      Top             =   180
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
      TabIndex        =   11
      Top             =   780
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
      TabIndex        =   10
      Top             =   1260
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
      TabIndex        =   9
      Top             =   1740
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
      TabIndex        =   8
      Top             =   2220
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
      TabIndex        =   7
      Top             =   2700
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      Height          =   1695
      Left            =   240
      Top             =   60
      Width           =   7455
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   7680
      Y1              =   660
      Y2              =   660
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   7680
      Y1              =   1140
      Y2              =   1140
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
      TabIndex        =   6
      Top             =   300
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
      TabIndex        =   5
      Top             =   300
      Width           =   1215
   End
End
Attribute VB_Name = "frmWaterPropertiesSurface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    frmWaterPropertiesSurface.Hide
End Sub

Private Sub Form_Activate()
    frmWaterPropertiesSurface.WindowState = 0

    'Position the form on the screen (Centered)
    If WindowState = 0 Then
       'don't attempt if screen Minimized or Maximized
      If frmPTADScreen1.Visible = True Then Move frmPTADScreen1.Left + (frmPTADScreen1.Width / 2) - (frmWaterPropertiesSurface.Width / 2), frmPTADScreen1.Top + (frmPTADScreen1.Height / 2) - (frmWaterPropertiesSurface.Height / 2)
      If frmPTADScreen2.Visible = True Then Move frmPTADScreen2.Left + (frmPTADScreen2.Width / 2) - (frmWaterPropertiesSurface.Width / 2), frmPTADScreen2.Top + (frmPTADScreen2.Height / 2) - (frmWaterPropertiesSurface.Height / 2)
      If frmBubble.Visible = True Then Move frmBubble.Left + (frmBubble.Width / 2) - (frmWaterPropertiesSurface.Width / 2), frmBubble.Top + (frmBubble.Height / 2) - (frmWaterPropertiesSurface.Height / 2)
      If frmSurface.Visible = True Then Move frmSurface.Left + (frmSurface.Width / 2) - (frmWaterPropertiesSurface.Width / 2), frmSurface.Top + (frmSurface.Height / 2) - (frmWaterPropertiesSurface.Height / 2)
    End If

End Sub

Private Sub Form_Load()


    lblUpdateValues.Caption = "Update" & Chr$(13) & "Values"
    lblSourceofValues.Caption = "Source of" & Chr$(13) & "Values"

    Call LabelsWaterPropertiesSurfaceSI
    
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
               sur.WaterDensity.ValChanged = ValueChanged
               sur.WaterDensity.UserInput = ValueChanged
               If sur.WaterDensity.UserInput Then sur.WaterDensity.Value = CDbl(txtAirWaterProperties(0).Text)
          Case 1
               sur.WaterViscosity.ValChanged = ValueChanged
               sur.WaterViscosity.UserInput = ValueChanged
               If sur.WaterViscosity.UserInput Then sur.WaterViscosity.Value = CDbl(txtAirWaterProperties(1).Text)
       End Select
    End If

  Call LostFocus_Handle(Me, txtAirWaterProperties(Index), flag_ok)


End Sub


