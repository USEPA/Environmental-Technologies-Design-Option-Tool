VERSION 5.00
Begin VB.Form frmantoine 
   Caption         =   "Fitting and Units Conversion - Antoine Coefficients"
   ClientHeight    =   6810
   ClientLeft      =   2715
   ClientTop       =   1680
   ClientWidth     =   9495
   ControlBox      =   0   'False
   LinkTopic       =   "FRMAntoine"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6810
   ScaleWidth      =   9495
   Begin VB.CommandButton CMDaccept 
      Caption         =   "Accept"
      Height          =   375
      Left            =   4800
      TabIndex        =   83
      Top             =   3480
      Width           =   1605
   End
   Begin VB.CommandButton CMDDIPPR 
      Caption         =   "DIPPR 801/911"
      Height          =   375
      Left            =   2430
      TabIndex        =   19
      Top             =   990
      Width           =   1785
   End
   Begin VB.Frame FrameInputs 
      Height          =   3950
      Left            =   4560
      TabIndex        =   1
      Top             =   0
      Width           =   4450
      Begin VB.CommandButton CMDCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   2580
         TabIndex        =   82
         Top             =   3480
         Width           =   1635
      End
      Begin VB.ComboBox CMBEquationForm 
         Height          =   315
         Left            =   1680
         TabIndex        =   40
         Text            =   "CMBEquationForm"
         Top             =   1710
         Width           =   2445
      End
      Begin VB.ComboBox CMBVPUnits 
         Height          =   315
         ItemData        =   "frmantoi.frx":0000
         Left            =   2160
         List            =   "frmantoi.frx":0002
         TabIndex        =   31
         Text            =   "CMBVPUnits"
         Top             =   420
         Width           =   1515
      End
      Begin VB.ComboBox CMBTempUnits 
         Height          =   315
         Left            =   1995
         TabIndex        =   30
         Text            =   "CMBTempUnits"
         Top             =   840
         Width           =   885
      End
      Begin VB.TextBox TXTTempFrom 
         Height          =   285
         Left            =   1650
         TabIndex        =   29
         Text            =   "TXTTempFrom"
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox TXTTempTo 
         Height          =   285
         Left            =   2880
         TabIndex        =   28
         Text            =   "TXTTempTo"
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox TXTRegressionPoints 
         Height          =   285
         Left            =   2190
         TabIndex        =   27
         Text            =   "TXTRegressionPoints"
         Top             =   2580
         Width           =   1245
      End
      Begin VB.CommandButton CMDConvert 
         Caption         =   "Convert Units"
         Height          =   375
         Left            =   2580
         TabIndex        =   26
         Top             =   3060
         Width           =   1605
      End
      Begin VB.ComboBox CMBLogForm 
         Height          =   315
         Left            =   1260
         TabIndex        =   25
         Text            =   "CMBLogForm"
         Top             =   1260
         Width           =   1275
      End
      Begin VB.CommandButton CMDRegress 
         Caption         =   "Regress"
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   3060
         Width           =   1605
      End
      Begin VB.Label Label8 
         Caption         =   "K"
         Height          =   285
         Left            =   3840
         TabIndex        =   86
         Top             =   2190
         Width           =   465
      End
      Begin VB.Label LBLEquationForm 
         Caption         =   "Equation Output Form:"
         Height          =   225
         Left            =   60
         TabIndex        =   39
         Top             =   1740
         Width           =   1635
      End
      Begin VB.Label LBLVPUnits 
         Caption         =   "Vapor Pressure Output Units:"
         Height          =   225
         Left            =   60
         TabIndex        =   38
         Top             =   480
         Width           =   2145
      End
      Begin VB.Label LBLTitleAntoine 
         Alignment       =   2  'Center
         Caption         =   "Antoine Inputs"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   30
         TabIndex        =   37
         Top             =   120
         Width           =   4365
      End
      Begin VB.Label LBLTempUnits 
         Caption         =   "Temperature Output Units:"
         Height          =   225
         Left            =   60
         TabIndex        =   36
         Top             =   870
         Width           =   1935
      End
      Begin VB.Label LBLTempRange 
         Caption         =   "Temperature Range:"
         Height          =   225
         Left            =   60
         TabIndex        =   35
         Top             =   2190
         Width           =   1545
      End
      Begin VB.Label LBLTempTo 
         Caption         =   "--"
         Height          =   195
         Left            =   2640
         TabIndex        =   34
         Top             =   2190
         Width           =   195
      End
      Begin VB.Label LBLRegPoints 
         Caption         =   "Number of Regression Points:"
         Height          =   255
         Left            =   60
         TabIndex        =   33
         Top             =   2610
         Width           =   2175
      End
      Begin VB.Label LBLLogForm 
         Caption         =   "Logarithm Form:"
         Height          =   225
         Left            =   60
         TabIndex        =   32
         Top             =   1290
         Width           =   1155
      End
   End
   Begin VB.Frame FrameVP 
      Height          =   3950
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4450
      Begin VB.VScrollBar VertScrollEquation 
         Height          =   915
         Left            =   4080
         TabIndex        =   23
         Top             =   2940
         Width           =   255
      End
      Begin VB.TextBox TXTEquation 
         Height          =   915
         Left            =   810
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Text            =   "TXTEquation"
         Top             =   2940
         Width           =   3285
      End
      Begin VB.ComboBox CMBEquationNumber 
         Height          =   315
         Left            =   2190
         TabIndex        =   21
         Text            =   "CMBEquationNumber"
         Top             =   2580
         Width           =   1300
      End
      Begin VB.CommandButton CMDAntoine 
         Caption         =   "Antoine "
         Height          =   375
         Left            =   300
         TabIndex        =   18
         Top             =   990
         Width           =   1785
      End
      Begin VB.TextBox TXTChemName 
         Height          =   285
         Left            =   1350
         TabIndex        =   8
         Text            =   "TXTChemName"
         Top             =   390
         Width           =   3015
      End
      Begin VB.TextBox TXTStartingCoeffA 
         Height          =   285
         Left            =   300
         TabIndex        =   7
         Text            =   "TXTStartingCoeffA"
         Top             =   1500
         Width           =   1500
      End
      Begin VB.TextBox TXTStartingCoeffC 
         Height          =   285
         Left            =   300
         TabIndex        =   6
         Text            =   "TXTStartingCoeffC"
         Top             =   2220
         Width           =   1500
      End
      Begin VB.TextBox TXTStartingCoeffB 
         Height          =   285
         Left            =   300
         TabIndex        =   5
         Text            =   "TXTStartingCoeffB"
         Top             =   1860
         Width           =   1500
      End
      Begin VB.TextBox TXTStartingCoeffE 
         Height          =   285
         Left            =   2460
         TabIndex        =   4
         Text            =   "TXTStartingCoeffE"
         Top             =   1860
         Width           =   1500
      End
      Begin VB.TextBox TXTStartingCoeffD 
         Height          =   285
         Left            =   2460
         TabIndex        =   3
         Text            =   "TXTStartingCoeffD"
         Top             =   1500
         Width           =   1500
      End
      Begin VB.Label LBLChooseCoeff 
         Caption         =   "Choose a Coefficient File:"
         Height          =   195
         Left            =   1290
         TabIndex        =   20
         Top             =   750
         Width           =   1845
      End
      Begin VB.Label LBLVP 
         Alignment       =   2  'Center
         Caption         =   "Vapor Pressure Correlations"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   105
         Width           =   4305
      End
      Begin VB.Label LBLName 
         Caption         =   "Chemical Name:"
         Height          =   225
         Left            =   90
         TabIndex        =   16
         Top             =   420
         Width           =   1215
      End
      Begin VB.Label LBLStartingCoeffA 
         Caption         =   "A:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   1530
         Width           =   225
      End
      Begin VB.Label LBLStartingCoeffB 
         Caption         =   "B:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   1890
         Width           =   225
      End
      Begin VB.Label LBLStartingCoeffC 
         Caption         =   "C:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   2280
         Width           =   225
      End
      Begin VB.Label LBLStartingCoeffD 
         Caption         =   "D:"
         Height          =   195
         Left            =   2250
         TabIndex        =   12
         Top             =   1530
         Width           =   225
      End
      Begin VB.Label LBLEquationNumber 
         Caption         =   "Equation #:"
         Height          =   225
         Left            =   1350
         TabIndex        =   11
         Top             =   2610
         Width           =   915
      End
      Begin VB.Label LBLEquation 
         Caption         =   "Equation:"
         Height          =   225
         Left            =   120
         TabIndex        =   10
         Top             =   2940
         Width           =   735
      End
      Begin VB.Label LBLStartingCoeffE 
         Caption         =   "E:"
         Height          =   195
         Left            =   2250
         TabIndex        =   9
         Top             =   1890
         Width           =   225
      End
   End
   Begin VB.Frame FrameRegress 
      Height          =   1905
      Left            =   120
      TabIndex        =   47
      Top             =   4080
      Width           =   6420
      Begin VB.CommandButton CMDRecalc 
         Caption         =   "Recalculate"
         Height          =   375
         Left            =   4740
         TabIndex        =   87
         Top             =   1320
         Width           =   1635
      End
      Begin VB.TextBox TXTANTTemp 
         Height          =   285
         Left            =   3420
         TabIndex        =   85
         Text            =   "TXTANTTemp"
         Top             =   1380
         Width           =   585
      End
      Begin VB.TextBox TXTANTPressUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   4560
         TabIndex        =   80
         Text            =   "TXTANTPressUnits"
         Top             =   990
         Width           =   645
      End
      Begin VB.TextBox TXTANTC 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   315
         TabIndex        =   68
         Text            =   "TXTANTC"
         Top             =   1575
         Width           =   1800
      End
      Begin VB.TextBox TXTANTB 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   315
         TabIndex        =   67
         Text            =   "TXTANTB"
         Top             =   1260
         Width           =   1800
      End
      Begin VB.TextBox TXTANTTempUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   4080
         TabIndex        =   59
         Text            =   "TXTANTTempUnits"
         Top             =   1410
         Width           =   645
      End
      Begin VB.TextBox TXTANTA 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   315
         TabIndex        =   56
         Text            =   "TXTANTA"
         Top             =   945
         Width           =   1800
      End
      Begin VB.TextBox TXTANTEquation 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   3045
         TabIndex        =   55
         Text            =   "TXTANTEquation"
         Top             =   420
         Width           =   2955
      End
      Begin VB.TextBox TXTANTEqnNum 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1470
         TabIndex        =   53
         Text            =   "TXTANTEqnNum"
         Top             =   420
         Width           =   645
      End
      Begin VB.Label TXTVaporPressure 
         Caption         =   "TXTVaporPressure"
         Height          =   255
         Left            =   3420
         TabIndex        =   81
         Top             =   990
         Width           =   1155
      End
      Begin VB.Label LBLANTTemp 
         Caption         =   "at Temperature:"
         Height          =   225
         Left            =   2205
         TabIndex        =   58
         Top             =   1395
         Width           =   1170
      End
      Begin VB.Label LBLVaporPressure 
         Caption         =   "Vapor Pressure:"
         Height          =   225
         Left            =   2205
         TabIndex        =   57
         Top             =   975
         Width           =   1170
      End
      Begin VB.Label LNLANTEqation 
         Caption         =   "Equation:"
         Height          =   225
         Left            =   2310
         TabIndex        =   54
         Top             =   420
         Width           =   750
      End
      Begin VB.Label LBLAntEqnNumber 
         Caption         =   "Equation Number:"
         Height          =   225
         Left            =   105
         TabIndex        =   52
         Top             =   420
         Width           =   1380
      End
      Begin VB.Label LBLANTC 
         Caption         =   "C:"
         Height          =   225
         Left            =   105
         TabIndex        =   51
         Top             =   1575
         Width           =   225
      End
      Begin VB.Label LBLANTB 
         Caption         =   "B:"
         Height          =   225
         Left            =   105
         TabIndex        =   50
         Top             =   1260
         Width           =   225
      End
      Begin VB.Label LBLANTA 
         Caption         =   "A:"
         Height          =   225
         Left            =   105
         TabIndex        =   49
         Top             =   945
         Width           =   225
      End
      Begin VB.Label LBLRegress 
         Alignment       =   2  'Center
         Caption         =   "Regression Results"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   105
         TabIndex        =   48
         Top             =   120
         Width           =   6210
      End
   End
   Begin VB.Frame FrameStatistics 
      Height          =   1905
      Left            =   6540
      TabIndex        =   2
      Top             =   4050
      Width           =   2430
      Begin VB.TextBox TXTRSQR 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   810
         TabIndex        =   91
         Text            =   "TXTRSQR"
         Top             =   540
         Width           =   1485
      End
      Begin VB.TextBox TXTRMPE 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   840
         TabIndex        =   46
         Text            =   "TXTRMPE"
         Top             =   1230
         Width           =   1485
      End
      Begin VB.Label TXTErr 
         Height          =   255
         Left            =   1440
         TabIndex        =   88
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label TXTRMSE 
         Caption         =   "TXTRMSE"
         Height          =   255
         Left            =   840
         TabIndex        =   84
         Top             =   870
         Width           =   885
      End
      Begin VB.Label LBLRMPE 
         Caption         =   "RMPE:"
         Height          =   225
         Left            =   105
         TabIndex        =   45
         Top             =   1215
         Width           =   540
      End
      Begin VB.Label LBLRMSE 
         Caption         =   "RMSE:"
         Height          =   225
         Left            =   105
         TabIndex        =   44
         Top             =   870
         Width           =   540
      End
      Begin VB.Label LBLSQR 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   210
         TabIndex        =   43
         Top             =   525
         Width           =   135
      End
      Begin VB.Label LBLR2 
         Caption         =   "R   :"
         Height          =   225
         Left            =   105
         TabIndex        =   42
         Top             =   525
         Width           =   330
      End
      Begin VB.Label LBLStatistics 
         Alignment       =   2  'Center
         Caption         =   "Regression Statistics"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   41
         Top             =   120
         Width           =   2220
      End
   End
   Begin VB.Frame FrameConvert 
      Height          =   1905
      Left            =   0
      TabIndex        =   62
      Top             =   4080
      Visible         =   0   'False
      Width           =   8940
      Begin VB.CommandButton CMDRecalc1 
         Caption         =   "Recalculate"
         Height          =   375
         Left            =   6840
         TabIndex        =   90
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox TXTTempUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   5400
         TabIndex        =   78
         Text            =   "TXTTempUnits"
         Top             =   1200
         Width           =   645
      End
      Begin VB.TextBox TXTTemp 
         Height          =   285
         Left            =   4320
         TabIndex        =   77
         Text            =   "TXTTemp"
         Top             =   1155
         Width           =   1065
      End
      Begin VB.TextBox TXTCC 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   315
         TabIndex        =   74
         Text            =   "TXTCC"
         Top             =   1365
         Width           =   1800
      End
      Begin VB.TextBox TXTBB 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   315
         TabIndex        =   73
         Text            =   "TXTBB"
         Top             =   1050
         Width           =   1800
      End
      Begin VB.TextBox TXTAA 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   315
         TabIndex        =   72
         Text            =   "TXTAA"
         Top             =   720
         Width           =   1800
      End
      Begin VB.TextBox TXTEquationNumber 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   3480
         TabIndex        =   66
         Text            =   "TXTANTEquation"
         Top             =   420
         Width           =   2955
      End
      Begin VB.TextBox TXTEqNum 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1470
         TabIndex        =   64
         Text            =   "TXTEqNum"
         Top             =   420
         Width           =   645
      End
      Begin VB.Label txtVP 
         Height          =   255
         Left            =   4200
         TabIndex        =   93
         Top             =   750
         Width           =   2055
      End
      Begin VB.Label TXTErr1 
         Height          =   255
         Left            =   8160
         TabIndex        =   92
         Top             =   600
         Width           =   855
      End
      Begin VB.Label TXTVPUnits 
         Height          =   255
         Left            =   6240
         TabIndex        =   89
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "at Temperature:"
         Height          =   225
         Left            =   3045
         TabIndex        =   76
         Top             =   1155
         Width           =   1170
      End
      Begin VB.Label Label6 
         Caption         =   "Vapor Pressure:"
         Height          =   225
         Left            =   3045
         TabIndex        =   75
         Top             =   750
         Width           =   1170
      End
      Begin VB.Label Label5 
         Caption         =   "C:"
         Height          =   225
         Left            =   105
         TabIndex        =   71
         Top             =   1365
         Width           =   225
      End
      Begin VB.Label Label4 
         Caption         =   "B:"
         Height          =   225
         Left            =   105
         TabIndex        =   70
         Top             =   1050
         Width           =   225
      End
      Begin VB.Label Label3 
         Caption         =   "A:"
         Height          =   225
         Left            =   105
         TabIndex        =   69
         Top             =   735
         Width           =   225
      End
      Begin VB.Label Label2 
         Caption         =   "Equation:"
         Height          =   225
         Left            =   2730
         TabIndex        =   65
         Top             =   420
         Width           =   750
      End
      Begin VB.Label Label1 
         Caption         =   "Equation Number:"
         Height          =   225
         Left            =   120
         TabIndex        =   63
         Top             =   420
         Width           =   1380
      End
      Begin VB.Label LBLConvertTitle 
         Alignment       =   2  'Center
         Caption         =   "Unit Conversion Output"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   105
         TabIndex        =   79
         Top             =   120
         Visible         =   0   'False
         Width           =   8940
      End
   End
   Begin VB.Frame FrameStart 
      Height          =   1905
      Left            =   60
      TabIndex        =   60
      Top             =   4080
      Visible         =   0   'False
      Width           =   8940
      Begin VB.Label LBLStart 
         Alignment       =   2  'Center
         Caption         =   "Fitting and Units Conversion - Antoine Coefficients"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   105
         TabIndex        =   61
         Top             =   420
         Width           =   8730
      End
   End
End
Attribute VB_Name = "frmantoine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Const Limit = 10


Private Sub CMBEquationNumber_Click()

Call antoine_equation_number

End Sub


Private Sub CMDAccept_Click()
    If frmantoine!txtVP.Visible = True Then
        InfoMethod(6).value(4) = Format(frmantoine!txtVP.caption)
        InfoMethod(6).Unit = frmantoine!TXTVPUnits.caption
    ElseIf frmantoine!TXTVaporPressure.Visible = True And frmantoine!TXTVaporPressure.caption <> "Null" Then
        InfoMethod(6).value(4) = Format(frmantoine!TXTVaporPressure.caption)
        InfoMethod(6).Unit = frmantoine!CMBVPUnits.Text
    Else
        InfoMethod(6).value(4) = 0
        InfoMethod(6).Unit = ""
    End If
    
    Call antoine_check_update_udb(Antoine_Info)
    
    Unload frmantoine
End Sub


Private Sub CMDAntoine_Click()
    On Error Resume Next
    Antoine_Info.MethodName = "Antoine"
    Call antoine_antoine
End Sub

Private Sub CMDCancel_Click()

    Unload frmantoine
End Sub

Private Sub CMDConvert_Click()
Call do_antoine_convert
Call FrameConvertOnTop
End Sub

Private Sub CMDConvintoDB_Click()
Call do_antoine_to_db
End Sub




Private Sub CMDDIPPR_Click()
On Error Resume Next
Antoine_Info.MethodName = "DIPPR 801/911"
Call antoine_dippr

End Sub






Private Sub CMDRecalc_Click()
    Call recalc_antoine
    CMDRegress_Click
End Sub

Private Sub CMDRecalc1_Click()

   Call recalc_one_antoine
   CMDConvert_Click

End Sub


Private Sub CMDRegress_Click()
    Call do_antoine_regress
End Sub











Public Sub FrameRegressOnTop()

FrameStart.Visible = False
FrameConvert.Visible = False

FrameRegress.Visible = True
FrameStatistics.Visible = True

End Sub


Public Sub FrameInputsDefault(TempFrom As Double, TempTo As Double, EqNum As Integer)
Call fill_antoine_input_defaults(TempFrom, TempTo, EqNum)
End Sub

Public Function FillEquationBox(Number As String) As String
Call fill_antoine_equation_box(Number)
End Function

Public Sub FrameConvertOnTop()
FrameStart.Visible = False
FrameConvert.Visible = True

FrameRegress.Visible = False
FrameStatistics.Visible = False

End Sub

Public Function CalcVP(EqNum As Integer, AAA As Double, BBB As Double, CCC As Double, TTT As Double)

Select Case EqNum
    Case 300
        CalcVP = 10 ^ (AAA - BBB / (TTT + CCC))
    Case 301
        CalcVP = 10 ^ (AAA - BBB / (TTT - CCC))
End Select
End Function

Private Sub intodb_Click()

Call antoine_into_db
End Sub




