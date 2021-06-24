VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0842D103-1E19-101B-9AAF-1A1626551E7C}#1.0#0"; "GRAPH32.OCX"
Begin VB.Form frmModelECMResults 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Results for the Equilibrium Column Model (ECM)"
   ClientHeight    =   6420
   ClientLeft      =   960
   ClientTop       =   1440
   ClientWidth     =   9405
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   9405
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   9120
      ScaleHeight     =   795
      ScaleWidth      =   1275
      TabIndex        =   22
      Top             =   5640
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Print Screen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7740
      TabIndex        =   21
      ToolTipText     =   "Click here to print current screen to selected printer"
      Top             =   5985
      Width           =   1455
   End
   Begin VB.ComboBox cboGlob 
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
      Height          =   315
      Left            =   7500
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   2910
      Width           =   1815
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   2775
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   9255
      _Version        =   65536
      _ExtentX        =   16325
      _ExtentY        =   4895
      _StockProps     =   14
      Caption         =   "Results:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "VTM         (mg GAC/L)"
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
         Left            =   5310
         TabIndex        =   14
         Top             =   210
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Wave velocity (cm/s)"
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
         Left            =   3930
         TabIndex        =   13
         Top             =   210
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Bed Volume Fed"
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
         Left            =   2790
         TabIndex        =   12
         Top             =   210
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Components"
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
         Left            =   510
         TabIndex        =   11
         Top             =   405
         Width           =   2235
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Zone"
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
         Left            =   90
         TabIndex        =   10
         Top             =   405
         Width           =   555
      End
      Begin VB.Label lblData3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111E22"
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
         Height          =   2055
         Left            =   5310
         TabIndex        =   9
         Top             =   630
         Width           =   1215
      End
      Begin VB.Label lblData2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1E-11"
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
         Height          =   2055
         Left            =   3930
         TabIndex        =   8
         Top             =   630
         Width           =   1335
      End
      Begin VB.Label lblData1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Height          =   2055
         Left            =   2790
         TabIndex        =   7
         Top             =   630
         Width           =   1095
      End
      Begin VB.Label lblZone 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
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
         Height          =   2055
         Left            =   90
         TabIndex        =   6
         Top             =   630
         Width           =   375
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Time to break through (days)"
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
         Left            =   6570
         TabIndex        =   5
         Top             =   210
         Width           =   1335
      End
      Begin VB.Label lblData4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblData4"
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
         Height          =   2055
         Left            =   6570
         TabIndex        =   4
         Top             =   630
         Width           =   1335
      End
      Begin VB.Label lblCompo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblCompo"
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
         Height          =   2055
         Left            =   510
         TabIndex        =   3
         Top             =   630
         Width           =   2235
      End
      Begin VB.Label lblData5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblData5"
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
         Height          =   2055
         Left            =   7950
         TabIndex        =   2
         Top             =   630
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Mass Bal. Err. (%)"
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
         Left            =   8040
         TabIndex        =   1
         Top             =   210
         Width           =   915
      End
   End
   Begin GraphLib.Graph grpGlob 
      Height          =   3435
      Left            =   60
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2910
      Width           =   7095
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   2646
      _StockProps     =   96
      BorderStyle     =   1
      RandomData      =   1
      ColorData       =   0
      ExtraData       =   0
      ExtraData[]     =   0
      FontFamily      =   4
      FontSize        =   4
      FontSize[0]     =   200
      FontSize[1]     =   150
      FontSize[2]     =   100
      FontSize[3]     =   100
      FontStyle       =   4
      GraphData       =   0
      GraphData[]     =   0
      LabelText       =   0
      LegendText      =   0
      PatternData     =   0
      SymbolData      =   0
      XPosData        =   0
      XPosData[]      =   0
   End
   Begin Threed.SSCommand cmdFile 
      Height          =   435
      Left            =   7740
      TabIndex        =   16
      Top             =   4650
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "Print to &File"
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog CMDialog1 
      Left            =   7200
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontSize        =   0
      MaxFileSize     =   256
   End
   Begin Threed.SSCommand cmdSelect 
      Height          =   435
      Left            =   7740
      TabIndex        =   17
      Top             =   3570
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "&Select Printer"
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSCommand cmdPrint 
      Height          =   435
      Left            =   7740
      TabIndex        =   18
      Top             =   4110
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "&Print"
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSCommand cmdClose 
      Height          =   435
      Left            =   7740
      TabIndex        =   20
      Top             =   5190
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "&Close"
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmModelECMResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Dim length As Double
Dim NumW As Integer
Dim Solid_ConcW(Number_Compo_Max, Number_Compo_Max), Liquid_ConcW(Number_Compo_Max, Number_Compo_Max), CoCW(Number_Compo_Max, Number_Compo_Max) As Double
Dim IndexW(Number_Compo_Max) As Integer
Dim Name_CompW(Number_Compo_Max) As String * 20
Dim Time_Break() As Double 'Breakthroug time for each chhemical
Dim Time_Min As Double, Time_Unit As String

Dim PopulatingScrollboxes As Integer




Const frmModelECMResults_declarations_end = True


Private Sub cboGlob_Click()
Dim i As Integer
    
  If (Not PopulatingScrollboxes) Then
    If cboGlob = "C/Co" Then i = 1
    If cboGlob = "Q" Then i = 2
    If cboGlob = "C (Liquid Conc.)" Then i = 3
    Call Draw(i)
    grpGlob.DrawMode = 2
  End If

End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub


Private Sub cmdFile_Click()
Dim f As Integer, Error_Code As Integer, temp As String
Dim i As Integer, J As Integer, k As Integer
Dim Filename_Input As String

On Error GoTo File_Error
    CMDialog1.CancelError = True
    CMDialog1.DialogTitle = "Print to File"
    CMDialog1.Filter = "All Files (*.*)|*.*|Text Files (*.txt)|*.txt|Data Files (*.dat)|*.dat"
    CMDialog1.FilterIndex = 2
    CMDialog1.flags = _
        cdlOFNOverwritePrompt + _
        cdlOFNPathMustExist
    CMDialog1.Action = 2

   'f = FileNameIsValid(Filename_Input, CMDialog1)
   'If Not (f) Then Exit Sub
      Filename_Input = CMDialog1.Filename
      f = FreeFile
      Open Filename_Input For Output As f
      Print #f, "Input data for the Equilibrium Colum Model"
    '-- Print Filename

      Print #f,
      Print #f, "From Data File : " & Filename
      Print #f, "Date/time stamp: " & Date$ & " " & Time$

      Print #f,
      Print #f, "Component"; Tab(30); "K*"; Tab(38); "1/n"; Tab(45); "Init. Conc."; Tab(59); "MW"
      Print #f, Tab(39); "-"; Tab(48); "mg/L", Tab(58); "g/mol"
                                                                     
      For i = 1 To Number_Component_ECM
        'K = Component_Index_ECM(i)
        k = IndexW(i)
'        Print #f, Trim$(Mid$(LTrim$(Component(K).Name), 1, 25)); Tab(29); Format$(Component(K).Use_K, "###,##0.000"); Tab(37); Format$(Component(K).Use_OneOverN, "0.000"); Tab(48); Format_It(Component(K).InitialConcentration, 2); Tab(58); Format$(Component(K).MW, "0.00")
        Print #f, Trim$(Mid$(LTrim$(Component(k).Name), 1, 25)); Tab(29); Format$(Component(k).Use_K, "###,##0.000"); Tab(37); Format$(Component(k).Use_OneOverN, "0.000"); Tab(48); Format_It(Component(k).InitialConcentration, 2); Tab(58); Format$(Component(k).MW, "0.00")
      Next i
      Print #f,
      Print #f, "* K in (mg/g)*(L/mg)^(1/n)"

      Print #f,

      '-----------------------Bed Data ----------------------
      Print #f, "Bed Data:"

      Print #f, Tab(5); "Bed Length: "; Tab(28); Format$(Bed.length, "0.000E+00") & " m"
      Print #f, Tab(5); "Bed Diameter: "; Tab(28); Format$(Bed.Diameter, "0.000E+00") & " m"
      Print #f, Tab(5); "Weight of GAC: "; Tab(28); Format$(Bed.Weight, "0.000E+00") & " kg"
      Print #f, Tab(5); "Inlet Flowrate: "; Tab(28); Format$(Bed.Flowrate, "0.000E+00") & " m" & Chr$(179) & "/s"
      Print #f, Tab(5); "EBCT: "; Tab(28); Format$(Bed.length * PI * Bed.Diameter * Bed.Diameter / 4# / Bed.Flowrate / 60#, "0.000E+00") & " mn"
      Print #f,
      Print #f, Tab(5); "Temperature:"; Tab(28); Format$(Bed.Temperature, "0.00") & " C"
      If Bed.Phase = 0 Then
        Print #f, Tab(5); "Water Density:"; Tab(28); Format$(Bed.WaterDensity, "0.0000") & " g/cm" & Chr$(179)
        Print #f, Tab(5); "Water Viscosity:"; Tab(28); Format$(Bed.WaterViscosity, "0.00E+00") & " g/cm.s"
      Else
        Print #f, Tab(5); "Pressure:"; Tab(28); Format$(Bed.Pressure, "0.00000") & " atm"
        Print #f, Tab(5); "Air Density:"; Tab(28); Format$(Bed.WaterDensity, "0.0000") & " g/cm" & Chr$(179)
        Print #f, Tab(5); "Air Viscosity:"; Tab(28); Format$(Bed.WaterViscosity, "0.00E+00") & " g/cm.s"
      End If
      Print #f,

      '-----------------Carbon Properties -------------------------------
      Print #f, "Carbon Properties:"

      Print #f, Tab(5); "Name: "; Tab(28); Trim$(Carbon.Name)
      Print #f, Tab(5); "Apparent Density: "; Tab(28); Format$(Carbon.Density, "0.000") & " g/cm" & Chr$(179)
      Print #f, Tab(5); "Particle Radius: "; Tab(28); Format$(Carbon.ParticleRadius * 100#, "0.000000") & " cm"
      Print #f, Tab(5); "Porosity: "; Tab(28); Format$(Carbon.Porosity, "0.000")
      Print #f, Tab(5); "Shape Factor: "; Tab(28); Format$(Carbon.ShapeFactor, "0.000")
      'Print #f, Tab(5); "Tortuosity: "; Tab(28); Format$(Carbon.Tortuosity, "0.000")
      Print #f,
      
      Print #f,
      '--- Print the results from the table
      Print #f, "Results for the Equilibrium Column Model"

      Print #f,
      Print #f, "Zone"; Tab(9); "Component"; Tab(35); "BVF"; Tab(44); "Wave Vel."; Tab(54); "TC"; Tab(63); "Breakthrough"
      Print #f, Tab(45); "cm/s"; Tab(54); "m3/kg"; Tab(63); Time_Unit
      For i = 1 To Number_Component_ECM
        Print #f, "Zone " & Format$(i, "0"); Tab(9); Mid$(Trim$(Component(IndexW(i)).Name), 1, 25); Tab(35); Format_It(Output_ECM(i).Bed_Volume_Fed, 2); Tab(45); Format_It(Output_ECM(i).Wave_Velocity, 2); Tab(54); Format_It(1 / Output_ECM(i).Carbon_Usage_Rate * 1000, 2); Tab(63); Format_It(Time_Break(i), 2)
        
      'Change made: (ejo, 3/1/96)
      '==========================
      'was: Format_It(Output_ECM(i).Carbon_Usage_Rate, 2)
      'is now: Format_It(1 / Output_ECM(i).Carbon_Usage_Rate * 1000, 2)
      
      Next i
      Print #f,
      Print #f, "TC (Treatment Capacity) is in m" & Chr$(179) & "  / kg of GAC"
      Print #f,

      For i = 1 To Number_Component_ECM
        Print #f, Mid$(Trim$(Component(IndexW(i)).Name), 1, 25)
        Print #f, "Zone "; Tab(9); "C/Co"; Tab(19); "C (mg/L)"; Tab(29); "Q (mg/L)"
        For J = 1 To Number_Component_ECM
          Print #f, "Zone " & Format$(J, "0"); Tab(9); Format_It(CoCW(i, J), 2); Tab(19); Format_It(Liquid_ConcW(i, J) / 1000#, 2); Tab(29); Format_It(Solid_ConcW(i, J) / 1000#, 2)
        Next J
        Print #f,
      Next i

      Print #f,
      Print #f,
      '--- Print the mass balance results
      Print #f, "Mass Balance Results"
      Print #f,
      Print #f, "Component"; Tab(30); "Percent Err."
      'Print #f, ""; Tab(30); "(ug/cm2/s)"; Tab(45); "(ug/cm2/s)"; Tab(60); "(%)"
      Print #f, ""; Tab(30); "(%)"
      For i = 1 To Number_Component_ECM
'        Print #f, Mid$(Trim$(Component(i).Name), 1, 25); Tab(30); Format_It(Output_ECM_MASSBAL.MASSBAL_C0_e_Vf(IndexW(i)), 3); Tab(45); Format_It(Output_ECM_MASSBAL.MASSBAL_TERM_SUM(IndexW(i)), 3); Tab(60); Format_It(Output_ECM_MASSBAL.MASSBAL_PERCENT_ERR(IndexW(i)), 3)
'        Print #f, Mid$(Trim$(Component(IndexW(i)).Name), 1, 25); Tab(30); Format_It(Output_ECM_MASSBAL.MASSBAL_C0_e_Vf(i), 3); Tab(45); Format_It(Output_ECM_MASSBAL.MASSBAL_TERM_SUM(i), 3); Tab(60); Format_It(Output_ECM_MASSBAL.MASSBAL_PERCENT_ERR(i), 3)
        Print #f, Mid$(Trim$(Component(IndexW(i)).Name), 1, 25); Tab(30); Format_It(Output_ECM_MASSBAL.MASSBAL_PERCENT_ERR(IndexW(i)), 3)
      Next i
      Print #f,
            
      
      Close (f)
    CMDialog1.Filename = ""
    Exit Sub

File_Error:
  If (Err.number = cdlCancel) Then
    'DO NOTHING.
  Else
    Call Show_Trapped_Error("cmdFile_Click")
  End If
  Resume Exit_Print_File
Exit_Print_File:
End Sub


Private Sub cmdPrint_Click()
Dim Error_Code As Integer, temp As String
Dim i As Integer, k As Integer, J As Integer

On Error GoTo Print_Error
'---Print other results-----------------------------------------------
  Printer.ScaleLeft = -1080  'Set a 3/4-inch margin
  Printer.ScaleTop = -1080
  Printer.CurrentX = 0
  Printer.CurrentY = 0

    Printer.FontSize = 12
    Printer.FontBold = True
    Printer.FontUnderline = True
    Printer.Print "Input data for the Equilibrium Colum Model"
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.FontUnderline = False
    '-- Print Filename
    Printer.Print
    Printer.Print "From Data File : " & Filename
    Printer.Print "Date/time stamp:" & Date$ & " " & Time$

    Printer.Print
    Printer.Print "Component"; Tab(30); "K*"; Tab(38); "1/n"; Tab(45); "Init. Conc."; Tab(59); "MW"
    Printer.Print Tab(39); "-"; Tab(48); "mg/L", Tab(58); "g/mol"
                                                                     
    For i = 1 To Number_Component_ECM
      'K = Component_Index_ECM(i)
      k = IndexW(i)
'      Printer.Print Trim$(Mid$(LTrim$(Component(K).Name), 1, 25)); Tab(29); Format$(Component(K).Use_K, "###,##0.000"); Tab(37); Format$(Component(K).Use_OneOverN, "0.000"); Tab(48); Format_It(Component(K).InitialConcentration, 2); Tab(58); Format$(Component(K).MW, "0.00")
      Printer.Print Trim$(Mid$(LTrim$(Component(k).Name), 1, 25)); Tab(29); Format$(Component(k).Use_K, "###,##0.000"); Tab(37); Format$(Component(k).Use_OneOverN, "0.000"); Tab(48); Format_It(Component(k).InitialConcentration, 2); Tab(58); Format$(Component(k).MW, "0.00")
    Next i
    Printer.Print
    Printer.Print "* K in (mg/g)*(L/mg)^(1/n)"

    Printer.Print

    '-----------------------Bed Data ----------------------
    Printer.FontUnderline = True
    Printer.Print "Bed Data:"
    Printer.FontUnderline = False

    Printer.Print Tab(5); "Bed Length: "; Tab(28); Format$(Bed.length, "0.000E+00") & " m"
    Printer.Print Tab(5); "Bed Diameter: "; Tab(28); Format$(Bed.Diameter, "0.000E+00") & " m"
    Printer.Print Tab(5); "Weight of GAC: "; Tab(28); Format$(Bed.Weight, "0.000E+00") & " kg"
    Printer.Print Tab(5); "Inlet Flowrate: "; Tab(28); Format$(Bed.Flowrate, "0.000E+00") & " m" & Chr$(179) & "/s"
    Printer.Print Tab(5); "EBCT: "; Tab(28); Format$(Bed.length * PI * Bed.Diameter * Bed.Diameter / 4# / Bed.Flowrate / 60#, "0.000E+00") & " mn"
    Printer.Print
    Printer.Print Tab(5); "Temperature:"; Tab(28); Format$(Bed.Temperature, "0.00") & " C"
    If Bed.Phase = 0 Then
      Printer.Print Tab(5); "Water Density:"; Tab(28); Format$(Bed.WaterDensity, "0.0000") & " g/cm" & Chr$(179)
      Printer.Print Tab(5); "Water Viscosity:"; Tab(28); Format$(Bed.WaterViscosity, "0.00E+00") & " g/cm.s"
    Else
      Printer.Print Tab(5); "Pressure:"; Tab(28); Format$(Bed.Pressure, "0.00000") & " atm"
      Printer.Print Tab(5); "Air Density:"; Tab(28); Format$(Bed.WaterDensity, "0.0000") & " g/cm" & Chr$(179)
      Printer.Print Tab(5); "Air Viscosity:"; Tab(28); Format$(Bed.WaterViscosity, "0.00E+00") & " g/cm.s"
    End If
    Printer.Print

    '-----------------Carbon Properties -------------------------------
    Printer.FontUnderline = True
    Printer.Print "Carbon Properties:"
    Printer.FontUnderline = False

    Printer.Print Tab(5); "Name: "; Tab(28); Trim$(Carbon.Name)
    Printer.Print Tab(5); "Apparent Density: "; Tab(28); Format$(Carbon.Density, "0.000") & " g/cm" & Chr$(179)
    Printer.Print Tab(5); "Particle Radius: "; Tab(28); Format$(Carbon.ParticleRadius * 100#, "0.000000") & " cm"
    Printer.Print Tab(5); "Porosity: "; Tab(28); Format$(Carbon.Porosity, "0.000")
    Printer.Print Tab(5); "Shape Factor: "; Tab(28); Format$(Carbon.ShapeFactor, "0.000")
    'Printer.Print Tab(5); "Tortuosity: "; Tab(28); Format$(Carbon.Tortuosity, "0.000")
    Printer.Print

                            
    Printer.Print
  '--- Print the results from the table
    Printer.FontSize = 12
    Printer.FontBold = True
    Printer.FontUnderline = True
    Printer.Print "Results for the Equilibrium Column Model"
    Printer.FontUnderline = False
    Printer.FontSize = 10
    Printer.FontBold = False

    Printer.Print
    Printer.Print "Zone"; Tab(9); "Component"; Tab(35); "BVF"; Tab(44); "Wave Vel."; Tab(54); "TC"; Tab(63); "Breakthrough"
    Printer.Print Tab(45); "cm/s"; Tab(54); "m3/kg"; Tab(63); Time_Unit
    For i = 1 To Number_Component_ECM
      Printer.Print "Zone " & Format$(i, "0"); Tab(9); Mid$(Trim$(Component(IndexW(i)).Name), 1, 25); Tab(35); Format_It(Output_ECM(i).Bed_Volume_Fed, 2); Tab(45); Format_It(Output_ECM(i).Wave_Velocity, 2); Tab(54); Format_It(1 / Output_ECM(i).Carbon_Usage_Rate * 1000, 2); Tab(63); Format_It(Time_Break(i), 2)
    
      'Change made: (ejo, 3/1/96)
      '==========================
      'was: Format_It(Output_ECM(i).Carbon_Usage_Rate, 2)
      'is now: Format_It(1 / Output_ECM(i).Carbon_Usage_Rate * 1000, 2)
    
    Next i
    Printer.Print
    Printer.Print "TC (Treatment Capacity) is in m" & Chr$(179) & "  / kg of GAC"
    Printer.Print

    For i = 1 To Number_Component_ECM
      Printer.FontBold = True
      Printer.Print Mid$(Trim$(Component(IndexW(i)).Name), 1, 25)
      Printer.FontBold = False
      Printer.Print "Zone "; Tab(9); "C/Co"; Tab(19); "C (mg/L)"; Tab(29); "Q (mg/L)"
      For J = 1 To Number_Component_ECM
       Printer.Print "Zone " & Format$(J, "0"); Tab(9); Format_It(CoCW(i, J), 2); Tab(19); Format_It(Liquid_ConcW(i, J) / 1000#, 2); Tab(29); Format_It(Solid_ConcW(i, J) / 1000#, 2)
      Next J
      Printer.Print
    Next i
    
    Printer.Print
  '--- Print the mass balance results
    Printer.FontSize = 12
    Printer.FontBold = True
    Printer.FontUnderline = True
    Printer.Print "Mass Balance Results"
    Printer.FontUnderline = False
    Printer.FontSize = 10
    Printer.FontBold = False

    Printer.Print
    'Printer.Print "Component"; Tab(30); "Left-Hand"; Tab(45); "Right-Hand"; Tab(60); "Percent Err."
    'Printer.Print ""; Tab(30); "(ug/cm2/s)"; Tab(45); "(ug/cm2/s)"; Tab(60); "(%)"
    Printer.Print "Component"; Tab(30); "Percent Err."
    Printer.Print ""; Tab(30); "(%)"
    For i = 1 To Number_Component_ECM
'      Printer.Print Mid$(Trim$(Component(i).Name), 1, 25); Tab(30); Format_It(Output_ECM_MASSBAL.MASSBAL_C0_e_Vf(i), 3); Tab(45); Format_It(Output_ECM_MASSBAL.MASSBAL_TERM_SUM(i), 3); Tab(60); Format_It(Output_ECM_MASSBAL.MASSBAL_PERCENT_ERR(i), 3)
'      Printer.Print Mid$(Trim$(Component(IndexW(i)).Name), 1, 25); Tab(30); Format_It(Output_ECM_MASSBAL.MASSBAL_C0_e_Vf(IndexW(i)), 3); Tab(45); Format_It(Output_ECM_MASSBAL.MASSBAL_TERM_SUM(IndexW(i)), 3); Tab(60); Format_It(Output_ECM_MASSBAL.MASSBAL_PERCENT_ERR(IndexW(i)), 3)
      Printer.Print Mid$(Trim$(Component(IndexW(i)).Name), 1, 25); Tab(30); Format_It(Output_ECM_MASSBAL.MASSBAL_PERCENT_ERR(IndexW(i)), 3)
    Next i
    Printer.Print


    Printer.EndDoc
    Exit Sub
Print_Error:
  Call Show_Trapped_Error("cmdPrint_Click")
  Resume Exit_Print
Exit_Print:

End Sub


Private Sub cmdSelect_Click()
Dim Error_Code As Integer
Dim temp As String
  On Error GoTo Select_Print_Error
  'CMDialog1.flags = PD_PRINTSETUP
  'CMDialog1.Action = 5
  CMDialog1.CancelError = False
  CMDialog1.ShowPrinter
  Exit Sub
Select_Print_Error:
  Call Show_Trapped_Error("cmdSelect_Click")
  Resume Exit_Select_Print
Exit_Select_Print:
End Sub


Private Sub Draw(GFlag As Integer)
Dim Num_Compo As Integer
  Num_Compo = Number_Component_ECM
Dim i As Integer, J As Integer, k As Integer
   Select Case GFlag
     Case 1
    grpGlob.NumSets = NumW
    grpGlob.GraphType = 4
    grpGlob.GraphStyle = 0
    
    For J = 1 To grpGlob.NumSets
     grpGlob.ThisSet = J
     If Num_Compo >= 2 Then
      grpGlob.NumPoints = NumW
     Else
      grpGlob.NumPoints = 2
     End If
    Next J
    grpGlob.GraphTitle = " C/Co for All Components"
    grpGlob.GridStyle = 3

    For J = 1 To grpGlob.NumSets
      grpGlob.ThisSet = J
       'K = IndexW(NumW - J + 1)
       k = J
       grpGlob.ThisPoint = J
       grpGlob.LegendText = Name_CompW(J)
      For i = 1 To grpGlob.NumPoints
        grpGlob.ThisPoint = i
        grpGlob.LabelText = "Zone " & Format$(i, "0")
        grpGlob.ThisPoint = i
        grpGlob.GraphData = CoCW(k, i)
      Next i
    Next J
   Case 2
    grpGlob.NumSets = NumW
    grpGlob.GraphType = 3
    grpGlob.GraphStyle = 0
    grpGlob.GraphTitle = " Q (" & Chr$(181) & "g/g) for All Components"
    grpGlob.GridStyle = 3
    
    For J = 1 To grpGlob.NumSets
     grpGlob.ThisSet = J
     If NumW > 2 Then
      grpGlob.NumPoints = NumW
     Else
      grpGlob.NumPoints = 2
     End If
    Next J

    For J = 1 To grpGlob.NumSets
      grpGlob.ThisSet = J
       'K = IndexW(NumW - J + 1)
       k = J
       grpGlob.ThisPoint = J
       grpGlob.LegendText = Name_CompW(k)
      For i = 1 To grpGlob.NumPoints
        grpGlob.ThisPoint = i
        grpGlob.LabelText = "Zone " & Format$(i, "0")
        grpGlob.ThisPoint = i
        grpGlob.GraphData = Solid_ConcW(k, i)
      Next i
    Next J

     Case 3
    grpGlob.NumSets = NumW
    grpGlob.GraphType = 3
    grpGlob.GraphStyle = 0
    grpGlob.GraphTitle = " C (" & Chr$(181) & "g/L) for All Components"
    grpGlob.GridStyle = 3
    
    For J = 1 To grpGlob.NumSets
     grpGlob.ThisSet = J
     If NumW > 2 Then
      grpGlob.NumPoints = NumW
     Else
      grpGlob.NumPoints = 2
     End If
    Next J

    For J = 1 To grpGlob.NumSets
      grpGlob.ThisSet = J
       'K = IndexW(NumW - J + 1)
        k = J
       grpGlob.ThisPoint = J
       grpGlob.LegendText = Name_CompW(k)
      For i = 1 To grpGlob.NumPoints
        grpGlob.ThisPoint = i
        grpGlob.LabelText = "Zone " & Format$(i, "0")
        grpGlob.ThisPoint = i
        grpGlob.GraphData = Liquid_ConcW(k, i)
      Next i
    Next J
   End Select
   If Number_Component_ECM = 1 Then
    grpGlob.ThisPoint = 2
    grpGlob.LabelText = ""
   End If
End Sub

Private Sub Form_Load()
Dim i As Integer, J As Integer
   
   'Move frmPFPSDM.Left + (frmPFPSDM.Width / 2) - (frmGlobal.Width / 2), frmPFPSDM.Top + (frmPFPSDM.Height / 2) - (frmGlobal.Height / 2)
   Call CenterOnForm(Me, frmMain)
    
    Label1 = "VTM" & Chr$(13) & "(m" & Chr$(179) & "/kg)"
    'Me.HelpContextID = Hlp_Global_Results
    ''''Caption = "Results for the Equilibrium Column Model"
    NumW = Number_Component_ECM
    For i = 1 To NumW
        IndexW(i) = Output_ECM(i).Index
        Name_CompW(i) = Component(IndexW(i)).Name
      For J = 1 To NumW
       Solid_ConcW(i, J) = Output_ECM(i).Solid_Concentration(J)
       Liquid_ConcW(i, J) = Output_ECM(i).Liquid_Concentration(J)
       CoCW(i, J) = Output_ECM(i).C_Over_C0(J)
      Next J
    Next i
    
    ''''fraGlob = "Results"
   lblZone = ""
    lblData1 = ""
    lblData2 = ""
    lblData3 = ""
    lblData4 = ""
    lblData5 = ""
    lblCompo = ""
    ReDim Time_Break(NumW)

    'Time_Min = 1E+100
    For i = 1 To NumW
      Time_Break(i) = Bed.length * 100# / Output_ECM(i).Wave_Velocity / 3600# / 24#
    '  If Time_Break(I) < Time_Min Then
    '   Time_Min = Time_Break(I)
    '  End If
    'Next I
    'If Time_Min < 1# Then
    ' For I = 1 To NumW
    '   Time_Break(I) = Time_Break(I) * 24# * 60#
     Next i
    ' Time_Unit = "mn"
    'Else Time_Unit = "days"
    'End If
    Time_Unit = "days"

    Label6 = "Breakthrough time(" & Time_Unit & ")"
    For i = 1 To NumW
      lblZone = lblZone & Format$(i, "0") & Chr$(10)
      lblCompo = lblCompo & LCase$(Trim$(Component(IndexW(i)).Name) & Chr$(10))
      lblData1 = lblData1 & Format_It(Output_ECM(i).Bed_Volume_Fed, 2) & Chr$(10)
      lblData2 = lblData2 & Format_It(Output_ECM(i).Wave_Velocity, 2) & Chr$(10)
      lblData3 = lblData3 & Format_It(1 / Output_ECM(i).Carbon_Usage_Rate * 1000, 2) & Chr$(10)
      lblData4 = lblData4 & Format_It(Time_Break(i), 3) & Chr$(10)
      lblData5 = lblData5 & Format_It(Output_ECM_MASSBAL.MASSBAL_PERCENT_ERR(i), 2) & Chr$(10)
    Next i
    
    Call Populate_Scrollboxes
    Call cboGlob_Click
    
    
    'Call Draw(1)
    'grpGlob.DrawMode = 2


End Sub

Private Sub Form_Unload(Cancel As Integer)

  Call UserPrefs_Save

End Sub


Private Sub Populate_Scrollboxes()

  PopulatingScrollboxes = True
    
  cboGlob.Clear
  cboGlob.AddItem "C/Co"
  cboGlob.AddItem "Q"
  cboGlob.AddItem "C (Liquid Conc.)"

  '-- Read in INI settings
  cboGlob.ListIndex = 0
  Call UserPrefs_Load
    
  PopulatingScrollboxes = False

End Sub

Private Sub UserPrefs_Load()
Dim X As Long

  On Error GoTo err_FRMGLOBAL_UserPrefs_Load

  X = CLng(INI_Getsetting("FRMGLOBAL_cboGlob"))
  If ((X >= 0) And (X <= cboGlob.ListCount - 1)) Then
    cboGlob.ListIndex = X
  End If
  Exit Sub

resume_err_FRMGLOBAL_UserPrefs_Load:
  Call UserPrefs_Save
  Exit Sub

err_FRMGLOBAL_UserPrefs_Load:
  Resume resume_err_FRMGLOBAL_UserPrefs_Load
           
End Sub

Private Sub UserPrefs_Save()
Dim X As Long

  X = cboGlob.ListIndex
  Call INI_PutSetting("FRMGLOBAL_cboGlob", Trim$(CStr(X)))

End Sub


