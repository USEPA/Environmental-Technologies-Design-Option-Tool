VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "Comdlg32.ocx"
Begin VB.Form frmPrintInputs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Model Inputs"
   ClientHeight    =   3285
   ClientLeft      =   3615
   ClientTop       =   5820
   ClientWidth     =   4395
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   4395
   Begin Threed.SSFrame SSFrame1 
      Height          =   2445
      Left            =   60
      TabIndex        =   0
      Top             =   150
      Width           =   4245
      _Version        =   65536
      _ExtentX        =   7488
      _ExtentY        =   4313
      _StockProps     =   14
      Caption         =   "Print:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSCheck chkSelect 
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   1
         Top             =   2070
         Width           =   3000
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "      Effluent Data"
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
      End
      Begin Threed.SSCheck chkSelect 
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   2
         Top             =   1770
         Width           =   3000
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "      Variable Influent Data"
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
      End
      Begin Threed.SSCheck chkSelect 
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   3
         Top             =   1470
         Width           =   3000
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "      Fouling Correlations"
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
      End
      Begin Threed.SSCheck chkSelect 
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   570
         Width           =   3000
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "      Bed Data"
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
      End
      Begin Threed.SSCheck chkSelect 
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   870
         Width           =   3000
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "      Carbon Properties"
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
      End
      Begin Threed.SSCheck chkSelect 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   270
         Width           =   3000
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "      Component Properties"
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
      End
      Begin Threed.SSCheck chkSelect 
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   1170
         Width           =   3000
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "      Kinetic Parameters"
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
      End
   End
   Begin Threed.SSCommand cmdPrint 
      Height          =   435
      Left            =   2850
      TabIndex        =   8
      Top             =   2730
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
   Begin Threed.SSCommand cmdCancel 
      Height          =   435
      Left            =   60
      TabIndex        =   9
      Top             =   2730
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "&Cancel"
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
      Left            =   1740
      Top             =   2700
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
      FontSize        =   0
      MaxFileSize     =   256
   End
End
Attribute VB_Name = "frmPrintInputs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Dim Filename_Input As String





Const frmPrintInputs_declarations_end = True


Private Sub chkSelect_Click(index As Integer, Value As Integer)
  chkSelect(index).Tag = CStr(Str$(Value))
End Sub


Private Sub cmdCancel_Click()
  Unload Me
End Sub


Private Sub cmdPrint_Click()
Dim Error_Code As Integer, temp As String, f  As Integer
Dim i As Integer, Dummy As Double, DFlag As Integer
Dim Eq1 As String, temporaryname As String, response As Integer
Dim s As String
Dim J As Integer

  DFlag = False
  For i = 0 To 4
    If chkSelect(i).Value Then DFlag = True
  Next i
  If Not (DFlag) Then
   Call Show_Error("You must select something to print!")
   Exit Sub
  End If

  If Print_To_Printer Then
On Error GoTo Print_Error
    Printer.ScaleLeft = -1080  'Set a 3/4-inch margin
    Printer.ScaleTop = -1080
    Printer.CurrentX = 0
    Printer.CurrentY = 0
    Printer.FontSize = 10
    Printer.Print Filename
    Printer.Print
    '---- Print Component Properties  ---------
    If chkSelect(0).Value Then
      Printer.FontSize = 12
      Printer.FontBold = True
      Printer.Print Tab(25); "Component Properties"
      Printer.FontSize = 10
      Printer.FontBold = False
      Printer.Print
      Printer.Print "Component"; Tab(30); "K*"; Tab(38); "1/n"; Tab(47); "C0"; Tab(57); "MW"; Tab(65); "Vm"; Tab(75); "NBP"
      Printer.Print Tab(39); "-"; Tab(46); "mg/L"; Tab(56); "g/mol"; Tab(65); "cm" & Chr$(179) & "/mol"; Tab(76); "C"
      For i = 1 To Number_Component
        Printer.Print Trim$(Mid$(LTrim$(Component(i).Name), 1, 25)); Tab(29); Format$(Component(i).Use_K, "###,##0.000"); Tab(37); Format$(Component(i).Use_OneOverN, "0.000"); Tab(46); Format_It(Component(i).InitialConcentration, 2); Tab(55); Format_It(Component(i).MW, 2); Tab(64); Format_It(Component(i).MolarVolume, 2); Tab(73); Format_It(Component(i).BP, 2)
      Next i
      Printer.Print
      Printer.Print "* K in (mg/g)*(L/mg)^(1/n) - Vm = Molar Volume at NBP"
      Printer.Print
    End If
    '---Print bed data-----
    Call GetMoreBedParameters
    If chkSelect(1).Value Then
      Printer.FontSize = 12
      Printer.FontBold = True
      Printer.Print Tab(25); "Fixed-Bed Properties"
      Printer.FontSize = 10
      Printer.FontBold = False
      Printer.Print
      Printer.Print "Bed Length:"; Tab(25); Format$(Bed.length, "0.000E+00") & " m"
      Printer.Print "Bed Diameter:"; Tab(25); Format$(Bed.Diameter, "0.000E+00") & " m"
      Printer.Print "Weight of GAC:"; Tab(25); Format$(Bed.Weight, "0.000E+00") & " kg"
      Printer.Print "Inlet Flowrate:"; Tab(25); Format$(Bed.Flowrate, "0.000E+00") & " m" & Chr$(179) & "/s"
      Printer.Print "EBCT:"; Tab(25); Format$(Bed.length * PI * Bed.Diameter * Bed.Diameter / 4# / Bed.Flowrate / 60#, "0.000E+00") & " mn"
      Printer.Print "Bed Density:"; Tab(25); Format$(Bed.Density, "0.000") & " g/cm3"
      Printer.Print "Bed Porosity:"; Tab(25); Format$(Bed.Porosity, "0.000")
      Printer.Print "Superficial Velocity:"; Tab(25); Format$(Bed.SuperficialVelocity * 3600#, "0.00E+00") & " m/hr"
      Printer.Print "Interstitial Velocity:"; Tab(25); Format$(Bed.InterstitialVelocity * 3600#, "0.00E+00") & " m/hr"
      Printer.Print
      Printer.Print "Temperature:"; Tab(25); Format$(Bed.Temperature, "0.00") & " C"
      Printer.Print "Water Density:"; Tab(25); Format$(Bed.WaterDensity, "0.0000") & " g/cm" & Chr$(179)
      Printer.Print "Water Viscosity:"; Tab(25); Format$(Bed.WaterViscosity, "0.00E+0") & " g/cm.s"
      Printer.Print
    End If
    '--- Print Carbon Properties -----
    If chkSelect(2).Value Then
      Printer.FontSize = 12
      Printer.FontBold = True
      Printer.Print Tab(25); "Carbon Properties"
      Printer.FontSize = 10
      Printer.FontBold = False
      Printer.Print
      Printer.Print "Name:"; Tab(19); Trim$(Carbon.Name)
      Printer.Print "Apparent Density:"; Tab(19); Format$(Carbon.Density, "0.000") & " g/cm" & Chr$(179)
      Printer.Print "Particle Radius:"; Tab(19); Format$(Carbon.ParticleRadius * 100#, "0.00000") & " cm"
      Printer.Print "Porosity:"; Tab(19); Format$(Carbon.Porosity, "0.000")
      Printer.Print "Shape Factor: "; Tab(19); Format$(Carbon.ShapeFactor, "0.000")
      'Printer.Print "Tortuosity:"; Tab(19); Format$(Carbon.Tortuosity, "0.000")
      Printer.Print
    End If
    '--- Print kinetic Parameters -----
    If chkSelect(3).Value Then
      Printer.FontSize = 12
      Printer.FontBold = True
      Printer.Print Tab(25); "Kinetic Parameters"
      Printer.FontSize = 10
      Printer.FontBold = False
      Printer.Print
      Printer.Print "Component"; Tab(24); "kf"; Tab(33); "Ds"; Tab(42); "Dp"; Tab(50); "St"; Tab(58); "Eds"; Tab(67); "Edp"; Tab(75); "SPDFR"
      Printer.Print Tab(23); "cm/s"; Tab(32); "cm" & Chr$(178) & "/s"; Tab(41); "cm" & Chr$(178) & "/s"; Tab(50); "-"; Tab(59); "-"; Tab(68); "-"; Tab(77); "-"
      For i = 1 To Number_Component
          Printer.Print Mid$(Trim$(Component(i).Name), 1, 20); Tab(22); Format_It(Component(i).kf, 2); Tab(31); Format$(Component(i).Ds, "0.00E+00"); Tab(40); Format$(Component(i).Dp, "0.00E+00"); Tab(49); Format_It(ST(i), 2); Tab(58); Format_It(Eds(i), 2); Tab(67); Format_It(Edp(i), 2); Tab(76); Format_It(Component(i).SPDFR, 2)
      Next i
    End If
    '--- Print fouling correlations ---
    If chkSelect(4).Value Then
        J = False
        
        For i = 1 To Number_Component
          'if and only if using correlation then print correlation
          If Component(i).Use_Tortuosity_Correlation = True Then J = True
        Next i
      
      If J Then
        Printer.Print
        Printer.FontSize = 12
        Printer.FontBold = True
        Printer.Print Tab(25); "Fouling Correlations"
        Printer.FontSize = 10
        Printer.FontBold = False
        Printer.Print
  
        Printer.Print " Water type : "; Trim$(Bed.Water_Correlation.Name)
        Eq1 = Format$(Bed.Water_Correlation.Coeff(1), "0.00")
  
        If Bed.Water_Correlation.Coeff(2) > 0 Then
          Eq1 = Eq1 & " + " & Format$(Bed.Water_Correlation.Coeff(2), "0.00E+00") & "* t "
        Else
        If Bed.Water_Correlation.Coeff(2) < 0 Then
          Eq1 = Eq1 & " - " & Format$(Abs(Bed.Water_Correlation.Coeff(2)), "0.00E+00") & "* t "
         End If
        End If
        If Bed.Water_Correlation.Coeff(3) > 0 Then
         Eq1 = Eq1 & " + " & Format$(Bed.Water_Correlation.Coeff(3), "0.00") & "* EXP("
        Else
         If Bed.Water_Correlation.Coeff(3) < 0 Then
          Eq1 = Eq1 & " - " & Format$(Abs(Bed.Water_Correlation.Coeff(3)), "0.00") & "* EXP("
         End If
        End If
        If Bed.Water_Correlation.Coeff(3) <> 0 Then
          If Bed.Water_Correlation.Coeff(4) > 0 Then
           Eq1 = Eq1 & Format$(Bed.Water_Correlation.Coeff(4), "0.00E+00") & "* t)"
          Else
           If Bed.Water_Correlation.Coeff(4) < 0 Then
            Eq1 = Eq1 & Format$(Bed.Water_Correlation.Coeff(4), "0.00E+00") & "* t)"
           End If
          End If
        End If
        Printer.Print "K(t)/K0 = " & Eq1
        Printer.Print "(t in minutes)"
        Printer.Print
  
        For i = 1 To Number_Component
         If Component(i).Use_Tortuosity_Correlation = True Then
            Eq1 = ""
            If Component(i).Correlation.Coeff(1) = 1# Then
            Eq1 = "(K/K0) "
           Else
            If Component(i).Correlation.Coeff(1) <> 0 Then Eq1 = Format$(Component(i).Correlation.Coeff(1), "0.00") & " * (K/K0) "
           End If
           If Component(i).Correlation.Coeff(2) > 0 Then
            Eq1 = Eq1 & "+ " & Format$(Component(i).Correlation.Coeff(2), "0.00")
           Else
            If Component(i).Correlation.Coeff(2) <> 0 Then Eq1 = Eq1 & "- " & Format$(Abs(Component(i).Correlation.Coeff(2)), "0.00")
           End If
           If Trim$(Eq1) = "" Then
             Eq1 = "K/K0"
           End If
           Printer.Print Trim$(Component(i).Name) & ":"
           Printer.Print Tab(10); "Correlation type: " & Trim$(Component(i).Correlation.Name)
    
           Printer.Print Tab(10); "K/K0 = " & Eq1
           If (Component(i).Use_Tortuosity_Correlation) Then
             If (Component(i).Constant_Tortuosity) Then
               Printer.Print "Correlation used when SOC competition is important:"
               Printer.Print " Tortuosity = 0.782 * EBCT^0.925 "
             Else
               Printer.Print "Correlation used when NOM fouling is important:"
               Printer.Print " Tortuosity = 1.0 if t< 70 days"
               Printer.Print " Tortuosity = 0.334 + 6.610E-06 * t   (t in minutes)"
             End If
           End If
           
           Printer.Print
         End If
        Next i
  
        Printer.Print
      End If
     End If

      '--- Print Variable Influent Data
      If chkSelect(5).Value Then
        Printer.Print
        Printer.Print Tab(25); "Variable Influent Data"
        Printer.Print
        s = "Time(days)"
        For J = 1 To Number_Component
          s = s & ",C of " & Trim$(Component(J).Name)
        Next J
        s = s & ":"
        Printer.Print s
        Printer.Print "(All C in mg/L)"
        For i = 1 To Number_Influent_Points
          s = Trim$(Str$(T_Influent(i) / 60# / 24#))     'Convert min--->days
          For J = 1 To Number_Component
            s = s & "," & Trim$(Str$(C_Influent(J, i)))
          Next J
          Printer.Print s
        Next i
      End If
      
      '--- Print Effluent Data
      If chkSelect(6).Value Then
        Printer.Print
        Printer.Print Tab(25); "Effluent Data"
        Printer.Print
        s = "Time(days)"
        For J = 1 To Number_Component
          s = s & ",C/C0 of " & Trim$(Component(J).Name)
        Next J
        s = s & ":"
        Printer.Print s
        Printer.Print "(All C/C0 are dimensionless and normalized)"
        For i = 1 To NData_Points
          s = Trim$(Str$(T_Data_Points(i)))
          For J = 1 To Number_Component
            s = s & "," & Trim$(Str$(C_Data_Points(J, i)))
          Next J
          Printer.Print s
        Next i
      End If

    Printer.EndDoc
  Else

On Error GoTo File_Error
    CMDialog1.CancelError = True
    CMDialog1.Filename = ""
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
      Print #f, Filename

      '---- Print Component Properties  ---------
      If chkSelect(0).Value Then
       Print #f, Tab(25); "Component Properties"
       Print #f,
       Print #f, "Component"; Tab(30); "K*"; Tab(38); "1/n"; Tab(47); "C0"; Tab(57); "MW"; Tab(65); "Vm"; Tab(75); "NBP"
       Print #f, Tab(39); "-"; Tab(46); "mg/L"; Tab(56); "g/mol"; Tab(65); "cm" & Chr$(179) & "/mol"; Tab(76); "C"
        For i = 1 To Number_Component
         Print #f, Trim$(Mid$(LTrim$(Component(i).Name), 1, 25)); Tab(29); Format$(Component(i).Use_K, "###,##0.000"); Tab(37); Format$(Component(i).Use_OneOverN, "0.000"); Tab(46); Format_It(Component(i).InitialConcentration, 2); Tab(55); Format_It(Component(i).MW, 2); Tab(64); Format_It(Component(i).MolarVolume, 2); Tab(73); Format_It(Component(i).BP, 2)
        Next i
       Print #f,
       Print #f, "* K in (mg/g)*(L/mg)^(1/n) - Vm = Molar Volume at NBP"
       Print #f,
      End If

      '---Print bed data-----
      Call GetMoreBedParameters
      If chkSelect(1).Value Then
        Print #f, Tab(25); "Bed Data"
        Print #f,
        Print #f, "Bed Length:"; Tab(18); Format$(Bed.length, "0.000E+00") & " m"
        Print #f, "Bed Diameter:"; Tab(18); Format$(Bed.Diameter, "0.000E+00") & " m"
        Print #f, "Weight of GAC:"; Tab(18); Format$(Bed.Weight, "0.000E+00") & " kg"
        Print #f, "Inlet Flowrate:"; Tab(18); Format$(Bed.Flowrate, "0.000E+00") & " m" & Chr$(179) & "/s"
        Print #f, "EBCT:"; Tab(18); Format$(Bed.length * PI * Bed.Diameter * Bed.Diameter / 4# / Bed.Flowrate / 60#, "0.000E+00") & " mn"
        Print #f, "Bed Density:"; Tab(25); Format$(Bed.Density, "0.000") & " g/cm3"
        Print #f, "Bed Porosity:"; Tab(25); Format$(Bed.Porosity, "0.000")
        Print #f, "Superficial Velocity:"; Tab(25); Format$(Bed.SuperficialVelocity * 3600#, "0.00E+00") & " m/hr"
        Print #f, "Interstitial Velocity:"; Tab(25); Format$(Bed.InterstitialVelocity * 3600#, "0.00E+00") & " m/hr"
        Print #f,
        Print #f, "Temperature:"; Tab(18); Format$(Bed.Temperature, "0.00") & " C"
        Print #f, "Water Density:"; Tab(18); Format$(Bed.WaterDensity, "0.0000") & " g/cm" & Chr$(179)
        Print #f, "Water Viscosity:"; Tab(18); Format$(Bed.WaterViscosity, "0.00E+0") & " g/cm.s"
        Print #f,
      End If
      '--- Print Carbon Properties -----
      If chkSelect(2).Value Then
        Print #f, Tab(25); "Carbon Properties"
        Print #f,
        Print #f, "Name:"; Tab(19); Trim$(Carbon.Name)
        Print #f, "Apparent Density:"; Tab(19); Format$(Carbon.Density, "0.000") & " g/cm" & Chr$(179)
        Print #f, "Particle Radius:"; Tab(19); Format$(Carbon.ParticleRadius * 100#, "0.00000") & " cm"
        Print #f, "Porosity:"; Tab(19); Format$(Carbon.Porosity, "0.000")
        Print #f, "Shape Factor: "; Tab(19); Format$(Carbon.ShapeFactor, "0.000")
        'Print #f, "Tortuosity:"; Tab(19); Format$(Carbon.Tortuosity, "0.000")
        Print #f,
      End If
      '--- Print kinetic Parameters -----
      If chkSelect(3).Value Then
       Print #f, Tab(25); "Kinetic Parameters"
       Print #f,
       Print #f, "Component"; Tab(24); "kf"; Tab(33); "Ds"; Tab(42); "Dp"; Tab(50); "St"; Tab(58); "Eds"; Tab(67); "Edp"; Tab(75); "SPDFR"
       Print #f, Tab(23); "cm/s"; Tab(32); "cm" & Chr$(178) & "/s"; Tab(41); "cm" & Chr$(178) & "/s"; Tab(50); "-"; Tab(59); "-"; Tab(68); "-"; Tab(77); "-"
        For i = 1 To Number_Component
         Print #f, Mid$(Trim$(Component(i).Name), 1, 20); Tab(22); Format_It(Component(i).kf, 2); Tab(31); Format$(Component(i).Ds, "0.00E+00"); Tab(40); Format$(Component(i).Dp, "0.00E+00"); Tab(49); Format_It(ST(i), 2); Tab(58); Format_It(Eds(i), 2); Tab(67); Format_It(Edp(i), 2); Tab(76); Format_It(Component(i).SPDFR, 2)
        Next i
      End If
    '--- Print fouling correlations ---
      If chkSelect(4).Value Then
        'check to see if fouling needed
        J = False
        For i = 1 To Number_Component
          'if and only if using correlation then print correlation
          If Component(i).Use_Tortuosity_Correlation = True Then J = True
        Next i

       If J Then
         Print #f,
         Print #f, Tab(25); "Fouling Correlations"
         Print #f,
  
         Print #f, " Water type : "; Trim$(Bed.Water_Correlation.Name)
          Eq1 = Format$(Bed.Water_Correlation.Coeff(1), "0.00")
  
          If Bed.Water_Correlation.Coeff(2) > 0 Then
            Eq1 = Eq1 & " + " & Format$(Bed.Water_Correlation.Coeff(2), "0.00E+00") & "* t "
          Else
          If Bed.Water_Correlation.Coeff(2) < 0 Then
            Eq1 = Eq1 & " - " & Format$(Abs(Bed.Water_Correlation.Coeff(2)), "0.00E+00") & "* t "
           End If
          End If
          If Bed.Water_Correlation.Coeff(3) > 0 Then
           Eq1 = Eq1 & " + " & Format$(Bed.Water_Correlation.Coeff(3), "0.00") & "* EXP("
          Else
           If Bed.Water_Correlation.Coeff(3) < 0 Then
            Eq1 = Eq1 & " - " & Format$(Abs(Bed.Water_Correlation.Coeff(3)), "0.00") & "* EXP("
           End If
          End If
          If Bed.Water_Correlation.Coeff(3) <> 0 Then
            If Bed.Water_Correlation.Coeff(4) > 0 Then
             Eq1 = Eq1 & Format$(Bed.Water_Correlation.Coeff(4), "0.00E+00") & "* t)"
            Else
             If Bed.Water_Correlation.Coeff(4) < 0 Then
              Eq1 = Eq1 & Format$(Bed.Water_Correlation.Coeff(4), "0.00E+00") & "* t)"
             End If
            End If
          End If
         Print #f, "K(t)/K0 = " & Eq1
         Print #f, "(t in minutes)"
         Print #f,
  
          For i = 1 To Number_Component
            'if and only if using correlation then print correlation
            If Component(i).Use_Tortuosity_Correlation = True Then
           
               Eq1 = ""
               If Component(i).Correlation.Coeff(1) = 1# Then
                Eq1 = "(K/K0) "
               Else
                If Component(i).Correlation.Coeff(1) <> 0 Then Eq1 = Format$(Component(i).Correlation.Coeff(1), "0.00") & " * (K/K0) "
               End If
               If Component(i).Correlation.Coeff(2) > 0 Then
                Eq1 = Eq1 & "+ " & Format$(Component(i).Correlation.Coeff(2), "0.00")
               Else
                If Component(i).Correlation.Coeff(2) <> 0 Then Eq1 = Eq1 & "- " & Format$(Abs(Component(i).Correlation.Coeff(2)), "0.00")
               End If
               If Trim$(Eq1) = "" Then
                 Eq1 = "K/K0"
               End If
              Print #f, Trim$(Component(i).Name) & ":"
              Print #f, Tab(10); "Correlation type: " & Trim$(Component(i).Correlation.Name)
      
              Print #f, Tab(10); "K/K0 = " & Eq1
              
              If (Component(i).Use_Tortuosity_Correlation) Then
                If (Component(i).Constant_Tortuosity) Then
                  Print #f, "Correlation used when SOC competition is important:"
                  Print #f, " Tortuosity = 0.782 * EBCT^0.925 "
                Else
                  Print #f, "Correlation used when NOM fouling is important:"
                  Print #f, " Tortuosity = 1.0 if t< 70 days"
                  Print #f, " Tortuosity = 0.334 + 6.610E-06 * t   (t in minutes)"
                End If
              End If
              Print #f,
            End If
          Next i
        End If
        'If Use_Tortuosity_Correlation Then
        '  If Constant_Tortuosity Then
        '   Print #f, "Correlation used when SOC competition is important:"
        '   Print #f, " Tortuosity = 0.782 * EBCT^0.925 "
        '  Else
        '   Print #f, "Correlation used when NOM fouling is important:"
        '   Print #f, " Tortuosity = 1.0 if t< 70 days"
        '   Print #f, " Tortuosity = 0.334 + 6.610E-06 * EBCT"
        '  End If
        'End If
       Print #f,
      End If
      
      '--- Print Variable Influent Data

      If chkSelect(5).Value Then
        Print #f,
        Print #f, Tab(25); "Variable Influent Data"
        Print #f,
        s = "Time(days)"
        For J = 1 To Number_Component
          s = s & ",C of " & Trim$(Component(J).Name)
        Next J
        s = s & ":"
        Print #f, s
        Print #f, "(All C in mg/L)"
        For i = 1 To Number_Influent_Points
          s = Trim$(Str$(T_Influent(i) / 60# / 24#))     'Convert min--->days
          For J = 1 To Number_Component
            s = s & "," & Trim$(Str$(C_Influent(J, i)))
          Next J
          Print #f, s
        Next i
      End If
      
      '--- Print Effluent Data
      If chkSelect(6).Value Then
        Print #f,
        Print #f, Tab(25); "Effluent Data"
        Print #f,
        s = "Time(days)"
        For J = 1 To Number_Component
          s = s & ",C/C0 of " & Trim$(Component(J).Name)
        Next J
        s = s & ":"
        Print #f, s
        Print #f, "(All C/C0 are dimensionless and normalized)"
        For i = 1 To NData_Points
          s = Trim$(Str$(T_Data_Points(i)))
          For J = 1 To Number_Component
            s = s & "," & Trim$(Str$(C_Data_Points(J, i)))
          Next J
          Print #f, s
        Next i
      End If

      Close (f)
    End If
  CMDialog1.Filename = ""
  Unload Me
  Exit Sub

Print_Error:
  If (Err.number = cdlCancel) Then
    'DO NOTHING.
  Else
    Call Show_Trapped_Error("cmdPrint_Click")
  End If
  Resume Exit_Print
File_Error:
  If (Err.number = cdlCancel) Then
    'DO NOTHING.
  Else
    Call Show_Trapped_Error("cmdPrint_Click")
  End If
  Resume Exit_Print
Exit_Print:
End Sub


Private Sub Form_Load()
Dim temp As String
Dim temp2 As String
Dim temp3 As String
Dim i As Integer
  ''''Me.HelpContextID = Hlp_Print_
  Call UserPrefs_Load
  For i = 0 To 6
    chkSelect(i).Tag = CStr(Str$(chkSelect(i).Value))
  Next i
  Call CenterOnForm(Me, frmMain)
  If Print_To_Printer Then
    Me.Caption = "Print to printer"
  Else
    Me.Caption = "Print to file"
  End If
  If (Number_Component <= 0) Then
    temp = chkSelect(0).Tag
    temp2 = chkSelect(3).Tag
    temp3 = chkSelect(4).Tag
    chkSelect(0).Enabled = False
    chkSelect(0).Value = False
    chkSelect(3).Enabled = False
    chkSelect(3).Value = False
    chkSelect(4).Enabled = False
    chkSelect(4).Value = False
    chkSelect(0).Tag = temp
    chkSelect(3).Tag = temp2
    chkSelect(4).Tag = temp3
  Else
    chkSelect(0).Value = True
    chkSelect(3).Value = True
    chkSelect(4).Value = True
  End If
  If (Number_Influent_Points = 0) Then
    temp = chkSelect(5).Tag
    chkSelect(5).Value = False
    chkSelect(5).Enabled = False
    chkSelect(5).Tag = temp
  End If
  If (NData_Points = 0) Then
    temp = chkSelect(6).Tag
    chkSelect(6).Value = False
    chkSelect(6).Enabled = False
    chkSelect(6).Tag = temp
  End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Call UserPrefs_Save
End Sub


Private Sub UserPrefs_Load()
Dim X As Long
Dim i As Integer
Dim varname As String
  On Error GoTo err_FRMPRINT_UserPrefs_Load
  For i = 0 To 6
    varname = "FRMPRINT_chkSelect(" & Trim$(Str$(i)) & ")"
    X = CLng(INI_Getsetting(varname))
    chkSelect(i) = X
  Next i
  Exit Sub
resume_err_FRMPRINT_UserPrefs_Load:
  Call UserPrefs_Save
  Exit Sub
err_FRMPRINT_UserPrefs_Load:
  Resume resume_err_FRMPRINT_UserPrefs_Load
End Sub
Private Sub UserPrefs_Save()
Dim X As Long
Dim i As Integer
Dim varname As String
  For i = 0 To 6
    varname = "FRMPRINT_chkSelect(" & Trim$(Str$(i)) & ")"
    X = CLng(chkSelect(i).Value)
    Call INI_PutSetting(varname, Trim$(CStr(X)))
  Next i
End Sub


