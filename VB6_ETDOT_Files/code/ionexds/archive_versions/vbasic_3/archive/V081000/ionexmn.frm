VERSION 2.00
Begin Form frmIonExchangeMain 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Ion Exchange Simulation Software"
   ClientHeight    =   6630
   ClientLeft      =   825
   ClientTop       =   1575
   ClientWidth     =   9510
   Height          =   7320
   Icon            =   IONEXMN.FRX:0000
   Left            =   765
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   9510
   Top             =   945
   Width           =   9630
   Begin CommonDialog CMDialog1 
      Left            =   4560
      Top             =   6240
   End
   Begin SSFrame fraIonsInSystem 
      Caption         =   "Ions in System"
      ForeColor       =   &H00000000&
      Height          =   1992
      Left            =   120
      TabIndex        =   65
      Top             =   1140
      Width           =   4572
      Begin ComboBox cboIons 
         Height          =   288
         Index           =   2
         Left            =   2940
         Style           =   2  'Dropdown List
         TabIndex        =   79
         Top             =   1020
         Width           =   1512
      End
      Begin ComboBox cboIons 
         Enabled         =   0   'False
         Height          =   288
         Index           =   1
         Left            =   1500
         Style           =   2  'Dropdown List
         TabIndex        =   76
         Top             =   1560
         Width           =   1272
      End
      Begin ComboBox cboIons 
         Height          =   288
         Index           =   0
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   75
         Top             =   1560
         Width           =   1272
      End
      Begin CommandButton cmdEditProperties 
         Caption         =   "Edit Properties"
         Height          =   252
         Left            =   2940
         TabIndex        =   74
         Top             =   1380
         Width           =   1512
      End
      Begin CommandButton cmdAddDeleteIons 
         Caption         =   "Add Anion"
         Height          =   252
         Index           =   2
         Left            =   2940
         TabIndex        =   73
         Top             =   540
         Width           =   1512
      End
      Begin CommandButton cmdAddDeleteIons 
         Caption         =   "Remove Ion"
         Height          =   252
         Index           =   1
         Left            =   2940
         TabIndex        =   72
         Top             =   1620
         Width           =   1512
      End
      Begin CommandButton cmdAddDeleteIons 
         Caption         =   "Add Cation"
         Height          =   252
         Index           =   0
         Left            =   2940
         TabIndex        =   71
         Top             =   180
         Width           =   1512
      End
      Begin ListBox lstIons 
         Enabled         =   0   'False
         Height          =   810
         Index           =   1
         Left            =   1500
         MultiSelect     =   1  'Simple
         TabIndex        =   67
         Top             =   480
         Width           =   1275
      End
      Begin ListBox lstIons 
         Height          =   810
         Index           =   0
         Left            =   120
         MultiSelect     =   1  'Simple
         TabIndex        =   66
         Top             =   480
         Width           =   1275
      End
      Begin Shape Shape3 
         Height          =   972
         Left            =   2880
         Top             =   960
         Width           =   1632
      End
      Begin Label lblAnions 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Presaturant"
         Height          =   192
         Index           =   1
         Left            =   1500
         TabIndex        =   78
         Top             =   1320
         Width           =   1272
      End
      Begin Label lblCations 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Presaturant"
         Height          =   192
         Index           =   1
         Left            =   120
         TabIndex        =   77
         Top             =   1320
         Width           =   1272
      End
      Begin Shape Shape2 
         Height          =   1692
         Left            =   1440
         Top             =   240
         Width           =   1392
      End
      Begin Shape Shape1 
         Height          =   1692
         Left            =   60
         Top             =   240
         Width           =   1392
      End
      Begin Label lblAnions 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Anions"
         Enabled         =   0   'False
         Height          =   192
         Index           =   0
         Left            =   1560
         TabIndex        =   69
         Top             =   240
         Width           =   1212
      End
      Begin Label lblCations 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cations"
         Height          =   192
         Index           =   0
         Left            =   120
         TabIndex        =   68
         Top             =   240
         Width           =   1212
      End
   End
   Begin SSFrame fraOperatingConditions 
      Caption         =   "Operating Conditions"
      ForeColor       =   &H00000000&
      Height          =   972
      Left            =   120
      TabIndex        =   58
      Top             =   60
      Width           =   4572
      Begin ComboBox cboOperatingConditionsUnits 
         BackColor       =   &H00C0C0C0&
         Height          =   288
         Index           =   1
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   600
         Width           =   1512
      End
      Begin ComboBox cboOperatingConditionsUnits 
         BackColor       =   &H00C0C0C0&
         Height          =   288
         Index           =   0
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   240
         Width           =   1512
      End
      Begin TextBox txtOperatingConditions 
         Height          =   285
         Index           =   1
         Left            =   1800
         TabIndex        =   63
         Top             =   600
         Width           =   1095
      End
      Begin TextBox txtOperatingConditions 
         Height          =   285
         Index           =   0
         Left            =   1800
         TabIndex        =   62
         Top             =   240
         Width           =   1095
      End
      Begin Label lblOperatingConditions 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pressure"
         Height          =   192
         Index           =   0
         Left            =   120
         TabIndex        =   61
         Top             =   300
         Width           =   1536
      End
      Begin Label lblOperatingConditions 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Temperature"
         Height          =   192
         Index           =   1
         Left            =   120
         TabIndex        =   60
         Top             =   660
         Width           =   1536
      End
   End
   Begin SSFrame fraKineticDimensionless 
      Caption         =   "Kinetic Parameters and Dimensionless Groups"
      ForeColor       =   &H00000000&
      Height          =   3072
      Left            =   4860
      TabIndex        =   34
      Top             =   3240
      Width           =   4572
      Begin SSFrame fraDimensionlessGroups 
         Caption         =   "Dimensionless Groups"
         ForeColor       =   &H00000000&
         Height          =   1752
         Left            =   2400
         TabIndex        =   38
         Top             =   780
         Width           =   2112
         Begin Label lblKineticDimensionless 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Bip"
            Height          =   192
            Index           =   9
            Left            =   60
            TabIndex        =   83
            Top             =   1500
            Width           =   432
         End
         Begin Label lblKineticDimensionlessValue 
            BorderStyle     =   1  'Fixed Single
            Height          =   192
            Index           =   8
            Left            =   600
            TabIndex        =   84
            Top             =   1500
            Width           =   1092
         End
         Begin Label lblKineticDimensionlessUnits 
            BackStyle       =   0  'Transparent
            Caption         =   "(-)"
            Height          =   192
            Index           =   8
            Left            =   1800
            TabIndex        =   85
            Top             =   1500
            Width           =   252
         End
         Begin Label lblKineticDimensionless 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Dgp"
            Height          =   192
            Index           =   8
            Left            =   60
            TabIndex        =   86
            Top             =   540
            Width           =   432
         End
         Begin Label lblKineticDimensionlessValue 
            BorderStyle     =   1  'Fixed Single
            Height          =   192
            Index           =   4
            Left            =   600
            TabIndex        =   87
            Top             =   540
            Width           =   1092
         End
         Begin Label lblKineticDimensionlessUnits 
            BackStyle       =   0  'Transparent
            Caption         =   "(-)"
            Height          =   192
            Index           =   4
            Left            =   1800
            TabIndex        =   88
            Top             =   540
            Width           =   252
         End
         Begin Label lblKineticDimensionless 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Dgs"
            Height          =   192
            Index           =   7
            Left            =   60
            TabIndex        =   82
            Top             =   300
            Width           =   432
         End
         Begin Label lblKineticDimensionlessValue 
            BorderStyle     =   1  'Fixed Single
            Height          =   195
            Index           =   3
            Left            =   600
            TabIndex        =   81
            Top             =   300
            Width           =   1095
         End
         Begin Label lblKineticDimensionlessUnits 
            BackStyle       =   0  'Transparent
            Caption         =   "(-)"
            Height          =   192
            Index           =   3
            Left            =   1800
            TabIndex        =   80
            Top             =   300
            Width           =   252
         End
         Begin Label lblKineticDimensionlessUnits 
            BackStyle       =   0  'Transparent
            Caption         =   "(-)"
            Height          =   192
            Index           =   7
            Left            =   1800
            TabIndex        =   39
            Top             =   1260
            Width           =   252
         End
         Begin Label lblKineticDimensionlessUnits 
            BackStyle       =   0  'Transparent
            Caption         =   "(-)"
            Height          =   192
            Index           =   6
            Left            =   1800
            TabIndex        =   40
            Top             =   1020
            Width           =   252
         End
         Begin Label lblKineticDimensionlessUnits 
            BackStyle       =   0  'Transparent
            Caption         =   "(-)"
            Height          =   192
            Index           =   5
            Left            =   1800
            TabIndex        =   41
            Top             =   780
            Width           =   252
         End
         Begin Label lblKineticDimensionlessValue 
            BorderStyle     =   1  'Fixed Single
            Height          =   192
            Index           =   7
            Left            =   600
            TabIndex        =   42
            Top             =   1260
            Width           =   1092
         End
         Begin Label lblKineticDimensionlessValue 
            BorderStyle     =   1  'Fixed Single
            Height          =   192
            Index           =   6
            Left            =   600
            TabIndex        =   43
            Top             =   1020
            Width           =   1092
         End
         Begin Label lblKineticDimensionlessValue 
            BorderStyle     =   1  'Fixed Single
            Height          =   192
            Index           =   5
            Left            =   600
            TabIndex        =   44
            Top             =   780
            Width           =   1092
         End
         Begin Label lblKineticDimensionless 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "St"
            Height          =   192
            Index           =   6
            Left            =   60
            TabIndex        =   45
            Top             =   1260
            Width           =   432
         End
         Begin Label lblKineticDimensionless 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Edp"
            Height          =   192
            Index           =   5
            Left            =   60
            TabIndex        =   46
            Top             =   1020
            Width           =   432
         End
         Begin Label lblKineticDimensionless 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Dgt"
            Height          =   192
            Index           =   4
            Left            =   60
            TabIndex        =   47
            Top             =   780
            Width           =   432
         End
      End
      Begin SSFrame fraKinetic 
         Caption         =   "Kinetic Parameters"
         ForeColor       =   &H00000000&
         Height          =   1332
         Left            =   60
         TabIndex        =   51
         Top             =   780
         Width           =   2292
         Begin Label lblKineticDimensionlessUnits 
            BackStyle       =   0  'Transparent
            Caption         =   "cm2/s"
            Height          =   192
            Index           =   2
            Left            =   1620
            TabIndex        =   48
            Top             =   1020
            Width           =   612
         End
         Begin Label lblKineticDimensionlessUnits 
            BackStyle       =   0  'Transparent
            Caption         =   "cm2/s"
            Height          =   192
            Index           =   1
            Left            =   1620
            TabIndex        =   49
            Top             =   660
            Width           =   612
         End
         Begin Label lblKineticDimensionlessUnits 
            BackStyle       =   0  'Transparent
            Caption         =   "cm/s"
            Height          =   192
            Index           =   0
            Left            =   1620
            TabIndex        =   50
            Top             =   300
            Width           =   612
         End
         Begin Label lblKineticDimensionlessValue 
            BorderStyle     =   1  'Fixed Single
            Height          =   192
            Index           =   2
            Left            =   480
            TabIndex        =   57
            Top             =   1020
            Width           =   1092
         End
         Begin Label lblKineticDimensionlessValue 
            BorderStyle     =   1  'Fixed Single
            Height          =   192
            Index           =   1
            Left            =   480
            TabIndex        =   56
            Top             =   660
            Width           =   1092
         End
         Begin Label lblKineticDimensionlessValue 
            BorderStyle     =   1  'Fixed Single
            Height          =   192
            Index           =   0
            Left            =   480
            TabIndex        =   55
            Top             =   300
            Width           =   1092
         End
         Begin Label lblKineticDimensionless 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Dp"
            Height          =   192
            Index           =   3
            Left            =   60
            TabIndex        =   54
            Top             =   1020
            Width           =   312
         End
         Begin Label lblKineticDimensionless 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Dl"
            Height          =   192
            Index           =   2
            Left            =   60
            TabIndex        =   53
            Top             =   660
            Width           =   312
         End
         Begin Label lblKineticDimensionless 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "kf"
            Height          =   192
            Index           =   1
            Left            =   60
            TabIndex        =   52
            Top             =   300
            Width           =   312
         End
      End
      Begin CommandButton cmdInputKineticParameters 
         Caption         =   "View Kinetic Parameters"
         Height          =   312
         Left            =   180
         TabIndex        =   37
         Top             =   2640
         Width           =   4212
      End
      Begin ComboBox cboKinDimComponent 
         Height          =   288
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   360
         Width           =   2532
      End
      Begin Label lblKineticDimensionless 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Select component:"
         Height          =   192
         Index           =   0
         Left            =   120
         TabIndex        =   35
         Top             =   420
         Width           =   1632
      End
   End
   Begin SSFrame fraAdsorbentProperties 
      Caption         =   "Adsorbent Properties"
      ForeColor       =   &H00000000&
      Height          =   3072
      Left            =   120
      TabIndex        =   16
      Top             =   3240
      Width           =   4572
      Begin ComboBox cboAdsorbents 
         Height          =   288
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   70
         Top             =   720
         Width           =   2712
      End
      Begin ComboBox cboAdsorbentPropertyUnits 
         BackColor       =   &H00C0C0C0&
         Height          =   300
         Index           =   5
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   2520
         Width           =   1515
      End
      Begin ComboBox cboAdsorbentPropertyUnits 
         BackColor       =   &H00C0C0C0&
         Height          =   300
         Index           =   2
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1515
      End
      Begin ComboBox cboAdsorbentPropertyUnits 
         BackColor       =   &H00C0C0C0&
         Height          =   300
         Index           =   1
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1515
      End
      Begin TextBox txtAdsorbentProperties 
         Height          =   285
         Index           =   5
         Left            =   1800
         TabIndex        =   28
         Top             =   2520
         Width           =   1092
      End
      Begin TextBox txtAdsorbentProperties 
         Height          =   285
         Index           =   4
         Left            =   1800
         TabIndex        =   27
         Top             =   2160
         Width           =   1092
      End
      Begin TextBox txtAdsorbentProperties 
         Height          =   285
         Index           =   3
         Left            =   1800
         TabIndex        =   26
         Top             =   1800
         Width           =   1095
      End
      Begin TextBox txtAdsorbentProperties 
         Height          =   285
         Index           =   2
         Left            =   1800
         TabIndex        =   25
         Top             =   1440
         Width           =   1095
      End
      Begin TextBox txtAdsorbentProperties 
         Height          =   285
         Index           =   1
         Left            =   1800
         TabIndex        =   24
         Top             =   1080
         Width           =   1095
      End
      Begin CommandButton cmdAdsorbentDatabase 
         Caption         =   "Adsorbent Database"
         Enabled         =   0   'False
         Height          =   315
         Left            =   180
         TabIndex        =   17
         Top             =   300
         Width           =   4215
      End
      Begin Label lblAdsorbentProperties 
         BackStyle       =   0  'Transparent
         Caption         =   "(-)"
         Height          =   195
         Index           =   7
         Left            =   3000
         TabIndex        =   31
         Top             =   2220
         Width           =   1095
      End
      Begin Label lblAdsorbentProperties 
         BackStyle       =   0  'Transparent
         Caption         =   "(-)"
         Height          =   195
         Index           =   6
         Left            =   3000
         TabIndex        =   30
         Top             =   1860
         Width           =   1095
      End
      Begin Label lblAdsorbentProperties 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Capacity"
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   23
         Top             =   2580
         Width           =   1540
      End
      Begin Label lblAdsorbentProperties 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tortuosity"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   22
         Top             =   2220
         Width           =   1540
      End
      Begin Label lblAdsorbentProperties 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Particle Porosity"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   21
         Top             =   1860
         Width           =   1540
      End
      Begin Label lblAdsorbentProperties 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Particle Radius"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   20
         Top             =   1500
         Width           =   1540
      End
      Begin Label lblAdsorbentProperties 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Apparent Density"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   19
         Top             =   1140
         Width           =   1540
      End
      Begin Label lblAdsorbentProperties 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   18
         Top             =   780
         Width           =   1540
      End
   End
   Begin SSFrame fraBedData 
      Caption         =   "Bed Data:"
      ForeColor       =   &H00000000&
      Height          =   2715
      Left            =   4860
      TabIndex        =   0
      Top             =   60
      Width           =   4575
      Begin ComboBox cboBedDataUnits 
         BackColor       =   &H00C0C0C0&
         Height          =   300
         Index           =   4
         Left            =   3420
         Style           =   2  'Dropdown List
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1740
         Width           =   1095
      End
      Begin ComboBox cboBedDataUnits 
         BackColor       =   &H00C0C0C0&
         Height          =   300
         Index           =   3
         Left            =   3420
         Style           =   2  'Dropdown List
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1380
         Width           =   1095
      End
      Begin ComboBox cboBedDataUnits 
         BackColor       =   &H00C0C0C0&
         Height          =   300
         Index           =   2
         Left            =   3420
         Style           =   2  'Dropdown List
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1095
      End
      Begin ComboBox cboBedDataUnits 
         BackColor       =   &H00C0C0C0&
         Height          =   300
         Index           =   1
         Left            =   3420
         Style           =   2  'Dropdown List
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   660
         Width           =   1095
      End
      Begin ComboBox cboBedDataUnits 
         BackColor       =   &H00C0C0C0&
         Height          =   300
         Index           =   0
         Left            =   3420
         Style           =   2  'Dropdown List
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   300
         Width           =   1095
      End
      Begin TextBox txtBedData 
         Height          =   288
         Index           =   4
         Left            =   2220
         TabIndex        =   10
         Top             =   1740
         Width           =   1092
      End
      Begin TextBox txtBedData 
         Height          =   288
         Index           =   3
         Left            =   2220
         TabIndex        =   9
         Top             =   1380
         Width           =   1092
      End
      Begin TextBox txtBedData 
         Height          =   288
         Index           =   2
         Left            =   2220
         TabIndex        =   8
         Top             =   1020
         Width           =   1092
      End
      Begin TextBox txtBedData 
         Height          =   288
         Index           =   1
         Left            =   2220
         TabIndex        =   7
         Top             =   660
         Width           =   1092
      End
      Begin TextBox txtBedData 
         Height          =   288
         Index           =   0
         Left            =   2220
         TabIndex        =   6
         Top             =   300
         Width           =   1092
      End
      Begin Label lblBedData 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "EBCT"
         Height          =   192
         Index           =   4
         Left            =   180
         TabIndex        =   5
         Top             =   1800
         Width           =   1800
      End
      Begin Label lblBedData 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Flow Rate"
         Height          =   192
         Index           =   3
         Left            =   180
         TabIndex        =   4
         Top             =   1440
         Width           =   1800
      End
      Begin Label lblBedData 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Bed Mass"
         Height          =   192
         Index           =   2
         Left            =   180
         TabIndex        =   3
         Top             =   1080
         Width           =   1800
      End
      Begin Label lblBedData 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Adsorber Diameter"
         Height          =   192
         Index           =   1
         Left            =   180
         TabIndex        =   2
         Top             =   720
         Width           =   1800
      End
      Begin Label lblBedData 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Bed Length"
         Height          =   192
         Index           =   0
         Left            =   180
         TabIndex        =   1
         Top             =   360
         Width           =   1800
      End
   End
   Begin Menu mnuFileMenu 
      Caption         =   "&File"
      Begin Menu mnuFile 
         Caption         =   "&New"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin Menu mnuFile 
         Caption         =   "&Open"
         Index           =   1
      End
      Begin Menu mnuFile 
         Caption         =   "&Save"
         Index           =   2
      End
      Begin Menu mnuFile 
         Caption         =   "Save &As"
         Index           =   3
      End
      Begin Menu mnuFile 
         Caption         =   "-"
         Index           =   4
      End
      Begin Menu mnuFile 
         Caption         =   "Select P&rinter"
         Index           =   5
      End
      Begin Menu mnuFile 
         Caption         =   "&Print"
         Index           =   6
         Begin Menu mnuFilePrint 
            Caption         =   "To &Printer"
            Index           =   0
         End
         Begin Menu mnuFilePrint 
            Caption         =   "To &File"
            Index           =   1
         End
      End
      Begin Menu mnuFile 
         Caption         =   "-"
         Index           =   7
      End
      Begin Menu mnuFile 
         Caption         =   "E&xit"
         Index           =   8
      End
   End
   Begin Menu mnuRunMenu 
      Caption         =   "&Run"
      Begin Menu mnuRun 
         Caption         =   "&PFPDM"
         Index           =   0
      End
   End
   Begin Menu mnuResultsMenu 
      Caption         =   "Re&sults"
      Begin Menu mnuResults 
         Caption         =   "&PFPDM"
         Index           =   0
      End
      Begin Menu mnuResults 
         Caption         =   "Compare to &Data"
         Index           =   1
      End
   End
   Begin Menu mnuOptionsMenu 
      Caption         =   "&Options"
      Begin Menu mnuOptions 
         Caption         =   "&Influent Concentrations"
         Index           =   0
      End
      Begin Menu mnuOptions 
         Caption         =   "Set &Number of Beds"
         Index           =   1
      End
      Begin Menu mnuOptions 
         Caption         =   "Set &Time Parameters"
         Index           =   2
      End
      Begin Menu mnuOptions 
         Caption         =   "Set &Collocation Points"
         Index           =   3
      End
      Begin Menu mnuOptions 
         Caption         =   "Set &Resin Phase Presaturant Conditions"
         Enabled         =   0   'False
         Index           =   4
      End
   End
End
Option Explicit

Dim Temp_Text As String
Dim IsError As Integer

Sub cboAdsorbentPropertyUnits_Click (Index As Integer)
    Dim ValueToDisplay As Double

    Select Case Index
       Case 1   'Apparent Density
            Select Case cboAdsorbentPropertyUnits(1).ListIndex
               Case APPARENT_DENSITY_G_per_ML    'g/ml
                    ValueToDisplay = Resin.ApparentDensity
               Case APPARENT_DENSITY_KG_per_M3    'kg/m3
                    ValueToDisplay = Resin.ApparentDensity * DensityConversionFactor(APPARENT_DENSITY_KG_per_M3)
               Case APPARENT_DENSITY_LB_per_FT3    'lb/ft3
                    ValueToDisplay = Resin.ApparentDensity * DensityConversionFactor(APPARENT_DENSITY_LB_per_FT3)
               Case APPARENT_DENSITY_LB_per_GAL    'lb/gal
                    ValueToDisplay = Resin.ApparentDensity * DensityConversionFactor(APPARENT_DENSITY_LB_per_GAL)
            End Select
            txtAdsorbentProperties(1).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       Case 2   'Particle Radius
            Select Case cboAdsorbentPropertyUnits(2).ListIndex
               Case LENGTH_M    'm
                    ValueToDisplay = Resin.ParticleRadius
               Case LENGTH_CM   'cm
                    ValueToDisplay = Resin.ParticleRadius * LengthConversionFactor(LENGTH_CM)
               Case LENGTH_FT   'ft
                    ValueToDisplay = Resin.ParticleRadius * LengthConversionFactor(LENGTH_FT)
               Case LENGTH_IN   'in
                    ValueToDisplay = Resin.ParticleRadius * LengthConversionFactor(LENGTH_IN)
            End Select
            txtAdsorbentProperties(2).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       Case 5   'Total Resin Capacity
            Select Case cboAdsorbentPropertyUnits(5).ListIndex
               Case RESIN_CAPACITY_MEQ_per_G   'meq/g resin
                    ValueToDisplay = Resin.TotalCapacity
               Case RESIN_CAPACITY_MEQ_per_MLbed   'meq/ml bed
                    ValueToDisplay = Resin.TotalCapacity * ResinCapacityConversionFactor(RESIN_CAPACITY_MEQ_per_MLbed)
               Case RESIN_CAPACITY_MEQ_per_MLresin   'meq/ml resin
                    ValueToDisplay = Resin.TotalCapacity * ResinCapacityConversionFactor(RESIN_CAPACITY_MEQ_per_MLresin)
            End Select
            txtAdsorbentProperties(5).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

    End Select

End Sub

Sub cboAdsorbents_Click ()
    Dim i As Integer

    Resin.Name = cboAdsorbents.List(cboAdsorbents.ListIndex)

    Select Case cboAdsorbents.ListIndex
       Case 0   'IRN-77
            'enable Cations List Box and Disable Anions list box for IRN-77
            cboIons(0).Enabled = True
            lstIons(0).Enabled = True
            lblCations(0).Enabled = True
            lblCations(1).Enabled = True
            cmdAddDeleteIons(0).Enabled = True
            cmdAddDeleteIons(1).Enabled = True

            cboIons(1).Enabled = False
            lstIons(1).Enabled = False
            lblAnions(0).Enabled = False
            lblAnions(1).Enabled = False
            cmdAddDeleteIons(2).Enabled = False

            cboIons(2).Clear
            cboKinDimComponent.Clear
            frmInputKineticParameters!cboIon.Clear
            'Load cations into Kinetic Parameters combo Box and onto frmInputKineticParameters
            If cboIons(0).ListCount > 0 Then
               For i = 0 To cboIons(0).ListCount - 1
                   cboIons(2).AddItem cboIons(0).List(i)
'                   cboKinDimComponent.AddItem cboIons(0).List(i)
                   frmInputKineticParameters!cboIon.AddItem cboIons(0).List(i)
               Next i
               cboIons(2).ListIndex = 0
'               cboKinDimComponent.ListIndex = 0
               frmInputKineticParameters!cboIon.ListIndex = 0
            End If
            Cations.Available = True
            Anions.Available = False
       Case 1   'IRN-78
            'disable Cations List Box and enable Anions list box for IRN-78
            cboIons(0).Enabled = False
            lstIons(0).Enabled = False
            lblCations(0).Enabled = False
            lblCations(1).Enabled = False
            cmdAddDeleteIons(0).Enabled = False
            cmdAddDeleteIons(1).Enabled = False

            cboIons(1).Enabled = True
            lstIons(1).Enabled = True
            lblAnions(0).Enabled = True
            lblAnions(1).Enabled = True
            cmdAddDeleteIons(2).Enabled = True

            cboIons(2).Clear
            cboKinDimComponent.Clear
            frmInputKineticParameters!cboIon.Clear
            'Load anions into Kinetic Parameters combo Box
            If cboIons(1).ListCount > 0 Then
               For i = 0 To cboIons(1).ListCount - 1
                   cboIons(2).AddItem cboIons(1).List(i)
'                   cboKinDimComponent.AddItem cboIons(1).List(i)
                   frmInputKineticParameters!cboIon.AddItem cboIons(1).List(i)
               Next i
               cboIons(2).ListIndex = 0
'               cboKinDimComponent.ListIndex = 0
               frmInputKineticParameters!cboIon.ListIndex = 0
            End If

            Cations.Available = False
            Anions.Available = True

       Case 2   'IRA-68
            'Enable both Anions and Cations for IRA-68
            cboIons(0).Enabled = True
            lstIons(0).Enabled = True
            lblCations(0).Enabled = True
            lblCations(1).Enabled = True
            cmdAddDeleteIons(0).Enabled = True
            cmdAddDeleteIons(1).Enabled = True

            cboIons(1).Enabled = True
            lstIons(1).Enabled = True
            lblAnions(0).Enabled = True
            lblAnions(1).Enabled = True
            cmdAddDeleteIons(2).Enabled = True

            cboIons(2).Clear
            cboKinDimComponent.Clear
            frmInputKineticParameters!cboIon.Clear

            If cboIons(0).ListCount > 0 Or cboIons(1).ListCount > 0 Then
               'Load cations into Kinetic Parameters combo Box
               For i = 0 To cboIons(0).ListCount - 1
                   cboIons(2).AddItem cboIons(0).List(i)
'                   cboKinDimComponent.AddItem cboIons(0).List(i)
                   frmInputKineticParameters!cboIon.AddItem cboIons(0).List(i)
               Next i
            
               'Load anions into Kinetic Parameters combo Box
               For i = 0 To cboIons(1).ListCount - 1
                   cboIons(2).AddItem cboIons(1).List(i)
'                   cboKinDimComponent.AddItem cboIons(1).List(i)
                   frmInputKineticParameters!cboIon.AddItem cboIons(1).List(i)
               Next i
               cboIons(2).ListIndex = 0
'               cboKinDimComponent.ListIndex = 0
               frmInputKineticParameters!cboIon.ListIndex = 0
            End If

            Cations.Available = True
            Anions.Available = True
    End Select

End Sub

Sub cboBedDataUnits_Click (Index As Integer)
    Dim ValueToDisplay As Double

    Select Case Index
       Case 0   'Bed Length
            Select Case cboBedDataUnits(0).ListIndex
               Case LENGTH_M    'm
                    ValueToDisplay = Bed.Length
               Case LENGTH_CM   'cm
                    ValueToDisplay = Bed.Length * LengthConversionFactor(LENGTH_CM)
               Case LENGTH_FT   'ft
                    ValueToDisplay = Bed.Length * LengthConversionFactor(LENGTH_FT)
               Case LENGTH_IN   'in
                    ValueToDisplay = Bed.Length * LengthConversionFactor(LENGTH_IN)
            End Select
            txtBedData(0).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       Case 1   'Bed Diameter
            Select Case cboBedDataUnits(1).ListIndex
               Case LENGTH_M    'm
                    ValueToDisplay = Bed.Diameter
               Case LENGTH_CM   'cm
                    ValueToDisplay = Bed.Diameter * LengthConversionFactor(LENGTH_CM)
               Case LENGTH_FT   'ft
                    ValueToDisplay = Bed.Diameter * LengthConversionFactor(LENGTH_FT)
               Case LENGTH_IN   'in
                    ValueToDisplay = Bed.Diameter * LengthConversionFactor(LENGTH_IN)
            End Select
            txtBedData(1).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       Case 2   'Bed Mass
            Select Case cboBedDataUnits(2).ListIndex
               Case MASS_KG   'kg
                    ValueToDisplay = Bed.Weight
               Case MASS_G    'g
                    ValueToDisplay = Bed.Weight * MassConversionFactor(MASS_G)
               Case MASS_LB   'lb
                    ValueToDisplay = Bed.Weight * MassConversionFactor(MASS_LB)
            End Select
            txtBedData(2).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       Case 3   'Flow Rate
            Select Case cboBedDataUnits(3).ListIndex
               Case FLOW_M3_per_S     'm3/s
                    ValueToDisplay = Bed.Flowrate.Value
               Case FLOW_M3_per_D     'm3/d
                    ValueToDisplay = Bed.Flowrate.Value * FlowConversionFactor(FLOW_M3_per_D)
               Case FLOW_CM3_per_S    'cm3/s
                    ValueToDisplay = Bed.Flowrate.Value * FlowConversionFactor(FLOW_CM3_per_S)
               Case FLOW_ML_per_MIN   'ml/min
                    ValueToDisplay = Bed.Flowrate.Value * FlowConversionFactor(FLOW_ML_per_MIN)
               Case FLOW_FT3_per_S    'ft3/s
                    ValueToDisplay = Bed.Flowrate.Value * FlowConversionFactor(FLOW_FT3_per_S)
               Case FLOW__FT3_per_D   'ft3/d
                    ValueToDisplay = Bed.Flowrate.Value * FlowConversionFactor(FLOW__FT3_per_D)
               Case FLOW_GPM   'gpm
                    ValueToDisplay = Bed.Flowrate.Value * FlowConversionFactor(FLOW_GPM)
               Case FLOW_GPD   'gpd
                    ValueToDisplay = Bed.Flowrate.Value * FlowConversionFactor(FLOW_GPD)
               Case FLOW_MGD   'MGD
                    ValueToDisplay = Bed.Flowrate.Value * FlowConversionFactor(FLOW_MGD)
            End Select
            txtBedData(3).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       Case 4   'EBCT
            Select Case cboBedDataUnits(4).ListIndex
               Case TIME_MIN   'min
                    ValueToDisplay = Bed.EBCT.Value
               Case TIME_S     's
                    ValueToDisplay = Bed.EBCT.Value * TimeConversionFactor(TIME_S)
               Case TIME_HR    'hr
                    ValueToDisplay = Bed.EBCT.Value * TimeConversionFactor(TIME_HR)
               Case TIME_D     'd
                    ValueToDisplay = Bed.EBCT.Value * TimeConversionFactor(TIME_D)
            End Select
            txtBedData(4).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))


    End Select

End Sub

Sub cboIons_Click (Index As Integer)
    Dim i As Integer

    If Index = 2 Then
       Exit Sub
    End If

    If Index = 0 Then      'Cations
       PresaturantCation = cboIons(0).ListIndex + 1
    ElseIf Index = 1 Then  'anions
       PresaturantAnion = cboIons(1).ListIndex + 1
    End If
    
    lstIons(Index).Clear
    For i = 0 To cboIons(Index).ListCount - 1
        If i <> cboIons(Index).ListIndex Then
           lstIons(Index).AddItem cboIons(Index).List(i)
        End If
    Next i

    For i = 0 To lstIons(Index).ListCount - 1
        lstIons(Index).Selected(i) = False
    Next i

    For i = 1 To MAX_CHEMICAL
        Resin.PresaturantPercentage(i) = 0#
    Next i

    'Start Cations_Selected or Anions_Selected Arrays
    Select Case Index
       Case 0   'Cations
          NumSelectedCations = 1
          Cations_Selected(1) = PresaturantCation
          Resin.PresaturantPercentage(Cations_Selected(1)) = 100#
       Case 1   'Anions
          NumSelectedAnions = 1
          Anions_Selected(1) = PresaturantAnion
          Resin.PresaturantPercentage(Anions_Selected(1)) = 100#
    End Select

    cboKinDimComponent.Clear
    cboKinDimComponent.Enabled = False

    For i = 3 To 8
        lblKineticDimensionlessValue(i).Caption = ""
    Next i

    mnuRun(0).Enabled = False
    mnuOptions(4).Enabled = False

End Sub

Sub cboKinDimComponent_Click ()
    Dim i As Integer, ListIndex As Integer
    Dim FoundAnion As Integer, FoundCation As Integer
    Dim ValueToDisplay As Double
    Dim NumberOfIonFound As Integer

    ListIndex = cboKinDimComponent.ListIndex

    'Display values for kf, Dl, and Dp on Main form

    If Cations.Available And Anions.Available Then   'May be editing either cations or anions
       FoundCation = False
       FoundAnion = False
       For i = 0 To cboKinDimComponent.ListCount - 1
           If Cation(Cations_Selected(i + 1)).Name = cboKinDimComponent.List(ListIndex) Then
              FoundCation = True
              NumberOfIonFound = Cations_Selected(i + 1)
              Exit For
           End If

           If Anion(Anions_Selected(i + 1)).Name = cboKinDimComponent.List(ListIndex) Then
              FoundAnion = True
              NumberOfIonFound = Anions_Selected(i + 1)
              Exit For
           End If
       Next i

       If FoundCation Then

          'Display kf
          ValueToDisplay = Cation(NumberOfIonFound).Kinetic.IonicTransportCoefficient.Value
          lblKineticDimensionlessValue(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          'Display Dl
          ValueToDisplay = Cation(NumberOfIonFound).Kinetic.LiquidDiffusivity.Value
          lblKineticDimensionlessValue(1).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          'Display Dp
          ValueToDisplay = Cation(NumberOfIonFound).Kinetic.PoreDiffusivity.Value
          lblKineticDimensionlessValue(2).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          If Not OKToGetCationDimensionless Then
             For i = 3 To 8
                 lblKineticDimensionlessValue(i).Caption = ""
             Next i
             Exit Sub
          End If

          'Display Dgs
          ValueToDisplay = Cation(NumberOfIonFound).Dimensionless.SurfaceDistributionParameter
          lblKineticDimensionlessValue(3).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          'Display Dgp
          ValueToDisplay = Cation(NumberOfIonFound).Dimensionless.PoreDistributionParameter
          lblKineticDimensionlessValue(4).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          'Display Dgt
          ValueToDisplay = Cation(NumberOfIonFound).Dimensionless.TotalDistributionParameter
          lblKineticDimensionlessValue(5).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          'Display Edp
          ValueToDisplay = Cation(NumberOfIonFound).Dimensionless.PoreDiffusionModulus
          lblKineticDimensionlessValue(6).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
     
          'Display St
          ValueToDisplay = Cation(NumberOfIonFound).Dimensionless.StantonNumber
          lblKineticDimensionlessValue(7).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          'Display Bip
          ValueToDisplay = Cation(NumberOfIonFound).Dimensionless.PoreBiotNumber
          lblKineticDimensionlessValue(8).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
       
       End If
       
       If FoundAnion Then
          'Display kf
          ValueToDisplay = Anion(NumberOfIonFound).Kinetic.IonicTransportCoefficient.Value
          lblKineticDimensionlessValue(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          'Display Dl
          ValueToDisplay = Anion(NumberOfIonFound).Kinetic.LiquidDiffusivity.Value
          lblKineticDimensionlessValue(1).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          'Display Dp
          ValueToDisplay = Anion(NumberOfIonFound).Kinetic.PoreDiffusivity.Value
          lblKineticDimensionlessValue(2).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          If Not OKToGetAnionDimensionless Then
             For i = 3 To 8
                 lblKineticDimensionlessValue(i).Caption = ""
             Next i
             Exit Sub
          End If

          'Display Dgs
          ValueToDisplay = Anion(NumberOfIonFound).Dimensionless.SurfaceDistributionParameter
          lblKineticDimensionlessValue(3).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          'Display Dgp
          ValueToDisplay = Anion(NumberOfIonFound).Dimensionless.PoreDistributionParameter
          lblKineticDimensionlessValue(4).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          'Display Dgt
          ValueToDisplay = Anion(NumberOfIonFound).Dimensionless.TotalDistributionParameter
          lblKineticDimensionlessValue(5).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          'Display Edp
          ValueToDisplay = Anion(NumberOfIonFound).Dimensionless.PoreDiffusionModulus
          lblKineticDimensionlessValue(6).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
     
          'Display St
          ValueToDisplay = Anion(NumberOfIonFound).Dimensionless.StantonNumber
          lblKineticDimensionlessValue(7).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          'Display Bip
          ValueToDisplay = Anion(NumberOfIonFound).Dimensionless.PoreBiotNumber
          lblKineticDimensionlessValue(8).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       End If

    ElseIf Cations.Available Then  'Only cations in list
       FoundCation = False
       For i = 0 To cboKinDimComponent.ListCount - 1
           If Cation(Cations_Selected(i + 1)).Name = cboKinDimComponent.List(ListIndex) Then
              FoundCation = True
              NumberOfIonFound = Cations_Selected(i + 1)
              Exit For
           End If

       Next i
       

          'Display kf
          ValueToDisplay = Cation(NumberOfIonFound).Kinetic.IonicTransportCoefficient.Value
          lblKineticDimensionlessValue(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          'Display Dl
          ValueToDisplay = Cation(NumberOfIonFound).Kinetic.LiquidDiffusivity.Value
          lblKineticDimensionlessValue(1).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          'Display Dp
          ValueToDisplay = Cation(NumberOfIonFound).Kinetic.PoreDiffusivity.Value
          lblKineticDimensionlessValue(2).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          If Not OKToGetCationDimensionless Then
             For i = 3 To 8
                 lblKineticDimensionlessValue(i).Caption = ""
             Next i
             Exit Sub
          End If

          'Display Dgs
          ValueToDisplay = Cation(NumberOfIonFound).Dimensionless.SurfaceDistributionParameter
          lblKineticDimensionlessValue(3).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          'Display Dgp
          ValueToDisplay = Cation(NumberOfIonFound).Dimensionless.PoreDistributionParameter
          lblKineticDimensionlessValue(4).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          'Display Dgt
          ValueToDisplay = Cation(NumberOfIonFound).Dimensionless.TotalDistributionParameter
          lblKineticDimensionlessValue(5).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          'Display Edp
          ValueToDisplay = Cation(NumberOfIonFound).Dimensionless.PoreDiffusionModulus
          lblKineticDimensionlessValue(6).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
     
          'Display St
          ValueToDisplay = Cation(NumberOfIonFound).Dimensionless.StantonNumber
          lblKineticDimensionlessValue(7).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          'Display Bip
          ValueToDisplay = Cation(NumberOfIonFound).Dimensionless.PoreBiotNumber
          lblKineticDimensionlessValue(8).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

    ElseIf Anions.Available Then  'Only anions in list
       FoundAnion = False
       For i = 0 To cboKinDimComponent.ListCount - 1
           If Anion(Anions_Selected(i + 1)).Name = cboKinDimComponent.List(ListIndex) Then
              FoundAnion = True
              NumberOfIonFound = Anions_Selected(i + 1)
              Exit For
           End If
       Next i
      

          'Display kf
          ValueToDisplay = Anion(NumberOfIonFound).Kinetic.IonicTransportCoefficient.Value
          lblKineticDimensionlessValue(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          'Display Dl
          ValueToDisplay = Anion(NumberOfIonFound).Kinetic.LiquidDiffusivity.Value
          lblKineticDimensionlessValue(1).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          'Display Dp
          ValueToDisplay = Anion(NumberOfIonFound).Kinetic.PoreDiffusivity.Value
          lblKineticDimensionlessValue(2).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          If Not OKToGetAnionDimensionless Then
             For i = 3 To 8
                 lblKineticDimensionlessValue(i).Caption = ""
             Next i
             Exit Sub
          End If

          'Display Dgs
          ValueToDisplay = Anion(NumberOfIonFound).Dimensionless.SurfaceDistributionParameter
          lblKineticDimensionlessValue(3).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          'Display Dgp
          ValueToDisplay = Anion(NumberOfIonFound).Dimensionless.PoreDistributionParameter
          lblKineticDimensionlessValue(4).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          'Display Dgt
          ValueToDisplay = Anion(NumberOfIonFound).Dimensionless.TotalDistributionParameter
          lblKineticDimensionlessValue(5).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          'Display Edp
          ValueToDisplay = Anion(NumberOfIonFound).Dimensionless.PoreDiffusionModulus
          lblKineticDimensionlessValue(6).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
     
          'Display St
          ValueToDisplay = Anion(NumberOfIonFound).Dimensionless.StantonNumber
          lblKineticDimensionlessValue(7).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          'Display Bip
          ValueToDisplay = Anion(NumberOfIonFound).Dimensionless.PoreBiotNumber
          lblKineticDimensionlessValue(8).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

    End If
    
End Sub

Sub cboOperatingConditionsUnits_Click (Index As Integer)
    Dim ValueToDisplay As Double

    Select Case Index
       Case 0   'Operating Pressure
            Select Case cboOperatingConditionsUnits(0).ListIndex
               Case PRESSURE_PA    'Pa
                    ValueToDisplay = Operating.Pressure
               Case PRESSURE_KPA   'kPa
                    ValueToDisplay = Operating.Pressure * PressureConversionFactor(PRESSURE_KPA)
               Case PRESSURE_BARS   'bars
                    ValueToDisplay = Operating.Pressure * PressureConversionFactor(PRESSURE_BARS)
               Case PRESSURE_ATM   'atm
                    ValueToDisplay = Operating.Pressure * PressureConversionFactor(PRESSURE_ATM)
               Case PRESSURE_PSI   'psi
                    ValueToDisplay = Operating.Pressure * PressureConversionFactor(PRESSURE_PSI)
               Case PRESSURE_MMHG   'mm Hg
                    ValueToDisplay = Operating.Pressure * PressureConversionFactor(PRESSURE_MMHG)
               Case PRESSURE_MH2O   'm H2O
                    ValueToDisplay = Operating.Pressure * PressureConversionFactor(PRESSURE_MH2O)
               Case PRESSURE_FTH2O   'ft H2O
                    ValueToDisplay = Operating.Pressure * PressureConversionFactor(PRESSURE_FTH2O)
               Case PRESSURE_INHG   'in. Hg
                    ValueToDisplay = Operating.Pressure * PressureConversionFactor(PRESSURE_INHG)
            End Select
            txtOperatingConditions(0).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       Case 1   'Operating Temperature
            Select Case cboOperatingConditionsUnits(1).ListIndex
               Case TEMPERATURE_K    'K
                    ValueToDisplay = Operating.Temperature
               Case TEMPERATURE_C   'C
                    ValueToDisplay = TemperatureConversion(TEMPERATURE_C, Operating.Temperature)
               Case TEMPERATURE_R   'R
                    ValueToDisplay = TemperatureConversion(TEMPERATURE_R, Operating.Temperature)
               Case TEMPERATURE_F   'F
                    ValueToDisplay = TemperatureConversion(TEMPERATURE_F, Operating.Temperature)
            End Select
            txtOperatingConditions(1).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

    End Select

End Sub

Sub cmdAddDeleteIons_Click (Index As Integer)
    Dim i As Integer
    Dim msg As String
    Dim ListIndex As Integer

    Select Case Index
       Case 0   'Add Cation

            If NumberOfCations = MAX_CHEMICAL Then
               msg = "Adding another cation would exceed the maximum number of "
               msg = msg & "cations allowed for a simulation.  If you would like to "
               msg = msg + "add another cation, you must remove one of the "
               msg = msg + "current cations first."
               MsgBox msg, MB_ICONSTOP, "Too Many Cations"
            End If

            frmAddComponent.Caption = "Add Cation"
            frmAddComponent!lblValenceSign.Caption = "+"
            frmAddComponent!txtAddIon(0).Text = "Cation"
            frmAddComponent!txtAddIon(1).Text = Trim$(Str$(DefaultCation.MolecularWeight))
            frmAddComponent!txtAddIon(2).Text = Trim$(Str$(DefaultCation.InitialConcentration))
            frmAddComponent!lblValence.Caption = Trim$(Str$(CInt(DefaultCation.Valence)))
            If NumberOfCations > 0 Then
               frmAddComponent!txtAlphaValue.Text = Trim$(Str$(DefaultCation.SeparationFactor))
            Else
               frmAddComponent!txtAlphaValue.Text = "1.00"
            End If

            SeparationFactorInput.Row = CationSeparationFactorInput.Row
            SeparationFactorInput.Value = CationSeparationFactorInput.Value
            If SeparationFactorInput.Row = True Then
               frmAddComponent!lblAlpha(1).Caption = frmAddComponent!txtAddIon(0).Text
            Else
               frmAddComponent!lblAlpha(2).Caption = frmAddComponent!txtAddIon(0).Text
            End If

            If NumberOfCations > 0 Then
               If SeparationFactorInput.Row = True Then
                  frmAddComponent!lblAlpha(2).Caption = Cation(SeparationFactorInput.Value - 10).Name
               Else
                  frmAddComponent!lblAlpha(1).Caption = Cation(SeparationFactorInput.Value).Name
               End If
            Else
               If SeparationFactorInput.Row = True Then
                  frmAddComponent!lblAlpha(2).Caption = frmAddComponent!lblAlpha(1).Caption
               Else
                  frmAddComponent!lblAlpha(1).Caption = frmAddComponent!lblAlpha(2).Caption
               End If
            End If

            If Trim$(frmAddComponent!lblAlpha(1).Caption) = Trim$(frmAddComponent!lblAlpha(2).Caption) Then
               frmAddComponent!txtAlphaValue.Enabled = False
            Else
               frmAddComponent!txtAlphaValue.Enabled = True
            End If
                  
            For i = 0 To frmAddComponent!cboAnion.ListCount - 1
                If DefaultCation.Kinetic.NernstHaskellAnion.Ion_Name = frmAddComponent!cboAnion.List(i) Then
                   frmAddComponent!cboAnion.ListIndex = i
                End If
            Next i
            For i = 0 To frmAddComponent!cboCation.ListCount - 1
                If DefaultCation.Kinetic.NernstHaskellCation.Ion_Name = frmAddComponent!cboCation.List(i) Then
                   frmAddComponent!cboCation.ListIndex = i
                End If
            Next i

            ChangedIon = DefaultCation

            If NumberOfCations = 0 Then
               ChangedIon.SeparationFactor = 1#
               frmAddComponent!txtAlphaValue.Enabled = False
            Else
               frmAddComponent!txtAlphaValue.Enabled = True
            End If

            'Generate click events on appropriate units
            ListIndex = frmAddComponent!cboAddIonUnits(0).ListIndex
            frmAddComponent!cboAddIonUnits(0).ListIndex = -1
            frmAddComponent!cboAddIonUnits(0).ListIndex = ListIndex

            ListIndex = frmAddComponent!cboAddIonUnits(1).ListIndex
            frmAddComponent!cboAddIonUnits(1).ListIndex = -1
            frmAddComponent!cboAddIonUnits(1).ListIndex = ListIndex

            AddingCation = True
            AddingAnion = False
            EditingCation = False
            EditingAnion = False
            NumberOfIons = NumberOfCations + 1
            NumberOfIonToEdit = NumberOfIons
            If NumberOfIons > 1 Then
               frmAddComponent!cmdViewSeparationFactors.Enabled = True
            Else
               frmAddComponent!cmdViewSeparationFactors.Enabled = False
            End If

            Cation(NumberOfIons).Name = "Cation"
            Cation(NumberOfIons).SeparationFactor = DefaultCation.SeparationFactor

            For i = 1 To NumberOfIons
                OneDimSeparationFactors(i) = Cation(i).SeparationFactor
            Next i

            frmAddComponent.Show 1

            CationSeparationFactorInput.Row = SeparationFactorInput.Row
            CationSeparationFactorInput.Value = SeparationFactorInput.Value

            AddingCation = False

       Case 1   'Remove Ion

       Case 2   'Add Anion

            If NumberOfAnions = MAX_CHEMICAL Then
               msg = "Adding another Anion would exceed the maximum number of "
               msg = msg & "anions allowed for a simulation.  If you would like to "
               msg = msg + "add another anion, you must remove one of the "
               msg = msg + "current anions first."
               MsgBox msg, MB_ICONSTOP, "Too Many Anions"
            End If

            frmAddComponent.Caption = "Add Anion"
            frmAddComponent!lblValenceSign.Caption = "-"
            frmAddComponent!txtAddIon(0).Text = "Anion"
            frmAddComponent!txtAddIon(1).Text = Trim$(Str$(DefaultAnion.MolecularWeight))
            frmAddComponent!txtAddIon(2).Text = Trim$(Str$(DefaultAnion.InitialConcentration))
            frmAddComponent!lblValence.Caption = Trim$(Str$(CInt(DefaultAnion.Valence)))
            frmAddComponent!txtAlphaValue.Text = Trim$(Str$(DefaultAnion.SeparationFactor))

            SeparationFactorInput.Row = AnionSeparationFactorInput.Row
            SeparationFactorInput.Value = AnionSeparationFactorInput.Value
            If SeparationFactorInput.Row = True Then
               frmAddComponent!lblAlpha(1).Caption = frmAddComponent!txtAddIon(0).Text
            Else
               frmAddComponent!lblAlpha(2).Caption = frmAddComponent!txtAddIon(0).Text
            End If

            If NumberOfAnions > 0 Then
               If SeparationFactorInput.Row = True Then
                  frmAddComponent!lblAlpha(2).Caption = Anion(SeparationFactorInput.Value - 10).Name
               Else
                  frmAddComponent!lblAlpha(1).Caption = Anion(SeparationFactorInput.Value).Name
               End If
            Else
               If SeparationFactorInput.Row = True Then
                  frmAddComponent!lblAlpha(2).Caption = frmAddComponent!lblAlpha(1).Caption
               Else
                  frmAddComponent!lblAlpha(1).Caption = frmAddComponent!lblAlpha(2).Caption
               End If
            End If

            If Trim$(frmAddComponent!lblAlpha(1).Caption) = Trim$(frmAddComponent!lblAlpha(2).Caption) Then
               frmAddComponent!txtAlphaValue.Enabled = False
            Else
               frmAddComponent!txtAlphaValue.Enabled = True
            End If

            For i = 0 To frmAddComponent!cboCation.ListCount - 1
                If DefaultAnion.Kinetic.NernstHaskellCation.Ion_Name = frmAddComponent!cboCation.List(i) Then
                   frmAddComponent!cboCation.ListIndex = i
                End If
            Next i
            For i = 0 To frmAddComponent!cboAnion.ListCount - 1
                If DefaultAnion.Kinetic.NernstHaskellAnion.Ion_Name = frmAddComponent!cboAnion.List(i) Then
                   frmAddComponent!cboAnion.ListIndex = i
                End If
            Next i

            ChangedIon = DefaultAnion
            
            If NumberOfAnions = 0 Then
               ChangedIon.SeparationFactor = 1#
               frmAddComponent!txtAlphaValue.Enabled = False
            Else
               frmAddComponent!txtAlphaValue.Enabled = True
            End If

            'Generate click events on appropriate units
            ListIndex = frmAddComponent!cboAddIonUnits(0).ListIndex
            frmAddComponent!cboAddIonUnits(0).ListIndex = -1
            frmAddComponent!cboAddIonUnits(0).ListIndex = ListIndex

            ListIndex = frmAddComponent!cboAddIonUnits(1).ListIndex
            frmAddComponent!cboAddIonUnits(1).ListIndex = -1
            frmAddComponent!cboAddIonUnits(1).ListIndex = ListIndex

            AddingCation = False
            AddingAnion = True
            EditingCation = False
            EditingAnion = False
            NumberOfIons = NumberOfAnions + 1
            NumberOfIonToEdit = NumberOfIons
            If NumberOfIons > 1 Then
               frmAddComponent!cmdViewSeparationFactors.Enabled = True
            Else
               frmAddComponent!cmdViewSeparationFactors.Enabled = False
            End If

            Anion(NumberOfIons).Name = "Anion"
            Anion(NumberOfIons).SeparationFactor = DefaultAnion.SeparationFactor

            For i = 1 To NumberOfIons
                OneDimSeparationFactors(i) = Anion(i).SeparationFactor
            Next i

            frmAddComponent.Show 1

            AnionSeparationFactorInput.Row = SeparationFactorInput.Row
            AnionSeparationFactorInput.Value = SeparationFactorInput.Value

            AddingAnion = False

    End Select

End Sub

Sub cmdEditProperties_Click ()
    Dim i As Integer
    Dim FoundCation As Integer, FoundAnion As Integer
    Dim ListIndex As Integer


    FoundCation = False
    FoundAnion = False
    NumberOfIonToEdit = 0
    For i = 1 To NumberOfCations
        If Trim$(Cation(i).Name) = Trim$(cboIons(2).List(cboIons(2).ListIndex)) Then
           NumberOfIonToEdit = i
           FoundCation = True
           Exit For
        End If
    Next i
    If Not FoundCation Then
       For i = 1 To NumberOfAnions
           If Trim$(Anion(i).Name) = Trim$(cboIons(2).List(cboIons(2).ListIndex)) Then
              NumberOfIonToEdit = i
              FoundAnion = True
              Exit For
           End If
       Next i
    End If

    If FoundCation = True Then
       frmAddComponent.Caption = "Edit Cation"
       frmAddComponent!lblValenceSign.Caption = "+"
       frmAddComponent!txtAddIon(0).Text = Trim$(Cation(NumberOfIonToEdit).Name)
       frmAddComponent!txtAddIon(0).Enabled = False
       frmAddComponent!txtAddIon(1).Text = Trim$(Str$(Cation(NumberOfIonToEdit).MolecularWeight))
       frmAddComponent!txtAddIon(2).Text = Trim$(Str$(Cation(NumberOfIonToEdit).InitialConcentration))
       frmAddComponent!lblValence.Caption = Trim$(Str$(CInt(Cation(NumberOfIonToEdit).Valence)))
       frmAddComponent!txtAlphaValue.Text = Trim$(Str$(Cation(NumberOfIonToEdit).SeparationFactor))

            SeparationFactorInput.Row = CationSeparationFactorInput.Row
            SeparationFactorInput.Value = CationSeparationFactorInput.Value
            If SeparationFactorInput.Row = True Then
               frmAddComponent!lblAlpha(1).Caption = frmAddComponent!txtAddIon(0).Text
            Else
               frmAddComponent!lblAlpha(2).Caption = frmAddComponent!txtAddIon(0).Text
            End If

            If NumberOfCations > 0 Then
               If SeparationFactorInput.Row = True Then
                  frmAddComponent!lblAlpha(2).Caption = Cation(SeparationFactorInput.Value - 10).Name
               Else
                  frmAddComponent!lblAlpha(1).Caption = Cation(SeparationFactorInput.Value).Name
               End If
            Else
               If SeparationFactorInput.Row = True Then
                  frmAddComponent!lblAlpha(2).Caption = frmAddComponent!lblAlpha(1).Caption
               Else
                  frmAddComponent!lblAlpha(1).Caption = frmAddComponent!lblAlpha(2).Caption
               End If
            End If

            If Trim$(frmAddComponent!lblAlpha(1).Caption) = Trim$(frmAddComponent!lblAlpha(2).Caption) Then
               frmAddComponent!txtAlphaValue.Enabled = False
            Else
               frmAddComponent!txtAlphaValue.Enabled = True
            End If

       For i = 0 To frmAddComponent!cboAnion.ListCount - 1
           If Cation(NumberOfIonToEdit).Kinetic.NernstHaskellAnion.Ion_Name = frmAddComponent!cboAnion.List(i) Then
              frmAddComponent!cboAnion.ListIndex = i
           End If
       Next i
       For i = 0 To frmAddComponent!cboCation.ListCount - 1
           If Cation(NumberOfIonToEdit).Kinetic.NernstHaskellCation.Ion_Name = frmAddComponent!cboCation.List(i) Then
              frmAddComponent!cboCation.ListIndex = i
           End If
       Next i

       ChangedIon = Cation(NumberOfIonToEdit)

       'Generate click events on appropriate units
       ListIndex = frmAddComponent!cboAddIonUnits(0).ListIndex
       frmAddComponent!cboAddIonUnits(0).ListIndex = -1
       frmAddComponent!cboAddIonUnits(0).ListIndex = ListIndex

       ListIndex = frmAddComponent!cboAddIonUnits(1).ListIndex
       frmAddComponent!cboAddIonUnits(1).ListIndex = -1
       frmAddComponent!cboAddIonUnits(1).ListIndex = ListIndex

       AddingCation = False
       AddingAnion = False
       EditingCation = True
       EditingAnion = False
       NumberOfIons = NumberOfCations
            If NumberOfIons > 1 Then
               frmAddComponent!cmdViewSeparationFactors.Enabled = True
            Else
               frmAddComponent!cmdViewSeparationFactors.Enabled = False
            End If

            For i = 1 To NumberOfIons
                OneDimSeparationFactors(i) = Cation(i).SeparationFactor
            Next i

       frmAddComponent.Show 1

            CationSeparationFactorInput.Row = SeparationFactorInput.Row
            CationSeparationFactorInput.Value = SeparationFactorInput.Value

       EditingCation = False

    ElseIf FoundAnion = True Then
       frmAddComponent.Caption = "Edit Anion"
       frmAddComponent!lblValenceSign.Caption = "-"
       frmAddComponent!txtAddIon(0).Text = Trim$(Anion(NumberOfIonToEdit).Name)
       frmAddComponent!txtAddIon(0).Enabled = False
       frmAddComponent!txtAddIon(1).Text = Trim$(Str$(Anion(NumberOfIonToEdit).MolecularWeight))
       frmAddComponent!txtAddIon(2).Text = Trim$(Str$(Anion(NumberOfIonToEdit).InitialConcentration))
       frmAddComponent!lblValence.Caption = Trim$(Str$(CInt(Anion(NumberOfIonToEdit).Valence)))
       frmAddComponent!txtAlphaValue.Text = Trim$(Str$(Anion(NumberOfIonToEdit).SeparationFactor))

            SeparationFactorInput.Row = AnionSeparationFactorInput.Row
            SeparationFactorInput.Value = AnionSeparationFactorInput.Value
            If SeparationFactorInput.Row = True Then
               frmAddComponent!lblAlpha(1).Caption = frmAddComponent!txtAddIon(0).Text
            Else
               frmAddComponent!lblAlpha(2).Caption = frmAddComponent!txtAddIon(0).Text
            End If

            If NumberOfAnions > 0 Then
               If SeparationFactorInput.Row = True Then
                  frmAddComponent!lblAlpha(2).Caption = Anion(SeparationFactorInput.Value - 10).Name
               Else
                  frmAddComponent!lblAlpha(1).Caption = Anion(SeparationFactorInput.Value).Name
               End If
            Else
               If SeparationFactorInput.Row = True Then
                  frmAddComponent!lblAlpha(2).Caption = frmAddComponent!lblAlpha(1).Caption
               Else
                  frmAddComponent!lblAlpha(1).Caption = frmAddComponent!lblAlpha(2).Caption
               End If
            End If

            If Trim$(frmAddComponent!lblAlpha(1).Caption) = Trim$(frmAddComponent!lblAlpha(2).Caption) Then
               frmAddComponent!txtAlphaValue.Enabled = False
            Else
               frmAddComponent!txtAlphaValue.Enabled = True
            End If

       For i = 0 To frmAddComponent!cboAnion.ListCount - 1
           If Anion(NumberOfIonToEdit).Kinetic.NernstHaskellAnion.Ion_Name = frmAddComponent!cboAnion.List(i) Then
              frmAddComponent!cboAnion.ListIndex = i
           End If
       Next i
       For i = 0 To frmAddComponent!cboCation.ListCount - 1
           If Anion(NumberOfIonToEdit).Kinetic.NernstHaskellCation.Ion_Name = frmAddComponent!cboCation.List(i) Then
              frmAddComponent!cboCation.ListIndex = i
           End If
       Next i

       ChangedIon = Anion(NumberOfIonToEdit)
       
       'Generate click events on appropriate units
       ListIndex = frmAddComponent!cboAddIonUnits(0).ListIndex
       frmAddComponent!cboAddIonUnits(0).ListIndex = -1
       frmAddComponent!cboAddIonUnits(0).ListIndex = ListIndex

       ListIndex = frmAddComponent!cboAddIonUnits(1).ListIndex
       frmAddComponent!cboAddIonUnits(1).ListIndex = -1
       frmAddComponent!cboAddIonUnits(1).ListIndex = ListIndex

       AddingCation = False
       AddingAnion = False
       EditingCation = False
       EditingAnion = True
       NumberOfIons = NumberOfAnions
            If NumberOfIons > 1 Then
               frmAddComponent!cmdViewSeparationFactors.Enabled = True
            Else
               frmAddComponent!cmdViewSeparationFactors.Enabled = False
            End If

            For i = 1 To NumberOfIons
                OneDimSeparationFactors(i) = Anion(i).SeparationFactor
            Next i

       frmAddComponent.Show 1

            AnionSeparationFactorInput.Row = SeparationFactorInput.Row
            AnionSeparationFactorInput.Value = SeparationFactorInput.Value

       EditingAnion = False

    End If

    frmAddComponent!txtAddIon(0).Enabled = True

End Sub

Sub cmdInputKineticParameters_Click ()
    Dim i As Integer, ListIndex As Integer

'       frmInputKineticParameters!cboIon.ListIndex = -1
'       frmInputKineticParameters!cboIon.ListIndex = cboKinDimComponent.ListIndex

    If cmdAddDeleteIons(0).Enabled And cmdAddDeleteIons(2).Enabled Then   'Both Cations and Anions can be modified
       EditingCation = True
       EditingAnion = True
       For i = 1 To NumberOfCations
           OldCationKineticParameters(i) = Cation(i).Kinetic
       Next i

       For i = 1 To NumberOfAnions
           OldAnionKineticParameters(i) = Anion(i).Kinetic
       Next i
    ElseIf cmdAddDeleteIons(0).Enabled Then   'Only cations can be modified
       EditingCation = True
       EditingAnion = False

       For i = 1 To NumberOfCations
           OldCationKineticParameters(i) = Cation(i).Kinetic
       Next i

    ElseIf cmdAddDeleteIons(2).Enabled Then   'Only anions can be modified
       EditingCation = True
       EditingAnion = False

       For i = 1 To NumberOfAnions
           OldAnionKineticParameters(i) = Anion(i).Kinetic
       Next i

    End If

    ViewingKineticParametersForm = True

    ListIndex = frmInputKineticParameters!cboIon.ListIndex
    frmInputKineticParameters!cboIon.ListIndex = -1
    frmInputKineticParameters!cboIon.ListIndex = ListIndex

    frmInputKineticParameters.Show 1  'Modal
    ViewingKineticParametersForm = False

    EditingAnion = False
    EditingCation = False

End Sub

Sub Form_Load ()

    Screen.MousePointer = 11

    Application_Name = "Simulation Program for Ion Exchange"

    'Set paths for program
   IonExchangePath = CurDir$
'    IonExchangePath = "d:\nasa\pfpdm\vbasic"
'    IonExchangePath = "c:\nasa\vbasic"
    ChDrive IonExchangePath
    ChDir IonExchangePath
    SaveAndLoadPath = CurDir$


    frmIonExchangeMain.WindowState = 0
    frmIonExchangeMain.Width = SCREEN_WIDTH_STANDARD
    frmIonExchangeMain.Height = SCREEN_HEIGHT_STANDARD

    'Center the form on the screen
    If WindowState = 0 Then
       'don't attempt if screen Minimized or Maximized
       Move (Screen.Width - frmIonExchangeMain.Width) / 2, (Screen.Height - frmIonExchangeMain.Height) / 2
    End If

    frmIonExchangeMain.Caption = "Ion Exchange Simulation Software - untitled.iex"
    OldFileName$ = "untitled.iex"
    FileName$ = ""

    Call LoadUnitsOperatingConditions
    Call LoadUnitsBedData
    Call LoadUnitsAdsorbentProperties
    Call LoadUnitsAddIon
    Call LoadUnitsKineticParameters
    Call LoadUnitsTimeParameters
    OKToGetCationDimensionless = False
    OKToGetAnionDimensionless = False
    
'    Call InitializeAvailableIons
    Call InitializeIonExchangeParameters
    Call LoadNernstHaskellDatabases
    Call InitializeDefaultIonProperties
    Call InitializeSeparationFactorInfo
    Call InitializeTimeAndCollocationInfo
    ViewingKineticParametersForm = False
    NumberOfCations = 0
    NumberOfAnions = 0
    VarInfluentFileCation = "NONE"
    VarInfluentFileAnion = "NONE"
    fraKineticDimensionless.Enabled = False
    ClickGeneratedFromcboIon = False
    NumSelectedCations = 0
    NumSelectedAnions = 0

    mnuRun(0).Enabled = False
    mnuResults(0).Enabled = False
    mnuResults(1).Enabled = False

    'Load forms
    Load frmInputKineticParameters
    Load frmAddComponent
    Load frmSeparationFactors
    Load frmOptionsInputParameters

    Unload frmFirst
    Screen.MousePointer = 1

    frmIonExchangeMain.Show


End Sub

Sub Form_Unload (Cancel As Integer)
    Unload frmInputKineticParameters
    Unload frmAddComponent
    Unload frmSeparationFactors
    End
End Sub

Sub GetAndDisplayEBCT ()
    Dim CurrentUnits As Integer, ValueToDisplay As Double

    Bed.EBCT.Value = EBCT(Bed.Length, Bed.Diameter, Bed.Flowrate.Value)
    CurrentUnits = cboBedDataUnits(4).ListIndex
    If CurrentUnits = 0 Then
       ValueToDisplay = Bed.EBCT.Value
    Else
       ValueToDisplay = Bed.EBCT.Value * TimeConversionFactor(CurrentUnits)
    End If
    txtBedData(4).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

End Sub

Sub GetAndDisplayFlowrate ()
    Dim CurrentUnits As Integer, ValueToDisplay As Double

    Bed.Flowrate.Value = Flowrate(Bed.Length, Bed.Diameter, Bed.EBCT.Value)
    CurrentUnits = cboBedDataUnits(3).ListIndex
    If CurrentUnits = 0 Then
       ValueToDisplay = Bed.Flowrate.Value
    Else
       ValueToDisplay = Bed.Flowrate.Value * FlowConversionFactor(CurrentUnits)
    End If
    txtBedData(3).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

End Sub

Sub lstIons_Click (Index As Integer)
    Dim i As Integer, j As Integer
    
    Select Case Index
       Case 0   'Cations
          EditingCation = True

          NumSelectedCations = 1
          Cations_Selected(1) = PresaturantCation
  
          For i = 1 To lstIons(0).ListCount
              If lstIons(0).Selected(i - 1) Then
                 For j = 1 To NumberOfCations
                     If Trim$(Cation(j).Name) = Trim$(lstIons(0).List(i - 1)) Then
                        NumSelectedCations = NumSelectedCations + 1
                        Cations_Selected(NumSelectedCations) = j
                        Exit For
                     End If
                 Next j
              End If
          Next i

          If NumSelectedCations < 2 Then
             mnuRun(0).Enabled = False
             mnuOptions(4).Enabled = False
             cboKinDimComponent.Clear
             cboKinDimComponent.Enabled = False
             For i = 0 To 8
                 lblKineticDimensionlessValue(i).Caption = ""
             Next i
          Else
             mnuRun(0).Enabled = True
             mnuOptions(4).Enabled = True
             fraKineticDimensionless.Enabled = True
             cboKinDimComponent.Clear
             For i = 1 To NumSelectedCations
                 cboKinDimComponent.AddItem Trim$(Cation(Cations_Selected(i)).Name)
             Next i
             cboKinDimComponent.Enabled = True

             Call CalculateSumEquivInitialConc
             For i = 1 To NumSelectedCations
                 NumberOfIonToEdit = Cations_Selected(i)
                 Call CalculateDimensionlessGroups
             Next i

             cboKinDimComponent.ListIndex = -1
             cboKinDimComponent.ListIndex = 0

             'Set Presaturant back to 100 % of initial resin phase concentration
             For i = 1 To MAX_CHEMICAL
                 If i = PresaturantCation Then
                    Resin.PresaturantPercentage(i) = 100#
                 Else
                    Resin.PresaturantPercentage(i) = 0#
                 End If
             Next i

             EditingCation = False
          End If

          Number_Component = NumSelectedCations

       Case 1   'Anions
          EditingAnion = True

          NumSelectedAnions = 1
          Anions_Selected(1) = PresaturantAnion
  
          For i = 1 To lstIons(1).ListCount
              If lstIons(1).Selected(i - 1) Then
                 For j = 1 To NumberOfAnions
                     If Trim$(Anion(j).Name) = Trim$(lstIons(1).List(i - 1)) Then
                        NumSelectedAnions = NumSelectedAnions + 1
                        Anions_Selected(NumSelectedAnions) = j
                        Exit For
                     End If
                 Next j
              End If
          Next i

          If NumSelectedAnions < 2 Then
             mnuRun(0).Enabled = False
             mnuOptions(4).Enabled = False
             cboKinDimComponent.Clear
             cboKinDimComponent.Enabled = False
             For i = 0 To 8
                 lblKineticDimensionlessValue(i).Caption = ""
             Next i
          Else
             mnuRun(0).Enabled = True
             mnuOptions(4).Enabled = True
             fraKineticDimensionless.Enabled = True
             cboKinDimComponent.Clear
             For i = 1 To NumSelectedAnions
                 cboKinDimComponent.AddItem Trim$(Anion(Anions_Selected(i)).Name)
             Next i
             cboKinDimComponent.Enabled = True

             Call CalculateSumEquivInitialConc
             For i = 1 To NumSelectedAnions
                 NumberOfIonToEdit = Anions_Selected(i)
                 Call CalculateDimensionlessGroups
             Next i

             cboKinDimComponent.ListIndex = -1
             cboKinDimComponent.ListIndex = 0

             'Set Presaturant back to 100 % of initial resin phase concentration
             For i = 1 To MAX_CHEMICAL
                 If i = PresaturantAnion Then
                    Resin.PresaturantPercentage(i) = 100#
                 Else
                    Resin.PresaturantPercentage(i) = 0#
                 End If
             Next i

             EditingAnion = False
          End If

          Number_Component = NumSelectedAnions

    End Select

End Sub

Sub mnuFile_Click (Index As Integer)

    Select Case Index
       Case 0   'New

       Case 1   'Open
            ChDrive SaveAndLoadPath
            ChDir SaveAndLoadPath
            Call LoadIonExchange
            SaveAndLoadPath = CurDir$
            txtOperatingConditions(0).SetFocus
            ChDrive IonExchangePath
            ChDir IonExchangePath

       Case 2   'Save
          ChDrive SaveAndLoadPath
          ChDir SaveAndLoadPath
          Call SaveIonExchange
          SaveAndLoadPath = CurDir$
          ChDir IonExchangePath
          ChDrive IonExchangePath

       Case 3   'Save As
          ChDrive SaveAndLoadPath
          ChDir SaveAndLoadPath
          OldFileName$ = FileName$
          FileName$ = ""
          Call SaveIonExchange
          SaveAndLoadPath = CurDir$
          ChDir IonExchangePath
          ChDrive IonExchangePath

       Case 5   'Select Printer
            On Error GoTo PrinterError
               CMDialog1.Flags = PD_PRINTSETUP
               CMDialog1.Action = 5
               Exit Sub
PrinterError:
               Resume ExitSelectPrint:

ExitSelectPrint:

       Case 6   'Print

       Case 8   'Exit
            Unload frmIonExchangeMain

    End Select

End Sub

Sub mnuFilePrint_Click (Index As Integer)

    Select Case Index
       Case 0   'Print to printer
          Call PrintIonExchange
       Case 1   'Print to file
          Call PrintIonExchangeToFile
    End Select

End Sub

Sub mnuOptions_Click (Index As Integer)
    Dim i As Integer, j As Integer
    
    Select Case Index
       Case 0   'Set Variable Influent Concentrations
          If Cations.Available And Anions.Available Then

          ElseIf Cations.Available Then
             For i = 1 To NumberOfCations
                 Ion(i) = Cation(i)
             Next i
             Total_NumberOfComponents = NumberOfCations
             Call ReadVarInfluentConcs

          ElseIf Anions.Available Then
             For i = 1 To NumberOfAnions
                 Ion(i) = Anion(i)
             Next i
             Total_NumberOfComponents = NumberOfAnions


          End If
             
          frmConcentrations.Show 1
       Case 1   'Set Number Of Beds
          frmOptionsInputParameters.Show 1
       Case 2   'Set Collocation Points
          frmOptionsInputParameters.Show 1
       Case 3   'Set Time Parameters
          frmOptionsInputParameters.Show 1
       Case 4   'Initial Resin Phase Concentrations
          frmResinPresaturantConditions.Show 1
    End Select

End Sub

Sub mnuResults_Click (Index As Integer)
    Dim i As Integer, j As Integer

    Select Case Index
       Case 0   'PFPDM results
          frmBreak.Show 1
       Case 1   'Compare to Data
          frmPlantData.Show 1
    End Select
End Sub

Sub mnuRun_Click (Index As Integer)
    Dim i As Integer, j As Integer

    Select Case Index
       Case 0   'PFPDM
'          Call GetSelectedComponents(0)

          If Cations.Available And Anions.Available Then

          ElseIf Cations.Available Then
             NumSelectedComponents_PFPDM = NumSelectedCations
             'Place Presaturant Ion in Last Element of Array
             For i = 2 To NumSelectedCations
                 Component_Index_PFPDM(i - 1) = Cations_Selected(i)
             Next i
             Component_Index_PFPDM(NumSelectedCations) = Cations_Selected(1)
             For i = 1 To NumberOfCations
                 Ion(i) = Cation(i)
             Next i

             'Determine Alpha_Input Array to send to PFPDM
             For i = 1 To NumberOfCations
                 OneDimSeparationFactors(i) = Cation(i).SeparationFactor
             Next i
             SeparationFactorInput.Row = CationSeparationFactorInput.Row
             SeparationFactorInput.Value = CationSeparationFactorInput.Value
             NumberOfIons = NumberOfCations
             Call CalculateSeparationFactors
             For i = 1 To NumSelectedCations
                 j = Cations_Selected(i)
                 AlphaInput(i) = TwoDimSeparationFactors(j, 1)
             Next i

             Call Call_PFPDM

          ElseIf Anions.Available Then

             NumSelectedComponents_PFPDM = NumSelectedAnions
             'Place Presaturant Ion in Last Element of Array
             For i = 2 To NumSelectedAnions
                 Component_Index_PFPDM(i - 1) = Anions_Selected(i)
             Next i
             Component_Index_PFPDM(NumSelectedAnions) = Anions_Selected(1)
             For i = 1 To NumberOfAnions
                 Ion(i) = Anion(i)
             Next i

             'Determine Alpha_Input Array to send to PFPDM
             For i = 1 To NumberOfAnions
                 OneDimSeparationFactors(i) = Anion(i).SeparationFactor
             Next i
             SeparationFactorInput.Row = AnionSeparationFactorInput.Row
             SeparationFactorInput.Value = AnionSeparationFactorInput.Value
             NumberOfIons = NumberOfAnions
             Call CalculateSeparationFactors
             For i = 1 To NumSelectedAnions
                 j = Anions_Selected(i)
                 AlphaInput(i) = TwoDimSeparationFactors(j, 1)
             Next i

             Call Call_PFPDM

          End If

    End Select

End Sub

Sub txtAdsorbentProperties_GotFocus (Index As Integer)
    Call TextGetFocus(txtAdsorbentProperties(Index), Temp_Text)
End Sub

Sub txtAdsorbentProperties_KeyPress (Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = 13 Then
       KeyAscii = 0
       SendKeys "{Tab}"
       Exit Sub
    End If

    If Index = 0 Then Exit Sub

    Call NumberCheck(KeyAscii)

End Sub

Sub txtAdsorbentProperties_LostFocus (Index As Integer)
    Dim ValueChanged As Integer, NewValue As Double
    Dim msg As String, CurrentUnits As Integer
    Dim OldValue As Double, ValueToDisplay As Double
    Dim i As Integer

    Call TextHandleError(IsError, txtAdsorbentProperties(Index), Temp_Text)

    If Not IsError Then
       NewValue = CDbl(txtAdsorbentProperties(Index).Text)
       'Convert NewValue to Standard Units if Necessary
       Select Case Index
          Case 1   'Apparent Density
               OldValue = Resin.ApparentDensity
               CurrentUnits = cboAdsorbentPropertyUnits(1).ListIndex
               If CurrentUnits <> 0 Then
                  NewValue = NewValue / DensityConversionFactor(CurrentUnits)
               End If
          Case 2   'Particle Radius
               OldValue = Resin.ParticleRadius
               CurrentUnits = cboAdsorbentPropertyUnits(2).ListIndex
               If CurrentUnits <> 0 Then
                  NewValue = NewValue / LengthConversionFactor(CurrentUnits)
               End If
          Case 3   'Particle Porosity
               OldValue = Resin.ParticlePorosity
          Case 4   'Tortuosity
               OldValue = Resin.Tortuosity
          Case 5   'Total Resin Capacity
               OldValue = Resin.TotalCapacity
               CurrentUnits = cboAdsorbentPropertyUnits(5).ListIndex
               If CurrentUnits <> 0 Then
                  NewValue = NewValue / ResinCapacityConversionFactor(CurrentUnits)
               End If
       End Select

       Select Case Index
          Case 1    'Apparent Density
             Call NumberChanged(ValueChanged, OldValue, NewValue)
             If ValueChanged Then
                If HaveValue(NewValue) Then
                   Resin.ApparentDensity = NewValue
                   
                   Call CalculateBedPorosity
                   Call CalculateEffectiveContactTime
                   Call CalculateInterstitialVelocity
                   
                   Call UpdateKineticParametersAllIons
                   Call UpdateDimensionlessGroupAllIons

                Else
                   txtAdsorbentProperties(1).Text = Temp_Text
                   txtAdsorbentProperties(1).SetFocus
                   Exit Sub
                End If
             End If

          Case 2    'Particle Radius
             Call NumberChanged(ValueChanged, OldValue, NewValue)
             If ValueChanged Then
                If HaveValue(NewValue) Then
                   Resin.ParticleRadius = NewValue
                   Call CalculateParticleDiameter

                   Call UpdateKineticParametersAllIons
                   Call UpdateDimensionlessGroupAllIons

                Else
                   txtAdsorbentProperties(2).Text = Temp_Text
                   txtAdsorbentProperties(2).SetFocus
                   Exit Sub
                End If
             End If

          Case 3    'Particle Porosity
             Call NumberChanged(ValueChanged, OldValue, NewValue)
             If ValueChanged Then
                If HaveValue(NewValue) Then
                   Resin.ParticlePorosity = NewValue

                   Call UpdateDimensionlessGroupAllIons

                Else
                   txtAdsorbentProperties(3).Text = Temp_Text
                   txtAdsorbentProperties(3).SetFocus
                   Exit Sub
                End If
             End If
             
          Case 4    'Tortuosity
             Call NumberChanged(ValueChanged, OldValue, NewValue)
             If ValueChanged Then
                If HaveValue(NewValue) Then
                   Resin.Tortuosity = NewValue

                   Call UpdateKineticParametersAllIons
                   Call UpdateDimensionlessGroupAllIons

                Else
                   txtAdsorbentProperties(4).Text = Temp_Text
                   txtAdsorbentProperties(4).SetFocus
                   Exit Sub
                End If
             End If

          Case 5    'Resin Capacity
             Call NumberChanged(ValueChanged, OldValue, NewValue)
             If ValueChanged Then
                If HaveValue(NewValue) Then
                   Resin.TotalCapacity = NewValue

                   Call UpdateDimensionlessGroupAllIons

                Else
                   txtAdsorbentProperties(5).Text = Temp_Text
                   txtAdsorbentProperties(5).SetFocus
                   Exit Sub
                End If
             End If


       End Select

    End If

End Sub

Sub txtBedData_GotFocus (Index As Integer)
    Call TextGetFocus(txtBedData(Index), Temp_Text)
End Sub

Sub txtBedData_KeyPress (Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
       KeyAscii = 0
       SendKeys "{Tab}"
       Exit Sub
    End If

    Call NumberCheck(KeyAscii)

End Sub

Sub txtBedData_LostFocus (Index As Integer)
    Dim ValueChanged As Integer, NewValue As Double
    Dim msg As String, CurrentUnits As Integer
    Dim OldValue As Double, ValueToDisplay As Double
    Dim i As Integer

    Call TextHandleError(IsError, txtBedData(Index), Temp_Text)

    If Not IsError Then
       NewValue = CDbl(txtBedData(Index).Text)
       'Convert NewValue to Standard Units if Necessary
       Select Case Index
          Case 0   'Bed Length
               OldValue = Bed.Length
               CurrentUnits = cboBedDataUnits(0).ListIndex
               If CurrentUnits <> 0 Then
                  NewValue = NewValue / LengthConversionFactor(CurrentUnits)
               End If
          Case 1   'Bed Diameter
               OldValue = Bed.Diameter
               CurrentUnits = cboBedDataUnits(1).ListIndex
               If CurrentUnits <> 0 Then
                  NewValue = NewValue / LengthConversionFactor(CurrentUnits)
               End If
          Case 2   'Bed Mass
               OldValue = Bed.Weight
               CurrentUnits = cboBedDataUnits(2).ListIndex
               If CurrentUnits <> 0 Then
                  NewValue = NewValue / MassConversionFactor(CurrentUnits)
               End If
          Case 3   'Flowrate
               OldValue = Bed.Flowrate.Value
               CurrentUnits = cboBedDataUnits(3).ListIndex
               If CurrentUnits <> 0 Then
                  NewValue = NewValue / FlowConversionFactor(CurrentUnits)
               End If
          Case 4   'EBCT
               OldValue = Bed.EBCT.Value
               CurrentUnits = cboBedDataUnits(4).ListIndex
               If CurrentUnits <> 0 Then
                  NewValue = NewValue / TimeConversionFactor(CurrentUnits)
               End If
       End Select

       Select Case Index
          Case 0    'Bed Length
             Call NumberChanged(ValueChanged, OldValue, NewValue)
             If ValueChanged Then
                If HaveValue(NewValue) Then
                   Bed.Length = NewValue
                   If Bed.Flowrate.UserInput Then
                      Call GetAndDisplayEBCT
                   Else
                      Call GetAndDisplayFlowrate
                   End If

                   Call CalculateBedVolume
                   Call CalculateBedDensity
                   Call CalculateBedPorosity
                   Call CalculateEffectiveContactTime
                   Call CalculateInterstitialVelocity
                   
                   Call UpdateKineticParametersAllIons
                   Call UpdateDimensionlessGroupAllIons

                Else
                   txtBedData(0).Text = Temp_Text
                   txtBedData(0).SetFocus
                   Exit Sub
                End If
             End If

          Case 1    'Bed Diameter
             Call NumberChanged(ValueChanged, OldValue, NewValue)
             If ValueChanged Then
                If HaveValue(NewValue) Then
                   Bed.Diameter = NewValue
                   If Bed.Flowrate.UserInput Then
                      Call GetAndDisplayEBCT
                   Else
                      Call GetAndDisplayFlowrate
                   End If

                   Call CalculateBedArea
                   Call CalculateBedVolume
                   Call CalculateBedDensity
                   Call CalculateBedPorosity
                   Call CalculateEffectiveContactTime
                   Call CalculateSuperficialVelocity
                   Call CalculateInterstitialVelocity

                   Call UpdateKineticParametersAllIons
                   Call UpdateDimensionlessGroupAllIons

                Else
                   txtBedData(1).Text = Temp_Text
                   txtBedData(1).SetFocus
                   Exit Sub
                End If
             End If

          Case 2    'Bed Mass
             Call NumberChanged(ValueChanged, OldValue, NewValue)
             If ValueChanged Then
                If HaveValue(NewValue) Then
                   Bed.Weight = NewValue
                   Call CalculateBedDensity
                   Call CalculateBedPorosity
                   Call CalculateInterstitialVelocity

                   Call CalculateKineticParameters
                   Call UpdateDimensionlessGroupAllIons

                Else
                   txtBedData(2).Text = Temp_Text
                   txtBedData(2).SetFocus
                   Exit Sub
                End If
             End If

             
          Case 3    'Flowrate
             Call NumberChanged(ValueChanged, OldValue, NewValue)
             If ValueChanged Then
                If HaveValue(NewValue) Then
                   Bed.Flowrate.Value = NewValue
                   Bed.Flowrate.UserInput = True
                   Bed.EBCT.UserInput = False
                   Call GetAndDisplayEBCT
                   Call CalculateEffectiveContactTime
                   Call CalculateSuperficialVelocity
                   Call CalculateInterstitialVelocity
                   
                   Call UpdateKineticParametersAllIons
                   Call UpdateDimensionlessGroupAllIons
                      
                Else
                   txtBedData(3).Text = Temp_Text
                   txtBedData(3).SetFocus
                   Exit Sub
                End If
             End If

          Case 4    'EBCT
             Call NumberChanged(ValueChanged, OldValue, NewValue)
             If ValueChanged Then
                If HaveValue(NewValue) Then
                   Bed.EBCT.Value = NewValue
                   Bed.EBCT.UserInput = True
                   Bed.Flowrate.UserInput = False
                   Call GetAndDisplayFlowrate
                   Call CalculateEffectiveContactTime
                   Call CalculateSuperficialVelocity
                   Call CalculateInterstitialVelocity
                  
                   Call UpdateKineticParametersAllIons
                   Call UpdateDimensionlessGroupAllIons

                Else
                   txtBedData(4).Text = Temp_Text
                   txtBedData(4).SetFocus
                   Exit Sub
                End If
             End If


       End Select

    End If

End Sub

Sub txtOperatingConditions_GotFocus (Index As Integer)
    Call TextGetFocus(txtOperatingConditions(Index), Temp_Text)
End Sub

Sub txtOperatingConditions_KeyPress (Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = 13 Then
       KeyAscii = 0
       SendKeys "{Tab}"
       Exit Sub
    End If

    Call NumberCheck(KeyAscii)

End Sub

Sub txtOperatingConditions_LostFocus (Index As Integer)
    Dim ValueChanged As Integer, NewValue As Double
    Dim msg As String, CurrentUnits As Integer
    Dim OldValue As Double, ValueToDisplay As Double
    Dim i As Integer

    Call TextHandleError(IsError, txtOperatingConditions(Index), Temp_Text)

    If Not IsError Then
       NewValue = CDbl(txtOperatingConditions(Index).Text)
       'Convert NewValue to Standard Units if Necessary
       Select Case Index
          Case 0   'Operating Pressure
               OldValue = Operating.Pressure
               CurrentUnits = cboOperatingConditionsUnits(0).ListIndex
               If CurrentUnits <> 0 Then
                  NewValue = NewValue / PressureConversionFactor(CurrentUnits)
               End If
          Case 1   'Operating Temperature
               OldValue = Operating.Temperature
               CurrentUnits = cboOperatingConditionsUnits(1).ListIndex
               If CurrentUnits <> 0 Then
                  NewValue = ReverseTemperatureConversion(CurrentUnits, NewValue)
               End If
       End Select

       Select Case Index
          Case 0    'Operating Pressure
             Call NumberChanged(ValueChanged, OldValue, NewValue)
             If ValueChanged Then
                If HaveValue(NewValue) Then
                   Operating.Pressure = NewValue
                Else
                   txtOperatingConditions(0).Text = Temp_Text
                   txtOperatingConditions(0).SetFocus
                   Exit Sub
                End If
             End If

          Case 1    'Operating Temperature
             Call NumberChanged(ValueChanged, OldValue, NewValue)
             If ValueChanged Then
                If HaveValue(NewValue) Then
                   Operating.Temperature = NewValue
                   Call CalculateLiquidDensity
                   Call CalculateLiquidViscosity

                Call UpdateKineticParametersAllIons
                Call UpdateDimensionlessGroupAllIons

                Else
                   txtOperatingConditions(1).Text = Temp_Text
                   txtOperatingConditions(1).SetFocus
                   Exit Sub
                End If
             End If

       End Select

    End If

End Sub

