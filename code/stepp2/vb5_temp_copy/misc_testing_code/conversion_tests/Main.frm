VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7575
   ClientLeft      =   2130
   ClientTop       =   1605
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   ScaleHeight     =   7575
   ScaleWidth      =   9975
   Begin VB.ComboBox cboTabSelector 
      Height          =   315
      Left            =   150
      Style           =   2  'Dropdown List
      TabIndex        =   193
      Top             =   60
      Width           =   645
   End
   Begin VB.Frame fraTabHolder 
      Height          =   6315
      Left            =   60
      TabIndex        =   0
      Top             =   450
      Width           =   9855
      Begin VB.Frame fraTab 
         Caption         =   "fraTab(8)"
         Height          =   2775
         Index           =   8
         Left            =   750
         TabIndex        =   1
         Top             =   2430
         Width           =   8985
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Autoignition Temperature"
            Height          =   255
            Index           =   26
            Left            =   120
            TabIndex        =   16
            Top             =   1290
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Flash Point"
            Height          =   255
            Index           =   25
            Left            =   120
            TabIndex        =   15
            Top             =   930
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Upper Flamibility Limit"
            Height          =   255
            Index           =   23
            Left            =   120
            TabIndex        =   14
            Top             =   210
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Lower Flamibility Limit"
            Height          =   255
            Index           =   24
            Left            =   120
            TabIndex        =   13
            Top             =   570
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Heat of Combustion"
            Height          =   255
            Index           =   27
            Left            =   120
            TabIndex        =   12
            Top             =   1650
            Width           =   3735
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   23
            Left            =   6000
            TabIndex        =   11
            Top             =   210
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   24
            Left            =   6000
            TabIndex        =   10
            Top             =   570
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   25
            Left            =   6000
            TabIndex        =   9
            Top             =   930
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   26
            Left            =   6000
            TabIndex        =   8
            Top             =   1290
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   27
            Left            =   6000
            TabIndex        =   7
            Top             =   1650
            Width           =   2175
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   23
            Left            =   4080
            TabIndex        =   6
            Top             =   210
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   24
            Left            =   4080
            TabIndex        =   5
            Top             =   570
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   25
            Left            =   4080
            TabIndex        =   4
            Top             =   930
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   26
            Left            =   4080
            TabIndex        =   3
            Top             =   1290
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   27
            Left            =   4080
            TabIndex        =   2
            Top             =   1650
            Width           =   1695
         End
      End
      Begin VB.Frame fraTab 
         Caption         =   "fraTab(7)"
         Height          =   3495
         Index           =   7
         Left            =   660
         TabIndex        =   17
         Top             =   2160
         Width           =   8985
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Surface Tension as f(T)"
            Height          =   255
            Index           =   18
            Left            =   150
            TabIndex        =   41
            Top             =   1440
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Vapor Viscosity as f(T)"
            Height          =   255
            Index           =   19
            Left            =   150
            TabIndex        =   40
            Top             =   1800
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Liquid Viscosity as f(T)"
            Height          =   255
            Index           =   20
            Left            =   150
            TabIndex        =   39
            Top             =   2160
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Liquid Thermal Conductivity as f(T)"
            Height          =   255
            Index           =   21
            Left            =   150
            TabIndex        =   38
            Top             =   2520
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Diffusivity in Air"
            Height          =   255
            Index           =   16
            Left            =   150
            TabIndex        =   37
            Top             =   720
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Surface Tension @ 298.15 K"
            Height          =   255
            Index           =   17
            Left            =   150
            TabIndex        =   36
            Top             =   1080
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Diffusivity in Water"
            Height          =   255
            Index           =   15
            Left            =   150
            TabIndex        =   35
            Top             =   360
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Vapor Thermal Conductivity as f(T)"
            Height          =   255
            Index           =   22
            Left            =   150
            TabIndex        =   34
            Top             =   2880
            Width           =   3735
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   15
            Left            =   6030
            TabIndex        =   33
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   16
            Left            =   6030
            TabIndex        =   32
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   17
            Left            =   6030
            TabIndex        =   31
            Top             =   1080
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   18
            Left            =   6030
            TabIndex        =   30
            Top             =   1440
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   19
            Left            =   6030
            TabIndex        =   29
            Top             =   1800
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   20
            Left            =   6030
            TabIndex        =   28
            Top             =   2160
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   21
            Left            =   6030
            TabIndex        =   27
            Top             =   2520
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   22
            Left            =   6030
            TabIndex        =   26
            Top             =   2880
            Width           =   2175
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   15
            Left            =   4110
            TabIndex        =   25
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   16
            Left            =   4110
            TabIndex        =   24
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   17
            Left            =   4110
            TabIndex        =   23
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   18
            Left            =   4110
            TabIndex        =   22
            Top             =   1440
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   19
            Left            =   4110
            TabIndex        =   21
            Top             =   1800
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   20
            Left            =   4110
            TabIndex        =   20
            Top             =   2160
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   21
            Left            =   4110
            TabIndex        =   19
            Top             =   2520
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   22
            Left            =   4110
            TabIndex        =   18
            Top             =   2880
            Width           =   1695
         End
      End
      Begin VB.Frame fraTab 
         Caption         =   "fraTab(6)"
         Height          =   2625
         Index           =   6
         Left            =   540
         TabIndex        =   42
         Top             =   1890
         Width           =   9015
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Daphnia Magna, 24h, EC50"
            Height          =   255
            Index           =   57
            Left            =   150
            TabIndex        =   60
            Top             =   300
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Daphnia Magna, 48h, EC50"
            Height          =   255
            Index           =   58
            Left            =   150
            TabIndex        =   59
            Top             =   660
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Daphnia Magna, 24h, LC50"
            Height          =   255
            Index           =   59
            Left            =   150
            TabIndex        =   58
            Top             =   1020
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Daphnia Magna, 48h, LC50"
            Height          =   255
            Index           =   60
            Left            =   150
            TabIndex        =   57
            Top             =   1380
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Mysid, 96, LC50"
            Height          =   255
            Index           =   61
            Left            =   150
            TabIndex        =   56
            Top             =   1740
            Width           =   3735
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   49
            Left            =   4110
            TabIndex        =   55
            Top             =   300
            Width           =   1695
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   49
            Left            =   6030
            TabIndex        =   54
            Top             =   300
            Width           =   2175
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Alternate Species"
            Height          =   255
            Index           =   62
            Left            =   150
            TabIndex        =   53
            Top             =   2100
            Width           =   3735
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   50
            Left            =   6030
            TabIndex        =   52
            Top             =   660
            Width           =   2175
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   50
            Left            =   4110
            TabIndex        =   51
            Top             =   660
            Width           =   1695
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   51
            Left            =   6030
            TabIndex        =   50
            Top             =   1020
            Width           =   2175
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   51
            Left            =   4110
            TabIndex        =   49
            Top             =   1020
            Width           =   1695
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   52
            Left            =   6030
            TabIndex        =   48
            Top             =   1380
            Width           =   2175
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   52
            Left            =   4110
            TabIndex        =   47
            Top             =   1380
            Width           =   1695
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   53
            Left            =   6030
            TabIndex        =   46
            Top             =   1740
            Width           =   2175
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   53
            Left            =   4110
            TabIndex        =   45
            Top             =   1740
            Width           =   1695
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   54
            Left            =   6030
            TabIndex        =   44
            Top             =   2100
            Width           =   2175
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   54
            Left            =   4110
            TabIndex        =   43
            Top             =   2100
            Width           =   1695
         End
      End
      Begin VB.Frame fraTab 
         Caption         =   "fraTab(5)"
         Height          =   3435
         Index           =   5
         Left            =   450
         TabIndex        =   61
         Top             =   1620
         Width           =   9015
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Vapor Pressure as f(T)"
            Height          =   255
            Index           =   6
            Left            =   180
            TabIndex        =   85
            Top             =   2460
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Melting Point"
            Height          =   255
            Index           =   3
            Left            =   180
            TabIndex        =   84
            Top             =   1380
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Normal Boiling Point (NBP)"
            Height          =   255
            Index           =   4
            Left            =   180
            TabIndex        =   83
            Top             =   1740
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Vapor Pressure @ 298.15 K"
            Height          =   255
            Index           =   5
            Left            =   180
            TabIndex        =   82
            Top             =   2100
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Liquid Density as f(T)"
            Height          =   255
            Index           =   2
            Left            =   180
            TabIndex        =   81
            Top             =   1020
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Liquid Density @ 298.15 K"
            Height          =   255
            Index           =   1
            Left            =   180
            TabIndex        =   80
            Top             =   660
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Molecular Weight"
            Height          =   255
            Index           =   0
            Left            =   180
            TabIndex        =   79
            Top             =   300
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Heat of Formation"
            Height          =   255
            Index           =   7
            Left            =   180
            TabIndex        =   78
            Top             =   2820
            Width           =   3735
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   0
            Left            =   6060
            TabIndex        =   77
            Top             =   300
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   1
            Left            =   6060
            TabIndex        =   76
            Top             =   660
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   2
            Left            =   6060
            TabIndex        =   75
            Top             =   1020
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   3
            Left            =   6060
            TabIndex        =   74
            Top             =   1380
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   4
            Left            =   6060
            TabIndex        =   73
            Top             =   1740
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   5
            Left            =   6060
            TabIndex        =   72
            Top             =   2100
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   6
            Left            =   6060
            TabIndex        =   71
            Top             =   2460
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   7
            Left            =   6060
            TabIndex        =   70
            Top             =   2820
            Width           =   2175
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   0
            Left            =   4140
            TabIndex        =   69
            Top             =   300
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   1
            Left            =   4140
            TabIndex        =   68
            Top             =   660
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   2
            Left            =   4140
            TabIndex        =   67
            Top             =   1020
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   3
            Left            =   4140
            TabIndex        =   66
            Top             =   1380
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   4
            Left            =   4140
            TabIndex        =   65
            Top             =   1740
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   5
            Left            =   4140
            TabIndex        =   64
            Top             =   2100
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   6
            Left            =   4140
            TabIndex        =   63
            Top             =   2460
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   7
            Left            =   4140
            TabIndex        =   62
            Top             =   2820
            Width           =   1695
         End
      End
      Begin VB.Frame fraTab 
         Caption         =   "fraTab(4)"
         Height          =   3435
         Index           =   4
         Left            =   360
         TabIndex        =   86
         Top             =   1350
         Width           =   8955
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Liquid Heat Capacity as f(T)"
            Height          =   255
            Index           =   8
            Left            =   210
            TabIndex        =   110
            Top             =   270
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Vapor Heat Capacity as f(T)"
            Height          =   255
            Index           =   9
            Left            =   210
            TabIndex        =   109
            Top             =   630
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Heat of Vaporization @ 298.15 K"
            Height          =   255
            Index           =   10
            Left            =   210
            TabIndex        =   108
            Top             =   990
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Heat of Vaporization  @ NBP"
            Height          =   255
            Index           =   11
            Left            =   210
            TabIndex        =   107
            Top             =   1350
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Heat of Vaporization as f(T)"
            Height          =   255
            Index           =   12
            Left            =   210
            TabIndex        =   106
            Top             =   1710
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Critical Temperature"
            Height          =   255
            Index           =   13
            Left            =   210
            TabIndex        =   105
            Top             =   2070
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Critical Pressure"
            Height          =   255
            Index           =   14
            Left            =   210
            TabIndex        =   104
            Top             =   2430
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Critical Volume"
            Height          =   255
            Index           =   38
            Left            =   210
            TabIndex        =   103
            Top             =   2790
            Width           =   3735
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   8
            Left            =   6000
            TabIndex        =   102
            Top             =   270
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   9
            Left            =   6090
            TabIndex        =   101
            Top             =   630
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   10
            Left            =   6090
            TabIndex        =   100
            Top             =   990
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   11
            Left            =   6090
            TabIndex        =   99
            Top             =   1350
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   12
            Left            =   6090
            TabIndex        =   98
            Top             =   1710
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   13
            Left            =   6090
            TabIndex        =   97
            Top             =   2070
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   14
            Left            =   6090
            TabIndex        =   96
            Top             =   2430
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   38
            Left            =   6090
            TabIndex        =   95
            Top             =   2790
            Width           =   2175
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   8
            Left            =   4170
            TabIndex        =   94
            Top             =   270
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   9
            Left            =   4170
            TabIndex        =   93
            Top             =   630
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   10
            Left            =   4170
            TabIndex        =   92
            Top             =   990
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   11
            Left            =   4170
            TabIndex        =   91
            Top             =   1350
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   12
            Left            =   4170
            TabIndex        =   90
            Top             =   1710
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   13
            Left            =   4170
            TabIndex        =   89
            Top             =   2070
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   14
            Left            =   4170
            TabIndex        =   88
            Top             =   2430
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   38
            Left            =   4170
            TabIndex        =   87
            Top             =   2790
            Width           =   1695
         End
      End
      Begin VB.Frame fraTab 
         Caption         =   "fraTab(3)"
         Height          =   2025
         Index           =   3
         Left            =   270
         TabIndex        =   111
         Top             =   1080
         Width           =   8955
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Combined (C + N) ThOD"
            Height          =   255
            Index           =   29
            Left            =   180
            TabIndex        =   123
            Top             =   750
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Biochemical Oxygen Demand"
            Height          =   255
            Index           =   31
            Left            =   180
            TabIndex        =   122
            Top             =   1470
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Carbonaceous ThOD"
            Height          =   255
            Index           =   28
            Left            =   180
            TabIndex        =   121
            Top             =   390
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Chemical Oxygen Demand"
            Height          =   255
            Index           =   30
            Left            =   180
            TabIndex        =   120
            Top             =   1110
            Width           =   3735
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   28
            Left            =   6060
            TabIndex        =   119
            Top             =   390
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   29
            Left            =   6060
            TabIndex        =   118
            Top             =   750
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   30
            Left            =   6060
            TabIndex        =   117
            Top             =   1110
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   31
            Left            =   6060
            TabIndex        =   116
            Top             =   1470
            Width           =   2175
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   28
            Left            =   4140
            TabIndex        =   115
            Top             =   390
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   29
            Left            =   4140
            TabIndex        =   114
            Top             =   750
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   30
            Left            =   4140
            TabIndex        =   113
            Top             =   1110
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   31
            Left            =   4140
            TabIndex        =   112
            Top             =   1470
            Width           =   1695
         End
      End
      Begin VB.Frame fraTab 
         Caption         =   "fraTab(2)"
         Height          =   3285
         Index           =   2
         Left            =   180
         TabIndex        =   124
         Top             =   810
         Width           =   8925
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Activity Coefficient of Chemical in Water"
            Height          =   255
            Index           =   34
            Left            =   240
            TabIndex        =   148
            Top             =   240
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Activity Coefficient of Water in Chemical"
            Height          =   255
            Index           =   32
            Left            =   240
            TabIndex        =   147
            Top             =   600
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Henry's Constant"
            Height          =   255
            Index           =   33
            Left            =   240
            TabIndex        =   146
            Top             =   960
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Solubility Limit of Chemical in Water"
            Height          =   255
            Index           =   39
            Left            =   240
            TabIndex        =   145
            Top             =   1320
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Solubility Limit of Water in Chemical"
            Height          =   255
            Index           =   40
            Left            =   240
            TabIndex        =   144
            Top             =   1680
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Log Kow"
            Height          =   255
            Index           =   35
            Left            =   240
            TabIndex        =   143
            Top             =   2040
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Log Koc"
            Height          =   255
            Index           =   36
            Left            =   240
            TabIndex        =   142
            Top             =   2400
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Bioconcentration Factor"
            Height          =   255
            Index           =   37
            Left            =   240
            TabIndex        =   141
            Top             =   2760
            Width           =   3735
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   32
            Left            =   6120
            TabIndex        =   140
            Top             =   600
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   34
            Left            =   6120
            TabIndex        =   139
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   33
            Left            =   6120
            TabIndex        =   138
            Top             =   960
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   39
            Left            =   6120
            TabIndex        =   137
            Top             =   1320
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   40
            Left            =   6120
            TabIndex        =   136
            Top             =   1680
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   35
            Left            =   6120
            TabIndex        =   135
            Top             =   2040
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   36
            Left            =   6120
            TabIndex        =   134
            Top             =   2400
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   37
            Left            =   6120
            TabIndex        =   133
            Top             =   2760
            Width           =   2175
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "PNLPropVal"
            Height          =   255
            Index           =   32
            Left            =   4200
            TabIndex        =   132
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   33
            Left            =   4200
            TabIndex        =   131
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   34
            Left            =   4200
            TabIndex        =   130
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   35
            Left            =   4200
            TabIndex        =   129
            Top             =   2040
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   36
            Left            =   4200
            TabIndex        =   128
            Top             =   2400
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   37
            Left            =   4200
            TabIndex        =   127
            Top             =   2760
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "PNLPropVal"
            Height          =   255
            Index           =   39
            Left            =   4200
            TabIndex        =   126
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   40
            Left            =   4200
            TabIndex        =   125
            Top             =   1680
            Width           =   1695
         End
      End
      Begin VB.Frame fraTab 
         Caption         =   "fraTab(1)"
         Height          =   3375
         Index           =   1
         Left            =   120
         TabIndex        =   149
         Top             =   540
         Width           =   8895
         Begin VB.TextBox TXTName 
            DataField       =   "Name"
            DataSource      =   "Data1"
            Height          =   285
            Left            =   4140
            Locked          =   -1  'True
            TabIndex        =   157
            Text            =   "TXTName"
            Top             =   1050
            Width           =   4575
         End
         Begin VB.TextBox TXTCAS 
            DataField       =   "CAS"
            DataSource      =   "Data1"
            Height          =   285
            Left            =   4140
            Locked          =   -1  'True
            TabIndex        =   156
            Text            =   "TXTCAS"
            Top             =   1410
            Width           =   3135
         End
         Begin VB.TextBox TXTFormula 
            DataField       =   "Formula"
            DataSource      =   "Data1"
            Height          =   285
            Left            =   4140
            Locked          =   -1  'True
            TabIndex        =   155
            Text            =   "TXTFormula"
            Top             =   1770
            Width           =   3135
         End
         Begin VB.TextBox TXTSource 
            DataField       =   "Source"
            DataSource      =   "Data1"
            Height          =   285
            Left            =   4140
            Locked          =   -1  'True
            TabIndex        =   154
            Text            =   "TXTSource"
            Top             =   2490
            Width           =   3135
         End
         Begin VB.TextBox TXTOpT 
            Height          =   285
            Left            =   4140
            TabIndex        =   153
            Text            =   "TXTOpT"
            Top             =   330
            Width           =   1335
         End
         Begin VB.TextBox TXTOpP 
            Height          =   285
            Left            =   4140
            TabIndex        =   152
            Text            =   "TXTOpP"
            Top             =   690
            Width           =   1335
         End
         Begin VB.TextBox TXTSMILES 
            DataField       =   "Smiles"
            DataSource      =   "Data1"
            Height          =   285
            Left            =   4140
            Locked          =   -1  'True
            TabIndex        =   151
            Text            =   "TXTSMILES"
            Top             =   2850
            Width           =   3135
         End
         Begin VB.TextBox TXTFamily 
            Height          =   285
            Left            =   4140
            Locked          =   -1  'True
            TabIndex        =   150
            Text            =   "TXTFamily"
            Top             =   2130
            Width           =   3135
         End
         Begin VB.Label LBLOpTUnits 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "LBLOpTUnits"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   5580
            TabIndex        =   167
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label LBLOpPUnits 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "LBLOpPUnits"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   5580
            TabIndex        =   166
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Operating Temperature"
            Height          =   255
            Index           =   41
            Left            =   180
            TabIndex        =   165
            Top             =   330
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "CAS"
            Height          =   255
            Index           =   44
            Left            =   180
            TabIndex        =   164
            Top             =   1410
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Name"
            Height          =   255
            Index           =   43
            Left            =   180
            TabIndex        =   163
            Top             =   1050
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Operating Pressure"
            Height          =   255
            Index           =   42
            Left            =   180
            TabIndex        =   162
            Top             =   690
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Formula"
            Height          =   255
            Index           =   45
            Left            =   180
            TabIndex        =   161
            Top             =   1770
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Family"
            Height          =   255
            Index           =   46
            Left            =   180
            TabIndex        =   160
            Top             =   2130
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Source"
            Height          =   255
            Index           =   47
            Left            =   180
            TabIndex        =   159
            Top             =   2490
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "SMILES"
            Height          =   255
            Index           =   48
            Left            =   180
            TabIndex        =   158
            Top             =   2850
            Width           =   3735
         End
      End
      Begin VB.Frame fraTab 
         Caption         =   "fraTab(0)"
         Height          =   3375
         Index           =   0
         Left            =   60
         TabIndex        =   168
         Top             =   240
         Width           =   8745
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Salmonidae, 96h, LC50"
            Height          =   255
            Index           =   56
            Left            =   210
            TabIndex        =   192
            Top             =   2880
            Width           =   3735
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   48
            Left            =   6090
            TabIndex        =   191
            Top             =   2880
            Width           =   2175
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   48
            Left            =   4170
            TabIndex        =   190
            Top             =   2880
            Width           =   1695
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Salmonidae, 48h, LC50"
            Height          =   255
            Index           =   55
            Left            =   210
            TabIndex        =   189
            Top             =   2520
            Width           =   3735
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   47
            Left            =   6090
            TabIndex        =   188
            Top             =   2520
            Width           =   2175
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   47
            Left            =   4170
            TabIndex        =   187
            Top             =   2520
            Width           =   1695
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Salmonidae, 24h, LC50"
            Height          =   255
            Index           =   54
            Left            =   210
            TabIndex        =   186
            Top             =   2160
            Width           =   3735
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   46
            Left            =   6090
            TabIndex        =   185
            Top             =   2160
            Width           =   2175
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   46
            Left            =   4170
            TabIndex        =   184
            Top             =   2160
            Width           =   1695
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Fathead Minnow, 96h, LC50"
            Height          =   255
            Index           =   53
            Left            =   210
            TabIndex        =   183
            Top             =   1800
            Width           =   3735
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   45
            Left            =   6090
            TabIndex        =   182
            Top             =   1800
            Width           =   2175
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   45
            Left            =   4170
            TabIndex        =   181
            Top             =   1800
            Width           =   1695
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Fathead Minnow, 48h, LC50"
            Height          =   255
            Index           =   52
            Left            =   210
            TabIndex        =   180
            Top             =   1440
            Width           =   3735
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   44
            Left            =   6090
            TabIndex        =   179
            Top             =   1440
            Width           =   2175
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   44
            Left            =   4170
            TabIndex        =   178
            Top             =   1440
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   43
            Left            =   4170
            TabIndex        =   177
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   43
            Left            =   6090
            TabIndex        =   176
            Top             =   1080
            Width           =   2175
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Fathead Minnow, 24h, LC50"
            Height          =   255
            Index           =   51
            Left            =   210
            TabIndex        =   175
            Top             =   1080
            Width           =   3735
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   41
            Left            =   4170
            TabIndex        =   174
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   41
            Left            =   6090
            TabIndex        =   173
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Fathead Minnow, 48h, EC50"
            Height          =   255
            Index           =   49
            Left            =   210
            TabIndex        =   172
            Top             =   360
            Width           =   3735
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   42
            Left            =   4170
            TabIndex        =   171
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   42
            Left            =   6090
            TabIndex        =   170
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Fathead Minnow, 96h, EC50"
            Height          =   255
            Index           =   50
            Left            =   210
            TabIndex        =   169
            Top             =   720
            Width           =   3735
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'
'/////////////////////////////////////////////////////////////////////////////////////////////
' HANDLE TAB SYSTEM.
'/////////////////////////////////////////////////////////////////////////////////////////////
Sub Redisplay_fraTabHolder()
Dim This_Idx As Integer
Dim i As Integer
  This_Idx = cboTabSelector.ItemData(cboTabSelector.ListIndex)
  For i = 0 To 8
    If (i <> This_Idx) Then
      fraTab(i).Visible = False
    End If
  Next i
  fraTab(This_Idx).Visible = True
End Sub
Private Sub cboTabSelector_Click()
  Call Redisplay_fraTabHolder
End Sub
Sub Populate_cboTabSelector()
Dim i As Integer
  cboTabSelector.Clear
  For i = 0 To 8
    cboTabSelector.AddItem Trim$(Str$(i))
    cboTabSelector.ItemData(cboTabSelector.NewIndex) = i
  Next i
  cboTabSelector.ListIndex = 0
  Call Redisplay_fraTabHolder
End Sub
'/////////////////////////////////////////////////////////////////////////////////////////////
' HANDLE TAB SYSTEM.
'/////////////////////////////////////////////////////////////////////////////////////////////
'


Private Sub Form_Load()
  '
  ' POPULATE cboTabSelector.
  '
  Call Populate_cboTabSelector
End Sub




