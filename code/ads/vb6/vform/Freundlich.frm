VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmFreundlich 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Freundlich Isotherm Parameters for {ComponentName}"
   ClientHeight    =   8145
   ClientLeft      =   4080
   ClientTop       =   675
   ClientWidth     =   9465
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   9465
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   9360
      ScaleHeight     =   795
      ScaleWidth      =   1275
      TabIndex        =   71
      Top             =   2160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin Threed.SSFrame fraIsothermDB 
      Height          =   4245
      Left            =   240
      TabIndex        =   27
      Top             =   3360
      Width           =   9255
      _Version        =   65536
      _ExtentX        =   16325
      _ExtentY        =   7488
      _StockProps     =   14
      Caption         =   "{What} Phase Isotherm Database"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSFrame fraTwo 
         Height          =   3825
         Left            =   4410
         TabIndex        =   35
         Top             =   300
         Width           =   4635
         _Version        =   65536
         _ExtentX        =   8176
         _ExtentY        =   6747
         _StockProps     =   14
         Caption         =   "{X} {What}-Phase isotherm(s) for {chemical_name}"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.ListBox lstRange 
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
            Height          =   1590
            Index           =   0
            Left            =   210
            TabIndex        =   37
            Top             =   900
            Width           =   1695
         End
         Begin VB.ListBox lstRange 
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
            Height          =   1590
            Index           =   1
            Left            =   2130
            TabIndex        =   36
            Top             =   900
            Width           =   2415
         End
         Begin VB.Label lblText 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "K"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   1470
            TabIndex        =   54
            Top             =   330
            Width           =   315
         End
         Begin VB.Label lblValue 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblValue(0)"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   1830
            TabIndex        =   53
            Top             =   315
            Width           =   975
         End
         Begin VB.Label lblValue 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblValue(1)"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   510
            TabIndex        =   52
            Top             =   315
            Width           =   975
         End
         Begin VB.Label lblText 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "1/n"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   60
            TabIndex        =   51
            Top             =   330
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Conc. Range (mg/L):"
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
            Left            =   2130
            TabIndex        =   50
            Top             =   690
            Width           =   2415
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "pH Range:"
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
            Left            =   210
            TabIndex        =   49
            Top             =   690
            Width           =   1695
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Adsorbent Type:"
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
            Height          =   225
            Left            =   180
            TabIndex        =   48
            Top             =   2715
            Width           =   1515
         End
         Begin VB.Label lblValue 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblValue(2)"
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
            Left            =   1770
            TabIndex        =   47
            Top             =   2700
            Width           =   2775
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Source:"
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
            TabIndex        =   46
            Top             =   3255
            Width           =   795
         End
         Begin VB.Label lblValue 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblValue(3)"
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
            Left            =   930
            TabIndex        =   45
            Top             =   3240
            Width           =   3615
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "(mg/g)*(L/mg)^(1/n)"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   2850
            TabIndex        =   44
            Top             =   330
            Width           =   1755
         End
         Begin VB.Label lblPhase 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblPhase"
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
            Left            =   3510
            TabIndex        =   43
            Top             =   2940
            Width           =   1035
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Temperature (C):"
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
            Height          =   195
            Left            =   180
            TabIndex        =   42
            Top             =   2955
            Width           =   1515
         End
         Begin VB.Label lblTemperature 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblTemperature"
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
            Left            =   1770
            TabIndex        =   41
            Top             =   2940
            Width           =   975
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Phase:"
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
            Height          =   195
            Left            =   2790
            TabIndex        =   40
            Top             =   2955
            Width           =   675
         End
         Begin VB.Label lblComments 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblComments"
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
            Left            =   930
            TabIndex        =   39
            Top             =   3480
            Width           =   3615
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Comment:"
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
            Height          =   195
            Left            =   -150
            TabIndex        =   38
            Top             =   3495
            Width           =   1035
         End
         Begin VB.Line Line2 
            X1              =   30
            X2              =   4590
            Y1              =   630
            Y2              =   630
         End
      End
      Begin Threed.SSFrame fraOne 
         Height          =   3825
         Left            =   90
         TabIndex        =   28
         Top             =   300
         Width           =   4335
         _Version        =   65536
         _ExtentX        =   7646
         _ExtentY        =   6747
         _StockProps     =   14
         Caption         =   "Select a component:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.ListBox lstCompo 
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
            Height          =   2370
            Left            =   120
            TabIndex        =   33
            Top             =   360
            Width           =   4095
         End
         Begin VB.ComboBox cboSortMethod 
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
            Left            =   2880
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   3420
            Width           =   1335
         End
         Begin Threed.SSCommand cmdSelect 
            Height          =   315
            Left            =   120
            TabIndex        =   30
            Top             =   3090
            Width           =   4095
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Select Chemic&al"
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
         Begin Threed.SSCommand cmdFind 
            Height          =   315
            Index           =   1
            Left            =   2160
            TabIndex        =   31
            Top             =   2760
            Width           =   2055
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Find A&gain"
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
         Begin Threed.SSCommand cmdFind 
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   32
            Top             =   2760
            Width           =   2055
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "&Find"
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
         Begin VB.Label lblEmpty_lstCompo 
            Alignment       =   2  'Center
            Caption         =   "No Components Available"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   120
            TabIndex        =   69
            Top             =   210
            Visible         =   0   'False
            Width           =   4095
         End
         Begin VB.Label lblInput 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Sorting Method:"
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
            Left            =   1140
            TabIndex        =   34
            Top             =   3480
            Width           =   1695
         End
      End
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   525
      Left            =   7920
      TabIndex        =   70
      Top             =   7320
      Visible         =   0   'False
      Width           =   1245
   End
   Begin Threed.SSFrame fraUserInput 
      Height          =   855
      Left            =   2400
      TabIndex        =   60
      Top             =   6960
      Width           =   4575
      _Version        =   65536
      _ExtentX        =   8070
      _ExtentY        =   1508
      _StockProps     =   14
      Caption         =   "User Input"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox UserK 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1860
         TabIndex        =   62
         Text            =   "UserK"
         Top             =   270
         Width           =   975
      End
      Begin VB.TextBox UserOneOverN 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   600
         TabIndex        =   61
         Text            =   "UserOneOverN"
         Top             =   270
         Width           =   975
      End
      Begin VB.Label lblText 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "K"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   1500
         TabIndex        =   65
         Top             =   300
         Width           =   315
      End
      Begin VB.Label lblText 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1/n"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   64
         Top             =   300
         Width           =   375
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "(mg/g)*(L/mg)^(1/n)"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   2880
         TabIndex        =   63
         Top             =   300
         Width           =   1575
      End
   End
   Begin Threed.SSFrame fraSource 
      Height          =   1665
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   9255
      _Version        =   65536
      _ExtentX        =   16325
      _ExtentY        =   2937
      _StockProps     =   14
      Caption         =   "Source of K and 1/n:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton Command2 
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
         Left            =   7560
         TabIndex        =   72
         ToolTipText     =   "Click here to print current screen to selected printer"
         Top             =   1240
         Width           =   1455
      End
      Begin Threed.SSOption optFreundlichSource 
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   1
         Top             =   900
         Width           =   3100
         _Version        =   65536
         _ExtentX        =   5468
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "U&ser Input"
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
         Enabled         =   0   'False
      End
      Begin Threed.SSOption optFreundlichSource 
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   3100
         _Version        =   65536
         _ExtentX        =   5468
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Isotherm Parameter &Estimation"
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
         Enabled         =   0   'False
      End
      Begin Threed.SSOption optFreundlichSource 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   300
         Width           =   3100
         _Version        =   65536
         _ExtentX        =   5468
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Isotherm &Database"
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
         Enabled         =   0   'False
      End
      Begin Threed.SSCommand cmdCancelOK 
         Height          =   495
         Index           =   1
         Left            =   7530
         TabIndex        =   58
         TabStop         =   0   'False
         ToolTipText     =   "Click here to save the changes you have made to the Freundlich parameters on this window"
         Top             =   690
         Width           =   1545
         _Version        =   65536
         _ExtentX        =   2725
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "&OK"
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
      Begin Threed.SSCommand cmdCancelOK 
         Height          =   495
         Index           =   0
         Left            =   7530
         TabIndex        =   59
         TabStop         =   0   'False
         ToolTipText     =   "Click here to abandon any changes you have made to the Freundlich parameters on this window"
         Top             =   210
         Width           =   1545
         _Version        =   65536
         _ExtentX        =   2725
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "&Cancel"
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
      Begin Threed.SSPanel sspanel_Warning 
         Height          =   1095
         Left            =   3330
         TabIndex        =   66
         Top             =   150
         Width           =   4065
         _Version        =   65536
         _ExtentX        =   7170
         _ExtentY        =   1931
         _StockProps     =   15
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.Label lblWarning 
            Alignment       =   2  'Center
            Caption         =   "lblWarning"
            ForeColor       =   &H000000FF&
            Height          =   915
            Left            =   120
            TabIndex        =   67
            Top             =   90
            Width           =   3825
         End
      End
   End
   Begin Threed.SSFrame fraIPES 
      Height          =   2505
      Left            =   90
      TabIndex        =   4
      Top             =   1800
      Width           =   9255
      _Version        =   65536
      _ExtentX        =   16325
      _ExtentY        =   4419
      _StockProps     =   14
      Caption         =   "Isotherm Parameter Estimation (IPE)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ComboBox cboMethod 
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
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   555
         Width           =   4095
      End
      Begin Threed.SSCommand cmdCalculate 
         Height          =   375
         Left            =   4470
         TabIndex        =   5
         Top             =   1650
         Width           =   4695
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "&Perform IPE Calculations"
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
      Begin Threed.SSFrame fraPolanyi 
         Height          =   1425
         Left            =   4470
         TabIndex        =   13
         Top             =   140
         Width           =   4695
         _Version        =   65536
         _ExtentX        =   8281
         _ExtentY        =   2514
         _StockProps     =   14
         Caption         =   "Polanyi Parameters"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox txtInput 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
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
            Index           =   13
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   17
            Text            =   "txtInput(13)"
            Top             =   240
            Width           =   2655
         End
         Begin VB.TextBox txtInput 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
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
            Left            =   3480
            Locked          =   -1  'True
            TabIndex        =   16
            Text            =   "txtInput(0)"
            Top             =   540
            Width           =   1095
         End
         Begin VB.TextBox txtInput 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
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
            Left            =   3480
            Locked          =   -1  'True
            TabIndex        =   15
            Text            =   "txtInput(1)"
            Top             =   780
            Width           =   1095
         End
         Begin VB.TextBox txtInput 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
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
            Index           =   10
            Left            =   3480
            Locked          =   -1  'True
            TabIndex        =   14
            Text            =   "txtInput(10)"
            Top             =   1020
            Width           =   1095
         End
         Begin Threed.SSCommand cmdEditPolanyi 
            Height          =   375
            Left            =   120
            TabIndex        =   68
            Top             =   900
            Width           =   1575
            _Version        =   65536
            _ExtentX        =   2778
            _ExtentY        =   661
            _StockProps     =   78
            Caption         =   "Edi&t Parameters"
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
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Adsorbent:"
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
            Height          =   195
            Left            =   900
            TabIndex        =   21
            Top             =   270
            Width           =   975
         End
         Begin VB.Label lblInput 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "W0 (cm3/g)"
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
            Left            =   1890
            TabIndex        =   20
            Top             =   570
            Width           =   1515
         End
         Begin VB.Label lblInput 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "BB (mol/cal)^GM"
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
            Left            =   1830
            TabIndex        =   19
            Top             =   810
            Width           =   1575
         End
         Begin VB.Label lblInput 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "GM"
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
            Left            =   1830
            TabIndex        =   18
            Top             =   1050
            Width           =   1575
         End
      End
      Begin Threed.SSFrame fraAdditional 
         Height          =   945
         Left            =   90
         TabIndex        =   22
         Top             =   1080
         Width           =   3705
         _Version        =   65536
         _ExtentX        =   6535
         _ExtentY        =   1667
         _StockProps     =   14
         Caption         =   "Additional Parameters"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox txtInput 
            Alignment       =   2  'Center
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
            Index           =   12
            Left            =   2460
            TabIndex        =   24
            Text            =   "txtInput(12)"
            Top             =   570
            Width           =   1095
         End
         Begin VB.TextBox txtInput 
            Alignment       =   2  'Center
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
            Index           =   11
            Left            =   2460
            TabIndex        =   23
            Text            =   "txtInput(11)"
            Top             =   270
            Width           =   1095
         End
         Begin VB.Label lblInput 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "No. of regression points:"
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
            Left            =   120
            TabIndex        =   26
            Top             =   600
            Width           =   2175
         End
         Begin VB.Label lblInput 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Order of magnitude:"
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
            Left            =   240
            TabIndex        =   25
            Top             =   300
            Width           =   2055
         End
      End
      Begin VB.Label lblText 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "K"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   5970
         TabIndex        =   12
         Top             =   2145
         Width           =   315
      End
      Begin VB.Label lblValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblValue(5)"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   6330
         TabIndex        =   11
         Top             =   2130
         Width           =   975
      End
      Begin VB.Label lblValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblValue(4)"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   5010
         TabIndex        =   10
         Top             =   2130
         Width           =   975
      End
      Begin VB.Label lblText 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1/n"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   4560
         TabIndex        =   9
         Top             =   2145
         Width           =   375
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "(mg/g)*(L/mg)^(1/n)"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   7350
         TabIndex        =   8
         Top             =   2145
         Width           =   1755
      End
      Begin VB.Label lblEstimationMethod 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Estimation Method:"
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
         TabIndex        =   7
         Top             =   330
         Width           =   1875
      End
   End
   Begin Threed.SSPanel sspanel_StatusBar 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   55
      Top             =   7740
      Width           =   9465
      _Version        =   65536
      _ExtentX        =   16695
      _ExtentY        =   714
      _StockProps     =   15
      ForeColor       =   -2147483640
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSPanel sspanel_Dirty 
         Height          =   285
         Left            =   60
         TabIndex        =   56
         Top             =   60
         Width           =   2115
         _Version        =   65536
         _ExtentX        =   3731
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "sspanel_Dirty"
         ForeColor       =   -2147483640
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin Threed.SSPanel sspanel_Status 
         Height          =   285
         Left            =   2220
         TabIndex        =   57
         Top             =   60
         Width           =   7200
         _Version        =   65536
         _ExtentX        =   12700
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "sspanel_Status"
         ForeColor       =   -2147483640
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
   End
End
Attribute VB_Name = "frmFreundlich"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'Dim FORM_MODE As Integer
'Const FORM_MODE_ADDNEW = 1
'Const FORM_MODE_EDIT = 2
Dim USER_HIT_CANCEL As Boolean
Dim USER_HIT_OK As Boolean
Dim frmFreundlich_Is_Dirty As Boolean
Dim ActivatedYet As Boolean

Dim HALT_OPTFREUNDLICHSOURCE As Boolean
Dim HALT_LSTCOMPO As Boolean
Dim HALT_LSTRANGE As Boolean
Dim HALT_CBOMETHOD As Boolean
Const CBOSORTMETHOD_NAME = 1
Const CBOSORTMETHOD_CAS = 2

Dim SaveOldComponent As ComponentPropertyType
  
Dim DB_Isotherm As Database
Dim Find_String As String




Const frmFreundlich_declarations_end = True


Sub frmFreundlich_Run( _
    OUTPUT_Raise_Dirty_Flag As Boolean)
  On Error GoTo err_frmFreundlich_Run
  'NOTE: THE FOLLOWING OpenDatabase() COMMAND MUST BE
  'SPECIFIED EXACTLY AS-IS, OR ELSE IT WILL FAIL.
  Set DB_Isotherm = _
      Ws1.OpenDatabase(fn_DB_Isotherm, True, False, _
      ";pwd=" & decrypt_string(Encrypted_User_Password))
  'Set DB_Isotherm = ws1.OpenDatabase(fn_DB_Isotherm)
  frmFreundlich.Show 1
  If (USER_HIT_OK) Then
    OUTPUT_Raise_Dirty_Flag = True
  Else
    OUTPUT_Raise_Dirty_Flag = False
  End If
  DB_Isotherm.Close
  Exit Sub
exit_err_frmFreundlich_Run:
  Exit Sub
err_frmFreundlich_Run:
  Call Show_Trapped_Error("frmFreundlich_Run")
  OUTPUT_Raise_Dirty_Flag = False
  Resume exit_err_frmFreundlich_Run
End Sub
Sub LOCAL___Reset_DemoVersionDisablings()
  If (IsThisADemo() = True) Then
    cmdCancelOK(1).Enabled = False
    cmdCalculate.Enabled = False
  End If
End Sub


Sub frmFreundlich_GenericStatus_Set(fn_Text As String)
  Me.sspanel_Status = fn_Text
End Sub
Sub frmFreundlich_DirtyStatus_Set(newVal As Boolean)
  If (newVal) Then
    Me.sspanel_Dirty = "Data Changed"
    Me.sspanel_Dirty.ForeColor = QBColor(12)
  Else
    Me.sspanel_Dirty = "Unchanged"
    Me.sspanel_Dirty.ForeColor = QBColor(0)
  End If
End Sub
Sub frmFreundlich_DirtyStatus_Set_Current()
  Call frmFreundlich_DirtyStatus_Set(frmFreundlich_Is_Dirty)
End Sub
Sub frmFreundlich_DirtyStatus_Throw()
  frmFreundlich_Is_Dirty = True
  Call frmFreundlich_DirtyStatus_Set_Current
End Sub
Sub frmFreundlich_DirtyStatus_Clear()
  frmFreundlich_Is_Dirty = False
  Call frmFreundlich_DirtyStatus_Set_Current
End Sub


Sub populate_lstCompo()
Dim Rs1 As Recordset
Dim Current_Criteria As String
Dim SAVE_CURRENT_POSITION As Long
Dim NEW_LISTINDEX As Integer
Dim This_ID As Long
Dim NumRecords As Long
Dim TempStr As String * 15
Dim Output_Line As String
Dim PhaseCode As String
Dim SortCode As String
Dim ThisChemicalName As String
Dim ThisChemicalCAS As String
Dim LastChemicalName As String
Dim LastChemicalCAS As String
  HALT_LSTCOMPO = True
  On Error GoTo err_populate_lstCompo
  '
  ' SAVE CURRENT POSITION.
  '
  If (lstCompo.ListCount > 0) And (lstCompo.ListIndex >= 0) Then
    SAVE_CURRENT_POSITION = lstCompo.ItemData(lstCompo.ListIndex)
  Else
    SAVE_CURRENT_POSITION = -1
  End If
  '
  ' SET UP SEARCH CRITERIA.
  '
  Select Case Bed.Phase
    Case 0: PhaseCode = "Liquid"
    Case 1: PhaseCode = "Gas"
  End Select
  Select Case cboSortMethod.ItemData(cboSortMethod.ListIndex)
    Case CBOSORTMETHOD_NAME: SortCode = "Name, [Component Number]"
    Case CBOSORTMETHOD_CAS: SortCode = "[Component Number], Name"
  End Select
  'Current_Criteria = "select * from [Chemicals] " & _
  '    "order by [Name]"
  Current_Criteria = "select * from Isotherms" & _
      " where Phase = '" & PhaseCode & "'" & _
      " order by " & SortCode
  '
  ' START SEARCH.
  '
  Set Rs1 = _
      DB_Isotherm.OpenRecordset(Current_Criteria)
  On Error Resume Next
  Rs1.MoveFirst
  Rs1.MoveLast
  Rs1.MoveFirst
  NumRecords = Rs1.RecordCount
  On Error GoTo err_populate_lstCompo
  '
  ' POPULATE LISTBOX.
  '
  lstCompo.Clear
  If (NumRecords = 0) Then
    ' NO RECORDS AVAILABLE.
    lstCompo.Visible = False
    lblEmpty_lstCompo.Move lstCompo.Left, lstCompo.Top
    lblEmpty_lstCompo.Visible = True
  Else
    '
    ' DISPLAY RECORDS.
    '
    lstCompo.Visible = True
    lblEmpty_lstCompo.Visible = False
    NEW_LISTINDEX = -1
    LastChemicalName = "n/a yet"
    LastChemicalCAS = "n/a yet"
    Do Until Rs1.EOF
      ThisChemicalName = Database_Get_String(Rs1, "Name")
      ThisChemicalCAS = Database_Get_String(Rs1, "Component Number")
      If ((ThisChemicalName <> LastChemicalName) Or _
          (ThisChemicalCAS <> LastChemicalCAS)) Then
        '
        ' THIS "IF" STATEMENT EXISTS IN ORDER TO ELIMINATE DUPLICATE
        ' CHEMICALS FROM THE LIST.  THIS CODE DEPENDS COMPLETELY
        ' ON THE FACT THAT THE LIST IS SORTED !
        '
        LastChemicalCAS = ThisChemicalCAS
        LastChemicalName = ThisChemicalName
        TempStr = ThisChemicalCAS
                'THIS STRING IS ENSURED TO BE 15 CHARACTERS LONG.
        Output_Line = _
            TempStr & _
            " " & _
            ThisChemicalName
        lstCompo.AddItem Output_Line
        This_ID = Database_Get_Long(Rs1, "ID")
        lstCompo.ItemData(lstCompo.NewIndex) = This_ID
        If (SAVE_CURRENT_POSITION <> -1) Then
          If (SAVE_CURRENT_POSITION = This_ID) Then
            NEW_LISTINDEX = lstCompo.NewIndex
          End If
        End If
      End If
      Rs1.MoveNext
    Loop
    If (lstCompo.ListCount > 0) And (NEW_LISTINDEX > -1) Then
      HALT_LSTCOMPO = True
      lstCompo.ListIndex = NEW_LISTINDEX
      HALT_LSTCOMPO = False
    End If
  End If
  '
  ' CLOSE DATABASE AND EXIT.
  '
  Rs1.Close
  HALT_LSTCOMPO = False
  Exit Sub
exit_err_populate_lstCompo:
  HALT_LSTCOMPO = False
  Exit Sub
err_populate_lstCompo:
  Call Show_Trapped_Error("populate_lstCompo")
  Resume exit_err_populate_lstCompo
End Sub
Sub populate_lstRange(ThisCAS As String, ThisChemical As String)
Dim PHASE_CODE As Integer
Dim Rs1 As Recordset
Dim Current_Criteria As String
Dim SAVE_CURRENT_POSITION As Long
Dim This_ID As Long
Dim NEW_LISTINDEX As Long
Dim NumRecords As Long
Dim PhaseCode As String
Dim ThisCMin As Double
Dim ThisCMax As Double
Dim ThisPHMin As Double
Dim ThisPHMax As Double
Dim ThisDbl As Double
Dim ThisOutput As String
  On Error GoTo err_populate_lstRange
  'GET PHASE CODE.
  Select Case Bed.Phase
    Case 0: PhaseCode = "Liquid"
    Case 1: PhaseCode = "Gas"
  End Select
  'SAVE CURRENT POSITION.
  If (lstRange(0).ListCount > 0) And (lstRange(0).ListIndex >= 0) Then
    SAVE_CURRENT_POSITION = lstRange(0).ItemData(lstRange(0).ListIndex)
  Else
    SAVE_CURRENT_POSITION = -1
  End If
  'SET UP SEARCH CRITERIA.
  If (Trim$(ThisCAS) = "0") Then ThisCAS = ""
  If (Trim$(ThisCAS) <> "") Then
    Current_Criteria = "select * from Isotherms" & _
        " where Phase = '" & PhaseCode & "'" & _
        " and [Component Number] = " & Trim$(ThisCAS) & _
        " order by CarbonName"
  Else
    Current_Criteria = "select * from Isotherms" & _
        " where Phase = '" & PhaseCode & "'" & _
        " and Name = '" & Trim$(ThisChemical) & "'" & _
        " order by CarbonName"
  End If
  'START SEARCH.
  Set Rs1 = _
      DB_Isotherm.OpenRecordset(Current_Criteria)
  On Error Resume Next
  Rs1.MoveFirst
  Rs1.MoveLast
  Rs1.MoveFirst
  NumRecords = Rs1.RecordCount
  On Error GoTo err_populate_lstRange
  'POPULATE LISTBOX.
  lstRange(0).Clear
  lstRange(1).Clear
  If (NumRecords = 0) Then
    'NO RECORDS AVAILABLE.
    fraTwo.Caption = "No Isotherms Available."
  Else
    'DISPLAY RECORDS.
    fraTwo.Caption = Trim$(Str$(NumRecords)) & " " & _
        PhaseCode & " Phase Isotherm" & _
        IIf(NumRecords = 1, "", "s") & _
        " Available"
    NEW_LISTINDEX = -1
    Do Until Rs1.EOF
      This_ID = Database_Get_Long(Rs1, "ID")
      ThisCMin = Database_Get_Double(Rs1, "C min")
      ThisCMax = Database_Get_Double(Rs1, "C max")
      ThisPHMin = Database_Get_Double(Rs1, "pH min")
      ThisPHMax = Database_Get_Double(Rs1, "pH max")
      If (ThisPHMin = 0#) And (ThisPHMax = 0#) Then
        ThisOutput = "No pH Range"
      Else
        If (ThisPHMin = 0#) Or (ThisPHMax = 0#) Then
          If (ThisPHMin <> 0#) Then ThisDbl = ThisPHMin
          If (ThisPHMax <> 0#) Then ThisDbl = ThisPHMax
          ThisOutput = Format$(ThisDbl, "0.000")
        Else
          ThisOutput = Format$(ThisPHMin, "0.000") & " - " & _
              Format$(ThisPHMax, "0.000")
        End If
      End If
      lstRange(0).AddItem ThisOutput
      lstRange(0).ItemData(lstRange(0).NewIndex) = This_ID
      If (ThisCMin = 0#) And (ThisCMax = 0#) Then
        ThisOutput = "No Conc. Range"
      Else
        If (ThisCMin = 0#) Or (ThisCMax = 0#) Then
          If (ThisCMin <> 0#) Then ThisDbl = ThisCMin
          If (ThisCMax <> 0#) Then ThisDbl = ThisCMax
          ThisOutput = Format$(ThisDbl, "0.000")
        Else
          ThisOutput = Format$(ThisCMin, "0.000") & " - " & _
              Format$(ThisCMax, "0.000")
        End If
      End If
      lstRange(1).AddItem ThisOutput
      lstRange(1).ItemData(lstRange(1).NewIndex) = This_ID
      If (SAVE_CURRENT_POSITION <> -1) Then
        If (SAVE_CURRENT_POSITION = This_ID) Then
          NEW_LISTINDEX = lstRange(0).NewIndex
        End If
      End If
      Rs1.MoveNext
    Loop
    If (lstRange(0).ListCount > 0) And (NEW_LISTINDEX > -1) Then
      lstRange(0).ListIndex = NEW_LISTINDEX
      lstRange(1).ListIndex = NEW_LISTINDEX
    End If
  End If
  'CLOSE DATABASE AND EXIT.
  Rs1.Close
  Exit Sub
exit_err_populate_lstRange:
  Exit Sub
err_populate_lstRange:
  Call Show_Trapped_Error("populate_lstRange")
  Resume exit_err_populate_lstRange
End Sub
Sub populate_lblValue(This_ID As Long)
Dim Rs1 As Recordset
Dim Current_Criteria As String
Dim NumRecords As Long
  On Error GoTo err_populate_lblValue
  'SET UP SEARCH CRITERIA.
  Current_Criteria = "select * from Isotherms" & _
      " where ID = " & Trim$(Str$(This_ID)) & _
      " order by CarbonName"
  'START SEARCH.
  Set Rs1 = _
      DB_Isotherm.OpenRecordset(Current_Criteria)
  On Error Resume Next
  Rs1.MoveFirst
  Rs1.MoveLast
  Rs1.MoveFirst
  NumRecords = Rs1.RecordCount
  On Error GoTo err_populate_lblValue
  'POPULATE LISTBOX.
  If (NumRecords = 0) Then
    'COULD NOT FIND THAT ISOTHERM (WEIRD PROBLEM).
  Else
    'DISPLAY RECORD.
    Call AssignCaptionAndTag(lblValue(2), Database_Get_String(Rs1, "CarbonName"))
    Call AssignCaptionAndTag(lblTemperature, Database_Get_Double(Rs1, "Tmin"))
    Call AssignCaptionAndTag(lblPhase, Database_Get_String(Rs1, "Phase"))
    Call AssignCaptionAndTag(lblValue(3), Database_Get_String(Rs1, "Source"))
    Call AssignCaptionAndTag(lblComments, Database_Get_String(Rs1, "Comments"))
    Component(0).IsothermDB_OneOverN = Database_Get_Double(Rs1, "1/n")
    Component(0).IsothermDB_K = Database_Get_Double(Rs1, "K")
  End If
  'CLOSE DATABASE AND EXIT.
  Rs1.Close
  Exit Sub
exit_err_populate_lblValue:
  Exit Sub
err_populate_lblValue:
  Call Show_Trapped_Error("populate_lblValue")
  Resume exit_err_populate_lblValue
End Sub


Sub Clear_lblValue()
  lblValue(0) = ""      'ISODB : K.
  lblValue(1) = ""      'ISODB : 1/N.
  lblValue(2) = ""      'ISODB : ADSORBENT TYPE.
  lblTemperature = ""   'ISODB : TEMP.
  lblPhase = ""         'ISODB : PHASE.
  lblValue(3) = ""      'ISODB : SOURCE.
  lblComments = ""      'ISODB : COMMENTS.
End Sub


'Returns:
'- TRUE = Succeeded
'- FALSE = Failed
Function Search_String( _
    J As Integer, _
    ShowErrorMessages As Integer) As Boolean
Dim i As Integer
Dim Res As Integer
  'If (fraIsothermDB.Visible) Then
  '  lstCompo.SetFocus
  'End If
  'For I = J + 1 To lstCompo.ListCount
  For i = J + 1 To lstCompo.ListCount - 1
    Res = InStr(1, lstCompo.List(i), Find_String, 1)
    If (Res > 0) Then
      'NOTE: BY HALTING lstCompo_Click(), THIS ALLOWS THE
      'COMPONENT TO BE SELECTED WITHOUT CLEARING THE ISOTHERM DB
      'VALUES OF K AND 1/N.
      'lstCompo.ListIndex = I
      HALT_LSTCOMPO = True
      Call Do_Select_Component(i)
      HALT_LSTCOMPO = False
      'If (fraIsothermDB.Visible) Then lstCompo.SetFocus
      Search_String = True
      Exit Function
    End If
  Next i
  For i = 0 To J
    Res = InStr(1, lstCompo.List(i), Find_String, 1)
    If (Res > 0) Then
      'NOTE: BY HALTING lstCompo_Click(), THIS ALLOWS THE
      'COMPONENT TO BE SELECTED WITHOUT CLEARING THE ISOTHERM DB
      'VALUES OF K AND 1/N.
      'lstCompo.ListIndex = I
      HALT_LSTCOMPO = True
      Call Do_Select_Component(i)
      HALT_LSTCOMPO = False
      'If (fraIsothermDB.Visible) Then lstCompo.SetFocus
      Search_String = True
      Exit Function
    End If
  Next i
  '----- If not found, show error message: -----
  If (ShowErrorMessages) Then
    Call Show_Error("String Not Found: " & Chr$(34) & _
        Trim$(Find_String) & Chr$(34))
  End If
  Search_String = False
End Function
'Returns:
'- TRUE = Succeeded
'- FALSE = Failed
Function Do_Search_For_Text(ShowErrorMessages As Integer) As Integer
Dim LIST_INDEX As Integer
  LIST_INDEX = lstCompo.ListIndex
  Do_Search_For_Text = Search_String(LIST_INDEX, ShowErrorMessages)
End Function


Sub Populate_cboSortMethod()
  cboSortMethod.Clear
  cboSortMethod.AddItem "By Name"
  cboSortMethod.ItemData(cboSortMethod.NewIndex) = CBOSORTMETHOD_NAME
  cboSortMethod.AddItem "By CAS"
  cboSortMethod.ItemData(cboSortMethod.NewIndex) = CBOSORTMETHOD_CAS
  cboSortMethod.ListIndex = 0
End Sub
Sub Populate_cboMethod()
Dim NewTag As Integer
Dim i As Integer
  HALT_CBOMETHOD = True
'xaxaxa (12/10/97)
  Select Case Bed.Phase
    Case 0
      ' ***** Phase = Water *****
      'lblInput(3).Visible = False
      'txtInput(3).Visible = False
      cboMethod.Clear
      cboMethod.AddItem "3 - Parameter Polanyi Isotherm Correlation"
      cboMethod.ItemData(cboMethod.NewIndex) = IPESMETHOD_LIQ_3PARAM
      cboMethod.AddItem "D-R Uniform Adsorbate"
      cboMethod.ItemData(cboMethod.NewIndex) = IPESMETHOD_LIQ_DRUNIFORM
      'cboMethod.AddItem "D-R Non-Uniform Adsorbate"
    Case 1
      ' ***** Phase = Air *****
      'lblInput(3).Visible = True
      'txtInput(3).Visible = True
      cboMethod.Clear
      cboMethod.AddItem "D-R based on Spreading Pressure Eval."
      cboMethod.ItemData(cboMethod.NewIndex) = IPESMETHOD_GAS_DRSPREADINGP
      'EJO 12/16/97 -- This correlation killed off today.
      'cboMethod.AddItem "D-R Isotherm Correlation for RH < 50%"
      'If (check_internal_to_mtu()) Then
      '  cboMethod.AddItem "D-R based on Spreading Pressure Eval. (MTU only)"
      'End If
      'cboMethod.AddItem "Calgon BPL"
      'cboMethod.AddItem "D-R based on Spreading Pressure Evaluation"
      'EJO 8/5/97 -- These two correlations were killed off!
  End Select
  'PERFORM LOOKUP FOR CURRENT METHOD.
  '[NEW AS OF 8/28/98.]
  NewTag = 0
  For i = 0 To cboMethod.ListCount - 1
    If (Component(0).IPES_EstimationMethod = _
        cboMethod.ItemData(i)) Then
      NewTag = i
      Exit For
    End If
  Next i
  cboMethod.ListIndex = NewTag
  HALT_CBOMETHOD = False
End Sub
Sub frmFreundlich_PopulateUnits()
  Call unitsys_register(frmFreundlich, lblInput(5), _
      txtInput(11), Nothing, "", _
      "", "", "", "", 100#, False)
  Call unitsys_register(frmFreundlich, lblInput(6), _
      txtInput(12), Nothing, "", _
      "", "", "", "", 100#, False)
  Call unitsys_register(frmFreundlich, lblText(5), _
      UserOneOverN, Nothing, "", _
      "", "", "", "", 100#, False)
  Call unitsys_register(frmFreundlich, lblText(4), _
      UserK, Nothing, "", _
      "", "", "", "", 100#, False)
End Sub


Private Sub cboMethod_Click()
Dim OldValue As Integer
Dim NewValue As Integer
  If (HALT_CBOMETHOD) Then Exit Sub
  OldValue = Component(0).IPES_EstimationMethod
  NewValue = cboMethod.ItemData(cboMethod.ListIndex)
  If (OldValue <> NewValue) Then
    Component(0).IPES_EstimationMethod = NewValue
    'THROW DIRTY FLAG.
    Call frmFreundlich_DirtyStatus_Throw
  End If
End Sub


Private Sub cboSortMethod_Click()
  Call populate_lstCompo
End Sub


Private Sub cmdCalculate_Click()
Dim WhichModule As Integer
Dim INPUT_NL As Integer
Dim INPUT_OMAG As Double
Dim Raise_Dirty_Flag As Boolean
  'CHECK FOR ZERO VALUES OF POLANYI PARAMETERS.
  If (Carbon.W0 = 0#) Or (Carbon.BB = 0#) Or (Carbon.PolanyiExponent = 0#) Then
    Call Show_Error("The Polanyi parameters have not been properly " & _
        "specified.  To properly specify the Polanyi parameters, you must " & _
        "enter a non-zero value for each of the following parameters: " & _
        "W0, BB, and GM (Polanyi Exponent).  " & _
        "Click on the button marked Edit Parameters to " & _
        "make these changes.")
    Exit Sub
  End If
  Select Case cboMethod.ItemData(cboMethod.ListIndex)
    Case IPESMETHOD_LIQ_3PARAM:
      WhichModule = 1
    Case IPESMETHOD_GAS_DRSPREADINGP:
      WhichModule = 4
    Case IPESMETHOD_LIQ_DRUNIFORM:
      WhichModule = 5
    Case Else:
      Call Show_Error("IPE calculation code #" & _
          Trim$(Str$(cboMethod.ItemData(cboMethod.ListIndex))) & _
          " is invalid.  Select another method.")
      Exit Sub
  End Select
  INPUT_NL = CInt(Component(0).IPES_NumRegressionPts)
  INPUT_OMAG = CDbl(Component(0).IPES_OrderOfMagnitude)
  Call ModelIPE_Go(WhichModule, _
      INPUT_NL, INPUT_OMAG, Raise_Dirty_Flag)
  If (Raise_Dirty_Flag) Then
    'THROW DIRTY FLAG.
    Call frmFreundlich_DirtyStatus_Throw
  End If
  'REFRESH WINDOW.
  Call frmFreundlich_Refresh
End Sub


Private Sub cmdCancelOK_Click(Index As Integer)
Dim WhichSelected As Integer
  Select Case Index
    Case 0:     'CANCEL.
      'ROLLBACK TO ORIGINAL COMPONENT DATA.
      Component(0) = SaveOldComponent
      'EXIT OUT OF HERE.
      USER_HIT_CANCEL = True
      USER_HIT_OK = False
      Unload Me
      Exit Sub
    Case 1:     'OK.
      'VERIFY THAT NEW K AND 1/N SOURCE IS VALID.
      If (optFreundlichSource(0)) Then WhichSelected = 0
      If (optFreundlichSource(1)) Then WhichSelected = 1
      If (optFreundlichSource(2)) Then WhichSelected = 2
      Select Case WhichSelected
        Case 0:
          If (Left$(optFreundlichSource(0).Caption, 1) = "(") Then
            'FORCE SOURCE TO USER-INPUT.
            Call Show_Error("Unable to validate isotherm " & _
              "database as source of K and 1/n: reverting " & _
              "to user-input as source of K and 1/n.")
            Component(0).Source_KandOneOverN = KNSOURCE_USERINPUT
          End If
        Case 1:
          If (Left$(optFreundlichSource(1).Caption, 1) = "(") Then
            'FORCE SOURCE TO USER-INPUT.
            Call Show_Error("Unable to validate IPES as " & _
              "source of K and 1/n: reverting to user-input " & _
              "as source of K and 1/n.")
            Component(0).Source_KandOneOverN = KNSOURCE_USERINPUT
          End If
        Case 2:
          'DO NOTHING!
      End Select
      'TRANSFER K AND 1/N DATA TO "USED" VARIABLES IN COMPONENT STRUCTURE.
      Select Case Component(0).Source_KandOneOverN
        Case KNSOURCE_ISOTHERMDB
          'ISOTHERM DATABASE.
          Component(0).Use_OneOverN = Component(0).IsothermDB_OneOverN
          Component(0).Use_K = Component(0).IsothermDB_K
        Case KNSOURCE_IPES
          'IPE CALCULATION.
          Component(0).Use_OneOverN = Component(0).IPESResult_OneOverN
          Component(0).Use_K = Component(0).IPESResult_K
        Case KNSOURCE_USERINPUT
          'USER INPUT.
          Component(0).Use_OneOverN = Component(0).UserEntered_OneOverN
          Component(0).Use_K = Component(0).UserEntered_K
      End Select
      'EXIT OUT OF HERE.
      USER_HIT_CANCEL = False
      USER_HIT_OK = True
      Unload Me
      Exit Sub
  End Select
End Sub


Private Sub cmdEditPolanyi_Click()
Dim Raise_Dirty_Flag As Boolean
  Call frmPolanyi.frmPolanyi_Edit(Me, Raise_Dirty_Flag)
  If (Raise_Dirty_Flag) Then
    'DO NOT THROW FREUNDLICH DIRTY FLAG; THROW MAIN DIRTY FLAG.
    'REASON: THE POLANYI PARAMETERS ARE _NOT_ SPECIFIC
    'TO THIS COMPONENT; THEY ARE SPECIFIC TO THE CARBON.
    '---------------------------------
    ''THROW DIRTY FLAG.
    'Call frmFreundlich_DirtyStatus_Throw
    'THROW (MAIN) DIRTY FLAG.
    Call DirtyStatus_Throw
  End If
End Sub


Private Sub cmdFind_Click(Index As Integer)
Dim NewName As String
Dim USER_HIT_CANCEL As Boolean
  Select Case Index
    Case 0:     'FIND.
      NewName = Find_String
      Do While (1 = 1)
        NewName = frmNewName.frmNewName_GetName( _
            "Search for String", _
            "Enter the string to find:", _
            NewName, _
            USER_HIT_CANCEL)
        If (USER_HIT_CANCEL) Then Exit Sub
        NewName = Trim$(NewName)
        If (NewName <> "") Then Exit Do
        Call Show_Error("You may only enter a non-blank search string.")
      Loop
      Find_String = NewName
      Call Do_Search_For_Text(True)
    Case 1:     'FIND AGAIN.
      Call Do_Search_For_Text(True)
  End Select
End Sub


Private Sub cmdSelect_Click()
  Call lstCompo_Click
End Sub


Private Sub Command2_Click()
    Set Picture1.Picture = CaptureActiveWindow()
    PrintPictureToFitPage Printer, Picture1.Picture
    Printer.EndDoc
    ' Set focus back to form.
    Me.SetFocus
End Sub


Private Sub Form_Activate()
Dim No_IsothermDBData As Boolean
  If (Not ActivatedYet) Then
    ActivatedYet = True
    'RE-POPULATE COMPONENT LIST FROM ISOTHERM DATABASE.
    Call populate_lstCompo
    'SETUP SOURCE OF K AND 1/N.
    ''''Debug.Print "xxxx1 " & Now
    HALT_OPTFREUNDLICHSOURCE = True
    Select Case Component(0).Source_KandOneOverN
      Case KNSOURCE_ISOTHERMDB: optFreundlichSource(0).Value = True
      Case KNSOURCE_IPES: optFreundlichSource(1).Value = True
      Case KNSOURCE_USERINPUT: optFreundlichSource(2).Value = True
    End Select
    HALT_OPTFREUNDLICHSOURCE = False
    '
    'OUTLINE OF SEARCHES:
    '====================
    '1.) If user has already selected an isotherm from the DB,
    '    select that one on-screen, along with their selected
    '    pH/conc. range.
    '2.) Otherwise, if they have a CAS# for this component,
    '    search for the CAS#.  If found, select it on-screen.
    '3.) If the above two criteria fail, search for an exact
    '    match on the component name.  If found, select
    '    it on-screen.
    '
    'SEARCH #1:
    '==========
    No_IsothermDBData = True
    If (Trim$(Component(0).IsothermDB_Component_Name) <> "") Then
      '----- Find the selected component: -----
      Find_String = Trim$(Component(0).IsothermDB_Component_Name)
      If (Do_Search_For_Text(False) = True) Then
        'SEARCH SUCEEDED; SELECT THIS ISOTHERM.
        'NOTE: BY HALTING lstCompo_Click(), THE COMPOUND CAN
        'BE SELECTED WITHOUT CLEARING THE ISOTHERM DB VALUES
        'OF K AND 1/N.
        'Call cmdSelect_Click
        HALT_LSTCOMPO = True
        Call Do_Select_Component(lstCompo.ListIndex)
        HALT_LSTCOMPO = False
        If (Component(0).IsothermDB_K <> -1#) And _
            (Component(0).IsothermDB_OneOverN <> -1#) Then
          '----- Find the selected pH/conc. range:
          If (Component(0).IsothermDB_Range_Num <> -1) Then
            'JUST BEING PARANOID ABOUT DATABASE CHANGES.
            If (Component(0).IsothermDB_Range_Num <= lstRange(0).ListCount - 1) Then
              HALT_LSTRANGE = True
              lstRange(0).ListIndex = Component(0).IsothermDB_Range_Num
              lstRange(1).ListIndex = Component(0).IsothermDB_Range_Num
              HALT_LSTRANGE = False
              'Call lstRange_Click(0)
              Call populate_lblValue(lstRange(0).ItemData(lstRange(0).ListIndex))
              No_IsothermDBData = False
            Else
              HALT_LSTRANGE = True
              lstRange(0).ListIndex = 0
              lstRange(1).ListIndex = 0
              HALT_LSTRANGE = False
              'Call lstRange_Click(0)
              Call populate_lblValue(lstRange(0).ItemData(lstRange(0).ListIndex))
              No_IsothermDBData = False
            End If
          End If
        End If
      End If
    End If
    '
    'SEARCH #2:
    '==========
    If (No_IsothermDBData) Then
      If (Component(0).CAS <> 0) Then
        Find_String = Trim$(Str$(Component(0).CAS)) & "   "
        If (Do_Search_For_Text(False) = True) Then
          'SEARCH SUCEEDED; SELECT THIS ISOTHERM.
          Call cmdSelect_Click
          No_IsothermDBData = False
        End If
      End If
    End If
    '
    'SEARCH #3:
    '==========
    If (No_IsothermDBData) Then
      Find_String = "   " & Trim$(Component(0).Name)
      If (Do_Search_For_Text(False) = True) Then
        'SEARCH SUCEEDED; SELECT THIS ISOTHERM.
        Call cmdSelect_Click
        No_IsothermDBData = False
      End If
    End If
    ''
    ''NONE OF THE ABOVE SUCCEEDED; K AND 1/N UNSPECIFIED SO FAR.
    ''
    'If (No_IsothermDBData) Then
    '  Component(0).IsothermDB_K = -1#
    '  Component(0).IsothermDB_OneOverN = -1#
    'End If
    '
    'SOME OF THE SELECTIONS ABOVE MAY HAVE SET THE DIRTY FLAG.
    'THIS CODE CLEARS IT.
    '
    Call frmFreundlich_DirtyStatus_Clear
    'REFRESH DISPLAY.
    Call frmFreundlich_Refresh
    optFreundlichSource(0).Enabled = True
    optFreundlichSource(1).Enabled = True
    optFreundlichSource(2).Enabled = True
    '
    'CLEAR HOURGLASS MOUSE POINTER.
    '
    Screen.MousePointer = 0
  End If
End Sub
Private Sub Form_Load()
  'SAVE OLD COMPONENT FOR CANCEL ROLLBACK.
  SaveOldComponent = Component(0)
  'MISC INITS.
  ActivatedYet = False
  Me.Height = 7200
  Me.Width = 9585
  Call CenterOnForm(Me, frmMain)
  Find_String = ""
  lblWarning.Caption = ""
  Call Populate_cboMethod
  Call Populate_cboSortMethod
  Me.Caption = "Freundlich Isotherm Parameters for " & Trim$(Component(0).Name)
  If (Bed.Phase = 0) Then
    fraIsothermDB.Caption = "Liquid Phase Isotherm Database"
    lblEstimationMethod.Caption = "Liquid Phase Estimation Method:"
  Else
    fraIsothermDB.Caption = "Gas Phase Isotherm Database"
    lblEstimationMethod.Caption = "Gas Phase Estimation Method:"
  End If
  Call Clear_lblValue
  Call frmFreundlich_DirtyStatus_Clear
  Call frmFreundlich_GenericStatus_Set("")
  'POPULATE UNIT CONTROLS.
  Call frmFreundlich_PopulateUnits
  'DEMO SETTINGS.
  Call LOCAL___Reset_DemoVersionDisablings
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Call unitsys_unregister_all_on_form(Me)
End Sub


Sub Do_Select_Component(WhichComp As Integer)
''''Dim THIS_ITEMDATA As Long
Dim ThisText As String
Dim ThisCAS As String
Dim ThisChemical As String
  If (WhichComp < 0) Or (lstCompo.ListCount <= 0) Then
    Exit Sub
  End If
  lstCompo.ListIndex = WhichComp
  ''''THIS_ITEMDATA = lstCompo.ItemData(lstCompo.ListIndex)
  'EXTRACT CAS NUMBER AND COMPONENT NAME.
  ThisText = lstCompo.List(WhichComp)
  ThisCAS = Trim$(Left$(ThisText, 15))
  ThisChemical = Trim$(Mid$(ThisText, 16, Len(ThisText) - 15))
  Call populate_lstRange(ThisCAS, ThisChemical)
End Sub
Private Sub lstCompo_Click()
  If (HALT_LSTCOMPO) Then Exit Sub
  HALT_LSTCOMPO = True
  Call Do_Select_Component(lstCompo.ListIndex)
  HALT_LSTCOMPO = False
  'CLEAR EXISTING RECORD DATA.
  Call Clear_lblValue
  'INVALIDATE EXISTING ISOTHERM RECORD LINK (IF ANY).
  Component(0).IsothermDB_OneOverN = -1#
  Component(0).IsothermDB_K = -1#
  HALT_LSTRANGE = True
  lstRange(0).ListIndex = -1
  lstRange(1).ListIndex = -1
  HALT_LSTRANGE = False
  Call frmFreundlich_Refresh
End Sub


Private Sub lstRange_Click(Index As Integer)
  If (HALT_LSTRANGE) Then Exit Sub
  HALT_LSTRANGE = True
  'KEEP THE RANGE LISTBOXES IN SYNCH.
  Select Case Index
    Case 0: lstRange(1).ListIndex = lstRange(0).ListIndex
    Case 1: lstRange(0).ListIndex = lstRange(1).ListIndex
  End Select
  'TRANSFER LINK TO COMPONENT(0) STRUCTURE.
  Component(0).IsothermDB_Component_Name = Trim$(lstCompo.List(lstCompo.ListIndex))
  Component(0).IsothermDB_Range_Num = lstRange(0).ListIndex
  'DISPLAY ISOTHERM RECORD.
  Call populate_lblValue(lstRange(0).ItemData(lstRange(0).ListIndex))
  'THROW DIRTY FLAG.
  Call frmFreundlich_DirtyStatus_Throw
  'REFRESH WINDOW.
  Call frmFreundlich_Refresh
  HALT_LSTRANGE = False
End Sub


Private Sub optFreundlichSource_Click(Index As Integer, Value As Integer)
Dim KandOneOverN_Enabled As Integer
Dim X As Integer
Dim temp As String
Dim WhichSelected As Integer
  If (HALT_OPTFREUNDLICHSOURCE) Then Exit Sub
  'DETERMINE WHICH OPTION WAS SELECTED.
  If (optFreundlichSource(0)) Then WhichSelected = 0
  If (optFreundlichSource(1)) Then WhichSelected = 1
  If (optFreundlichSource(2)) Then WhichSelected = 2
'Debug.Print "optFreundlichSource_Click; WhichSelected = " & _
Trim$(Str$(WhichSelected))
  'TRANSFER K AND 1/N TO "USED" VARIABLES IN COMPONENT STRUCTURE.
  Select Case WhichSelected
    Case 0
      Component(0).Source_KandOneOverN = KNSOURCE_ISOTHERMDB
      Component(0).Use_K = Component(0).IsothermDB_K
      Component(0).Use_OneOverN = Component(0).IsothermDB_OneOverN
      KandOneOverN_Enabled = False
    Case 1
      Component(0).Source_KandOneOverN = KNSOURCE_IPES
      Component(0).Use_K = Component(0).IPESResult_K
      Component(0).Use_OneOverN = Component(0).IPESResult_OneOverN
      KandOneOverN_Enabled = False
    Case 2
      Component(0).Source_KandOneOverN = KNSOURCE_USERINPUT
      Component(0).Use_K = Component(0).UserEntered_K
      Component(0).Use_OneOverN = Component(0).UserEntered_OneOverN
      KandOneOverN_Enabled = True
  End Select
  'REFRESH WINDOW.
  Call frmFreundlich_Refresh
End Sub



Private Sub UCtl_GotFocus(Ctl As Control)
Dim StatusMessagePanel As String
Dim CtlIndex As Integer
  Call unitsys_control_txtx_gotfocus(Ctl)
  CtlIndex = 0
  On Error Resume Next
  CtlIndex = Ctl.Index
  On Error GoTo 0
  If (Trim$(UCase$(Ctl.Name)) = Trim$(UCase$("txtInput"))) And _
      (CtlIndex = 11) Then
    StatusMessagePanel = "Type in the order of magnitude"
  ElseIf (Trim$(UCase$(Ctl.Name)) = Trim$(UCase$("txtInput"))) And _
      (CtlIndex = 12) Then
    StatusMessagePanel = "Type in the number of regression points"
  ElseIf (Trim$(UCase$(Ctl.Name)) = Trim$(UCase$("UserOneOverN"))) Then
    StatusMessagePanel = "Type in the user-input Freundlich K value"
  ElseIf (Trim$(UCase$(Ctl.Name)) = Trim$(UCase$("UserK"))) Then
    StatusMessagePanel = "Type in the user-input Freundlich 1/n value"
  Else
    'NOT RECOGNIZED -- DO NOTHING.
  End If
  Call frmFreundlich_GenericStatus_Set(StatusMessagePanel)
End Sub
Sub UCtl_LostFocus(Ctl As Control)
Dim NewValue_Okay As Integer
Dim NewValue As Double
Dim Val_Low As Double
Dim Val_High As Double
Dim Raise_Dirty_Flag As Boolean
Dim CtlIndex As Integer
  'NOTE: LOW AND HIGH VALUES IN BASE UNITS.
  CtlIndex = 0
  On Error Resume Next
  CtlIndex = Ctl.Index
  On Error GoTo 0
  If (Trim$(UCase$(Ctl.Name)) = Trim$(UCase$("txtInput"))) And _
      (CtlIndex = 11) Then
    Val_Low = 1#: Val_High = 10#
  ElseIf (Trim$(UCase$(Ctl.Name)) = Trim$(UCase$("txtInput"))) And _
      (CtlIndex = 12) Then
    Val_Low = 1#: Val_High = 10000#
  ElseIf (Trim$(UCase$(Ctl.Name)) = Trim$(UCase$("UserOneOverN"))) Then
    Val_Low = 1E-40: Val_High = 1E+40
  ElseIf (Trim$(UCase$(Ctl.Name)) = Trim$(UCase$("UserK"))) Then
    Val_Low = 1E-40: Val_High = 1E+40
  Else
    'NOT RECOGNIZED -- DO NOTHING.
  End If
  NewValue_Okay = False
  If (unitsys_control_txtx_lostfocus_validate(Ctl, Val_Low, Val_High, NewValue, Raise_Dirty_Flag)) Then
    NewValue_Okay = True
  End If
  Call unitsys_control_txtx_lostfocus(Ctl, NewValue)
  Call frmFreundlich_GenericStatus_Set("")
  If (NewValue_Okay) Then
    If (Raise_Dirty_Flag) Then
      'STORE TO MEMORY.
      If (Trim$(UCase$(Ctl.Name)) = Trim$(UCase$("txtInput"))) And _
          (CtlIndex = 11) Then
        Component(0).IPES_OrderOfMagnitude = NewValue
      ElseIf (Trim$(UCase$(Ctl.Name)) = Trim$(UCase$("txtInput"))) And _
          (CtlIndex = 12) Then
        Component(0).IPES_NumRegressionPts = CInt(NewValue)
      ElseIf (Trim$(UCase$(Ctl.Name)) = Trim$(UCase$("UserOneOverN"))) Then
        Component(0).UserEntered_OneOverN = NewValue
      ElseIf (Trim$(UCase$(Ctl.Name)) = Trim$(UCase$("UserK"))) Then
        Component(0).UserEntered_K = NewValue
      Else
        'NOT RECOGNIZED -- DO NOTHING.
      End If
      'RAISE DIRTY FLAG IF NECESSARY.
      If (Raise_Dirty_Flag) Then
        'THROW DIRTY FLAG.
        Call frmFreundlich_DirtyStatus_Throw
      End If
      'REFRESH WINDOW.
      Call frmFreundlich_Refresh
    End If
  End If
End Sub


Private Sub txtInput_GotFocus(Index As Integer)
  Dim Ctl As Control: Set Ctl = txtInput(Index): Call UCtl_GotFocus(Ctl)
End Sub
Private Sub txtInput_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = Global_NumericKeyPress(KeyAscii)
End Sub
Private Sub txtInput_LostFocus(Index As Integer)
  Dim Ctl As Control: Set Ctl = txtInput(Index): Call UCtl_LostFocus(Ctl)
End Sub

Private Sub UserK_GotFocus()
  Dim Ctl As Control: Set Ctl = UserK: Call UCtl_GotFocus(Ctl)
End Sub
Private Sub UserK_KeyPress(KeyAscii As Integer)
  KeyAscii = Global_NumericKeyPress(KeyAscii)
End Sub
Private Sub UserK_LostFocus()
  Dim Ctl As Control: Set Ctl = UserK: Call UCtl_LostFocus(Ctl)
End Sub

Private Sub UserOneOverN_GotFocus()
  Dim Ctl As Control: Set Ctl = UserOneOverN: Call UCtl_GotFocus(Ctl)
End Sub
Private Sub UserOneOverN_KeyPress(KeyAscii As Integer)
  KeyAscii = Global_NumericKeyPress(KeyAscii)
End Sub
Private Sub UserOneOverN_LostFocus()
  Dim Ctl As Control: Set Ctl = UserOneOverN: Call UCtl_LostFocus(Ctl)
End Sub




