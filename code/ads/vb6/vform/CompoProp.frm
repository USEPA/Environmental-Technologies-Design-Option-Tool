VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmCompoProp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Component Properties"
   ClientHeight    =   6480
   ClientLeft      =   450
   ClientTop       =   1050
   ClientWidth     =   8895
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   8895
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
      Left            =   500
      TabIndex        =   55
      Top             =   4440
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   9240
      ScaleHeight     =   795
      ScaleWidth      =   1275
      TabIndex        =   54
      Top             =   3960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin Threed.SSFrame SSFrame3 
      Height          =   1575
      Left            =   2430
      TabIndex        =   46
      Top             =   4410
      Width           =   6315
      _Version        =   65536
      _ExtentX        =   11139
      _ExtentY        =   2778
      _StockProps     =   14
      Caption         =   "Freundlich Isotherm Parameters"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtDataComponentProperty 
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
         Height          =   288
         Index           =   6
         Left            =   2340
         TabIndex        =   11
         Text            =   "element 6"
         Top             =   660
         Width           =   1212
      End
      Begin VB.TextBox txtDataComponentProperty 
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
         Height          =   288
         Index           =   5
         Left            =   2340
         TabIndex        =   10
         Text            =   "element 5"
         Top             =   300
         Width           =   1212
      End
      Begin VB.ComboBox cboSource 
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
         Left            =   2340
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1020
         Width           =   3915
      End
      Begin VB.ComboBox txtPropUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Index           =   5
         Left            =   3660
         Style           =   2  'Dropdown List
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   270
         Width           =   2595
      End
      Begin VB.Label lblComponentProperty 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Freundlich 1/n"
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
         Left            =   810
         TabIndex        =   50
         Top             =   690
         Width           =   1455
      End
      Begin VB.Label lblComponentProperty 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Freundlich K"
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
         Left            =   810
         TabIndex        =   49
         Top             =   330
         Width           =   1455
      End
      Begin VB.Label lblComponentProperty 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Source of K and 1/n"
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
         Left            =   390
         TabIndex        =   48
         Top             =   1080
         Width           =   1875
      End
   End
   Begin Threed.SSFrame ssframe_StEPP 
      Height          =   1905
      Left            =   30
      TabIndex        =   42
      Top             =   270
      Width           =   2355
      _Version        =   65536
      _ExtentX        =   4154
      _ExtentY        =   3360
      _StockProps     =   14
      Caption         =   "StEPP Link"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSCommand cmdImportClipboard 
         Height          =   525
         Left            =   90
         TabIndex        =   43
         TabStop         =   0   'False
         ToolTipText     =   "Click here to import properties from a StEPP clipboard transfer"
         Top             =   1260
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   926
         _StockProps     =   78
         Caption         =   "Clip&board"
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
      Begin Threed.SSCommand cmdImportFromFile 
         Height          =   525
         Left            =   90
         TabIndex        =   44
         TabStop         =   0   'False
         ToolTipText     =   "Click here to import properties from a StEPP export file"
         Top             =   750
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   926
         _StockProps     =   78
         Caption         =   "StEPP &Export File"
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
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Obtain properties from StEPP via:"
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
         Height          =   380
         Left            =   90
         TabIndex        =   45
         Top             =   270
         Width           =   2100
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   2865
      Left            =   9030
      TabIndex        =   37
      Top             =   600
      Visible         =   0   'False
      Width           =   3255
      _Version        =   65536
      _ExtentX        =   5741
      _ExtentY        =   5054
      _StockProps     =   14
      Caption         =   "Unused -- Invisible"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmdCancelOld 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
         Height          =   495
         Left            =   540
         TabIndex        =   41
         Top             =   1500
         Width           =   2175
      End
      Begin Threed.SSCommand cmdCaData 
         Height          =   345
         Left            =   1650
         TabIndex        =   38
         Top             =   540
         Visible         =   0   'False
         Width           =   1170
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "&Isotherms"
         ForeColor       =   255
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
      Begin Threed.SSCommand cmdIPES 
         Height          =   375
         Left            =   1620
         TabIndex        =   39
         Top             =   900
         Visible         =   0   'False
         Width           =   1170
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "IP&ES"
         ForeColor       =   255
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
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   " * K is in (mg/g)*(L/mg)^(1/n)"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Visible         =   0   'False
         Width           =   2655
      End
   End
   Begin VB.TextBox txtDataComponentProperty 
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
      Index           =   0
      Left            =   4770
      TabIndex        =   0
      Text            =   "element 0"
      Top             =   690
      Width           =   2895
   End
   Begin VB.TextBox txtDataComponentProperty 
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
      Height          =   288
      Index           =   1
      Left            =   4770
      TabIndex        =   1
      Text            =   "element 1"
      Top             =   1050
      Width           =   1212
   End
   Begin VB.TextBox txtDataComponentProperty 
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
      Height          =   288
      Index           =   2
      Left            =   4770
      TabIndex        =   2
      Text            =   "element 2"
      Top             =   1410
      Width           =   1212
   End
   Begin VB.TextBox txtDataComponentProperty 
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
      Height          =   288
      Index           =   4
      Left            =   4770
      TabIndex        =   4
      Text            =   "element 4"
      Top             =   2130
      Width           =   1212
   End
   Begin VB.TextBox txtDataComponentProperty 
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
      Height          =   288
      Index           =   3
      Left            =   4770
      TabIndex        =   3
      Text            =   "element 3"
      Top             =   1770
      Width           =   1212
   End
   Begin VB.TextBox txtDataComponentProperty 
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
      Height          =   288
      Index           =   7
      Left            =   4770
      TabIndex        =   7
      Text            =   "element 7"
      Top             =   3210
      Width           =   1212
   End
   Begin VB.TextBox txtDataComponentProperty 
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
      Height          =   288
      Index           =   8
      Left            =   4770
      TabIndex        =   8
      Text            =   "element 8"
      Top             =   3570
      Width           =   1212
   End
   Begin VB.TextBox txtDataComponentProperty 
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
      Height          =   288
      Index           =   9
      Left            =   4770
      TabIndex        =   6
      Text            =   "element 9"
      Top             =   2850
      Width           =   1212
   End
   Begin VB.TextBox txtDataComponentProperty 
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
      Height          =   288
      Index           =   10
      Left            =   4770
      TabIndex        =   5
      Text            =   "element 10"
      Top             =   2490
      Width           =   1212
   End
   Begin VB.ComboBox txtPropUnits 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Index           =   1
      Left            =   6030
      Style           =   2  'Dropdown List
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1020
      Width           =   1275
   End
   Begin VB.ComboBox txtPropUnits 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Index           =   2
      Left            =   6030
      Style           =   2  'Dropdown List
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1380
      Width           =   1275
   End
   Begin VB.ComboBox txtPropUnits 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Index           =   7
      Left            =   6030
      Style           =   2  'Dropdown List
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3180
      Width           =   1275
   End
   Begin VB.ComboBox txtPropUnits 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Index           =   3
      Left            =   6030
      Style           =   2  'Dropdown List
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1740
      Width           =   1275
   End
   Begin VB.ComboBox txtPropUnits 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Index           =   4
      Left            =   6030
      Style           =   2  'Dropdown List
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2100
      Width           =   1275
   End
   Begin VB.ComboBox txtPropUnits 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Index           =   10
      Left            =   6030
      Style           =   2  'Dropdown List
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2460
      Width           =   1275
   End
   Begin VB.ComboBox txtPropUnits 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Index           =   9
      Left            =   6030
      Style           =   2  'Dropdown List
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2820
      Width           =   1275
   End
   Begin VB.ComboBox cboChemName 
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
      Left            =   4770
      Style           =   2  'Dropdown List
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   30
      Width           =   2895
   End
   Begin VB.TextBox txtDataComponentProperty 
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
      Height          =   288
      Index           =   11
      Left            =   4770
      TabIndex        =   9
      Text            =   "element 11"
      Top             =   3930
      Width           =   1212
   End
   Begin Threed.SSCommand cmdFreundlich 
      Height          =   855
      Left            =   120
      TabIndex        =   21
      TabStop         =   0   'False
      ToolTipText     =   "Click here to edit the Freundlich K and 1/n values for this component"
      Top             =   3450
      Width           =   2175
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "&Freundlich K and 1/n"
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
      Index           =   1
      Left            =   120
      TabIndex        =   22
      TabStop         =   0   'False
      ToolTipText     =   "Click here to save the changes to the component(s) you have edited"
      Top             =   5490
      Width           =   2175
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
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
   Begin Threed.SSCommand cmdKinetics 
      Height          =   855
      Left            =   120
      TabIndex        =   23
      TabStop         =   0   'False
      ToolTipText     =   "Click here to edit the kinetic parameters for this component"
      Top             =   2610
      Width           =   2175
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "&Kinetics"
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
      Left            =   120
      TabIndex        =   36
      TabStop         =   0   'False
      ToolTipText     =   "Click here to abandon any changes you have made to this component"
      Top             =   5010
      Width           =   2175
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
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
   Begin Threed.SSPanel SSPanel2 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   51
      Top             =   6075
      Width           =   8895
      _Version        =   65536
      _ExtentX        =   15690
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
         TabIndex        =   52
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
         TabIndex        =   53
         Top             =   60
         Width           =   6520
         _Version        =   65536
         _ExtentX        =   11501
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
   Begin VB.Label lblComponentProperty 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
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
      Left            =   2640
      TabIndex        =   35
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label lblComponentProperty 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Molecular Weight"
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
      Left            =   2640
      TabIndex        =   34
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label lblComponentProperty 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Molar Volume @ NBP"
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
      Left            =   2640
      TabIndex        =   33
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label lblComponentProperty 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Initial Concentration"
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
      Left            =   2640
      TabIndex        =   32
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label lblComponentProperty 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Boiling Point"
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
      Left            =   2640
      TabIndex        =   31
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label lblComponentProperty 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Vapor Pressure"
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
      Left            =   2640
      TabIndex        =   30
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label lblUnit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "-"
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
      Left            =   6030
      TabIndex        =   29
      Top             =   3600
      Width           =   1275
   End
   Begin VB.Label lblComponentProperty 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Refractive Index"
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
      Left            =   2640
      TabIndex        =   28
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label lblComponentProperty 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Solubility"
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
      Left            =   2640
      TabIndex        =   27
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label lblComponentProperty 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Liquid Density"
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
      Left            =   2640
      TabIndex        =   26
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label lblComponentProperty 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CAS Number"
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
      Index           =   12
      Left            =   2640
      TabIndex        =   25
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label lblUnit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "-"
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
      Left            =   6030
      TabIndex        =   24
      Top             =   3960
      Width           =   1275
   End
End
Attribute VB_Name = "frmCompoProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim FORM_MODE As Integer
Const FORM_MODE_ADDNEW = 1
Const FORM_MODE_EDIT = 2
Dim USER_HIT_CANCEL As Boolean
Dim USER_HIT_OK As Boolean
Dim frmCompoProp_Is_Dirty As Boolean

Dim START_AT_COMPNUMBER As Integer
Dim CurrentCompNumber As Integer
Dim TempComponents(1 To Number_Compo_Max) As ComponentPropertyType
        'USED TO STORE TEMPORARY COPIES OF COMPONENTS.
        'THIS ALLOWS A ROLLBACK IF THE USER HITS CANCEL.

Dim HALT_CBOCHEMNAME As Boolean
Dim HALT_CBOSOURCE As Boolean





Const frmCompoProp_declarations_end = True


Sub frmCompoProp_Add( _
    OUTPUT_Raise_Dirty_Flag As Boolean)
  FORM_MODE = FORM_MODE_ADDNEW
  frmCompoProp.Show 1
  If (USER_HIT_OK) Then
    OUTPUT_Raise_Dirty_Flag = True
  Else
    OUTPUT_Raise_Dirty_Flag = False
  End If
End Sub
Sub frmCompoProp_Edit( _
    OUTPUT_Raise_Dirty_Flag As Boolean, _
    INPUT_Start_At_CompNumber)
  FORM_MODE = FORM_MODE_EDIT
  START_AT_COMPNUMBER = INPUT_Start_At_CompNumber
  frmCompoProp.Show 1
  If (USER_HIT_OK) Then
    OUTPUT_Raise_Dirty_Flag = True
  Else
    OUTPUT_Raise_Dirty_Flag = False
  End If
End Sub
'RETURNS:
'- true = it's okay to unload now.
'- false = cancel the unload.
Function frmCompoProp_Query_Unload() As Integer
Dim RetVal As Integer
Dim msg As String
  If (Not frmCompoProp_Is_Dirty) Then
    frmCompoProp_Query_Unload = True
    Exit Function
  End If
  msg = "Are you sure you want to abandon the changes " & _
      "made to " & IIf(Number_Component = 1, "this component", "these components") & " ?" & _
      vbCrLf & vbCrLf & _
      "If you want to abandon the changes, click Yes." & _
      vbCrLf & _
      "If you want to save the " & _
      "changes, click No, then click OK."
  RetVal = MsgBox(msg, vbCritical + vbYesNo, _
      AppName_For_Display_Short & " : Abandon Changes ?")
  Select Case RetVal
    Case vbYes:
      frmCompoProp_Query_Unload = True
      Exit Function
    Case vbNo:
      frmCompoProp_Query_Unload = False
      Exit Function
  End Select
End Function
Sub LOCAL___Reset_DemoVersionDisablings()
  If (IsThisADemo() = True) Then
    cmdCancelOK(1).Enabled = False
    cmdImportFromFile.Enabled = False
    cmdImportClipboard.Enabled = False
  End If
End Sub


Sub Update_Display_of_KFreundlich(IndexOfChangedProperty As Integer)
''Sub Update_Display_of_KFreundlich(IndexOfChangedProperty As Integer, OldPropertyValue As Double)
''NOTE: THIS SUBROUTINE IS _NEVER_ CALLED UNLESS THE PROPERTY HAS CHANGED,
''SO THE VALUE OldPropertyValue IS NO LONGER NEEDED.
'Dim NoUnitsOnChangedProperty As Integer
'Dim OldK As String
'Dim NewK As String
'Dim KUnits As String
'Dim temp As String
'Dim CheckINI As String
'
'Dim A_Property_Changed As Integer
'Dim The_MW_Changed As Integer
'Dim Invalidate_Isotherm_K As Integer
'Dim Invalidate_IPES_K As Integer
'Dim K_Reverted As Integer
'
'Dim ReverseConversionFactor As Double
'Dim ThisValue As Double
'Dim i As Integer
'
'Dim Current_CheckValue As Integer
'
'  'THE PHILOSOPHY OF FREUNDLICH K DATA ENTRY.
'  '==========================================
'  'If the source of the K and 1/n for a chemical is:
'  '(1). User-entry.  They may do what they wish with any chemical
'  '     property.  If they change the 1/n or MW, the internally-stored
'  '     value for Freundlich K is updated so that it remains constant
'  '     on the screen.
'  '(2). IPES Routines.  They should not modify any chemical property.
'  '     If they do, a message box will pop up which tells them that
'  '     their K and 1/n from IPES have been invalidated, and their
'  '     K and 1/n will revert to the user-entered numbers.
'  '(3). Isotherm Database.  They should not modify molecular weight.
'  '     If they do, a message box will pop up which tells them that
'  '     their K and 1/n from the Isotherm Database have been
'  '     invalidated, and their K and 1/n will revert to the user-entered
'  '     numbers.
'
'  'Update display of Freundlich K, which is dependent upon MW and 1/n
'  '(depending on which units for Freundlich K have been chosen).
'  OldK = txtDataComponentProperty(5).Text
'  Call txtPropUnits_Click(5)
'  NewK = txtDataComponentProperty(5).Text
'  KUnits = txtPropUnits(5).List(txtPropUnits(5).ListIndex)
'
'  '----- INVALIDATE K AND 1/N FROM IPES OR ISOTHERM DATABASE IF NECESSARY
'  A_Property_Changed = False
'  The_MW_Changed = False
'  Invalidate_Isotherm_K = False
'  Invalidate_IPES_K = False
'  K_Reverted = False
'  If ((IndexOfChangedProperty = 1) Or (IndexOfChangedProperty = 2) Or (IndexOfChangedProperty = 3) Or (IndexOfChangedProperty = 4) Or (IndexOfChangedProperty = 7) Or (IndexOfChangedProperty = 8) Or (IndexOfChangedProperty = 9) Or (IndexOfChangedProperty = 10)) Then
'    If (OldPropertyValue <> txtDataComponentProperty(IndexOfChangedProperty).Text) Then
'      A_Property_Changed = True
'    End If
'  End If
'  If (IndexOfChangedProperty = 1) Then
'    If (OldPropertyValue <> txtDataComponentProperty(IndexOfChangedProperty).Text) Then
'      The_MW_Changed = True
'    End If
'  End If
'  If (A_Property_Changed) Then
'    'The IPES K and 1/n are invalidated if any property has changed.
'    If (Component(0).IPESResult_K <> -1) Then
'      Invalidate_IPES_K = True
'    End If
'  End If
'  If (The_MW_Changed) Then
'    'The Isotherm Database K and 1/n are invalidated if MW has changed.
'    If (Component(0).IsothermDB_K <> -1) Then
'      Invalidate_Isotherm_K = True
'    End If
'  End If
'
'  If (Invalidate_IPES_K) Then
'    '-- IPES K and 1/n are now invalid!
'    Component(0).IPESResult_K = -1
'    Component(0).IPESResult_OneOverN = -1
'    If (cboSource.ListIndex = 1) Then
'      '-- Change from IPES to user-entry.
'      cboSource.ListIndex = 2
'      K_Reverted = True
'    End If
'    Call Update_cboSource
'  End If
'  If (Invalidate_Isotherm_K) Then
'    '-- Isotherm Database K and 1/n are now invalid!
'    Component(0).IsothermDB_K = -1
'    Component(0).IsothermDB_OneOverN = -1
'    If (cboSource.ListIndex = 0) Then
'      '-- Change from IPES to user-entry.
'      cboSource.ListIndex = 2
'      K_Reverted = True
'    End If
'    Call Update_cboSource
'  End If
'
'  '-- Inform user of what the hell just happened.
'  If ((Invalidate_IPES_K) Or (Invalidate_Isotherm_K)) Then
'    temp = "A property has changed.  "
'    temp = temp & "This has caused the values of Freundlich K and 1/n from the "
'    If (Invalidate_IPES_K) Then
'      temp = temp & "IPES Results"
'    End If
'    If (Invalidate_Isotherm_K) Then
'      If (Invalidate_IPES_K) Then
'        temp = temp & " and the "
'      End If
'      temp = temp & "Isotherm Database"
'    End If
'    temp = temp & " to become invalidated for this chemical."
'    If (K_Reverted) Then
'      temp = temp & "  This chemical has reverted to user-entered K and 1/n."
'    End If
'    MsgBox temp, MB_ICONEXCLAMATION, AppName_For_Display_long
'''''''    Exit Sub
'  End If
'
'  '----- UPON CHANGE OF MW OR FREUNDLICH 1/N, UPDATE K
'  If ((IndexOfChangedProperty = 1) Or (IndexOfChangedProperty = 6)) Then
'    If (OldPropertyValue <> txtDataComponentProperty(IndexOfChangedProperty).Text) Then
'      '--- Note: Freundlich K is stored internally in (mg/g)*(L/mg)^(1/n).
'      '... In order to keep K constant in the currently-displayed set of units,
'      '... it is (sometimes) necessary to change the internally-stored value.
'      '... The user should be informed of what the heck is going on here.
'
'      'STEP ONE: Determine the new K in (mg/g)*(L/mg)^(1/n) units
'      ReverseConversionFactor = 1 / KFreundlichConversionFactor(CInt(txtPropUnits(5).ListIndex), Component(0).Use_OneOverN, Component(0).MW)
'      Component(0).Use_K = CDbl(OldK) * ReverseConversionFactor
'      'Update display of K:
'      Call txtPropUnits_Click(5)
'
'      'STEP TWO: Inform the user of what just happened
'      'frmGoAway_Caption = "Warning: Freundlich K Updated"
'      If (IndexOfChangedProperty = 1) Then
'        temp = "molecular weight"
'        NoUnitsOnChangedProperty = False
'      ElseIf (IndexOfChangedProperty = 6) Then
'        temp = "Freundlich 1/n"
'        NoUnitsOnChangedProperty = True
'      End If
'      temp = "The property of " & temp & " was manually changed to "
'      temp = temp & txtDataComponentProperty(IndexOfChangedProperty).Text
'      If (Not NoUnitsOnChangedProperty) Then
'        temp = temp & " ("
'        temp = temp & txtPropUnits(IndexOfChangedProperty).List(txtPropUnits(IndexOfChangedProperty).ListIndex)
'        temp = temp & ")"
'      End If
'      temp = temp & "." & Chr$(13)
'      temp = temp & "This causes a change in the value of Freundlich K:" & Chr$(13)
'      temp = temp & "    Old value = " & OldK & " (" & KUnits & ")" & Chr$(13)
'      temp = temp & "    New values:" & Chr$(13)
'      For i = 0 To 3
'        ThisValue = _
'            Component(0).Use_K * _
'            KFreundlichConversionFactor(i, Component(0).Use_OneOverN, Component(0).MW)
'        temp = temp & "        " & _
'            Format$(ThisValue, NumericalFormat(5)) & _
'            " (" & txtPropUnits(5).List(i) & ")" & Chr$(13)
'      Next i
'      'frmGoAway_Text = temp
'
'      'The property of molecular weight was manually changed to 156.77 mg/mmol.
'      'This causes a change in the value of Freundlich K.
'      '    Old value = ___________ (unit)
'      '    New values:
'      '        ___________ (unit)
'      '        ___________ (unit)
'      '        ___________ (unit)
'      '        ___________ (unit)
'
'      CheckINI = ini_getsetting("has_seen_freundlichK_warning")
'      If (CheckINI = "1") Then
'        Current_CheckValue = 1
'      Else
'        Current_CheckValue = 0
'      End If
'      'frmGoAway_CheckText = "Never display this warning again"
'      If (Current_CheckValue <> 1) Then
'        Call frmGoAway.frmGoAway_Run( _
'            frmCompoProp, _
'            "Warning: Freundlich K Was Updated", _
'            temp, _
'            "Never display this warning again", _
'            Current_CheckValue)
'      End If
'      Call ini_putsetting("has_seen_freundlichK_warning", Trim$(Str$(frmGoAway_CheckValue)))
'      Exit Sub
'    End If
'  End If
End Sub
Sub Update_cboSource()
  HALT_CBOSOURCE = True
  'POPULATE FREUNDLICH K AND 1/N SOURCE BOX.
  cboSource.Clear
  If (Component(0).IsothermDB_K > 0#) And (Component(0).IsothermDB_OneOverN > 0#) Then
    'ENABLE ISOTHERM DATABASE AS SOURCE.
    cboSource.AddItem "Isotherm Database"
  Else
    cboSource.AddItem "(Isotherm Database)"
  End If
  If (Component(0).IPESResult_K > 0#) And (Component(0).IPESResult_OneOverN > 0#) Then
    'ENABLE IPES AS SOURCE.
    cboSource.AddItem "Isotherm Parameter Estimation"
  Else
    cboSource.AddItem "(Isotherm Parameter Estimation)"
  End If
  cboSource.AddItem "User Entry"
  'DISPLAY CURRENT SOURCE.
  Select Case Component(0).Source_KandOneOverN
    Case KNSOURCE_ISOTHERMDB: cboSource.ListIndex = 0
    Case KNSOURCE_IPES: cboSource.ListIndex = 1
    Case KNSOURCE_USERINPUT: cboSource.ListIndex = 2
  End Select
  HALT_CBOSOURCE = False
End Sub
Sub frmCompoProp_PopulateUnits()
  'MAIN BLOCK.
  Call unitsys_register(frmCompoProp, lblComponentProperty(1), _
      txtDataComponentProperty(1), txtPropUnits(1), "molecular_weight", _
      PropertyUnits.MW, "g/gmol", "", "", 100#, True)
  Call unitsys_register(frmCompoProp, lblComponentProperty(2), _
      txtDataComponentProperty(2), txtPropUnits(2), "molar_volume", _
      PropertyUnits.MolarVolume, "mL/gmol", "", "", 100#, True)
  Call unitsys_register(frmCompoProp, lblComponentProperty(3), _
      txtDataComponentProperty(3), txtPropUnits(3), "temperature", _
      PropertyUnits.BP, "C", "", "", 100#, True)
  Call unitsys_register(frmCompoProp, lblComponentProperty(4), _
      txtDataComponentProperty(4), txtPropUnits(4), "concentration", _
      PropertyUnits.InitialConcentration, "mg/L", "", "", 100#, True)
  Call unitsys_register(frmCompoProp, lblComponentProperty(10), _
      txtDataComponentProperty(10), txtPropUnits(10), "density", _
      PropertyUnits.Liquid_Density, "g/mL", "", "", 100#, True)
  Call unitsys_register(frmCompoProp, lblComponentProperty(9), _
      txtDataComponentProperty(9), txtPropUnits(9), "concentration", _
      PropertyUnits.Aqueous_Solubility, "mg/L", "", "", 100#, True)
  Call unitsys_register(frmCompoProp, lblComponentProperty(7), _
      txtDataComponentProperty(7), txtPropUnits(7), "pressure", _
      PropertyUnits.Vapor_Pressure, "Pa", "", "", 100#, True)
  Call unitsys_register(frmCompoProp, lblComponentProperty(8), _
      txtDataComponentProperty(8), Nothing, "", _
      "", "", "", "", 100#, False)
  Call unitsys_register(frmCompoProp, lblComponentProperty(11), _
      txtDataComponentProperty(11), Nothing, "", _
      "", "", "0", "0", 100#, False)
  'FREUNDLICH K AND 1/N BLOCK.
  Call unitsys_register(frmCompoProp, lblComponentProperty(5), _
      txtDataComponentProperty(5), txtPropUnits(5), "freundlich_k", _
      PropertyUnits.k, "(mg/g)*(L/mg)^(1/n)", "", "", 100#, True)
  Call unitsys_register(frmCompoProp, lblComponentProperty(6), _
      txtDataComponentProperty(6), Nothing, "", _
      "", "", "", "", 100#, False)
End Sub
Sub Store_Unit_Settings()
  PropertyUnits.MW = unitsys_get_units(txtPropUnits(1))
  PropertyUnits.MolarVolume = unitsys_get_units(txtPropUnits(2))
  PropertyUnits.BP = unitsys_get_units(txtPropUnits(3))
  PropertyUnits.InitialConcentration = unitsys_get_units(txtPropUnits(4))
  PropertyUnits.Liquid_Density = unitsys_get_units(txtPropUnits(10))
  PropertyUnits.Aqueous_Solubility = unitsys_get_units(txtPropUnits(9))
  PropertyUnits.Vapor_Pressure = unitsys_get_units(txtPropUnits(7))
  PropertyUnits.k = unitsys_get_units(txtPropUnits(5))
End Sub


Private Sub cboChemName_Click()
  If (HALT_CBOCHEMNAME) Then Exit Sub
  TempComponents(CurrentCompNumber) = Component(0)
  CurrentCompNumber = cboChemName.ListIndex + 1
  Component(0) = TempComponents(CurrentCompNumber)
  Call frmCompoProp_Refresh
  Call Update_cboSource
End Sub
Private Sub cboSource_Click()
Dim KandOneOverN_Enabled As Integer
Dim X As Integer
Dim temp As String
  If (HALT_CBOSOURCE) Then Exit Sub
  If (Left$(cboSource.List(cboSource.ListIndex), 1) = "(") Then
    'UNABLE TO USE THAT SOURCE!
    X = cboSource.ListIndex
    cboSource.ListIndex = 2
    Select Case X
      Case 0
        temp = "You must first select an isotherm from the isotherm database.  "
        temp = temp & "Click on the button marked " & Chr$(34) & _
            "Freundlich K and 1/n" & Chr$(34) & " to do so."
        Call Show_Error(temp)
      Case 1
        temp = "You must first calculate K and 1/n using IPES.  "
        temp = temp & "To do so, click on the button marked " & Chr$(34) & _
            "Freundlich K and 1/n" & Chr$(34) & " and then click on " & _
            Chr$(34) & "Re-calculate" & Chr$(34) & " from within the next window."
        Call Show_Error(temp)
    End Select
  End If
  'UPDATE INTERNAL RECORDS.
  Select Case cboSource.ListIndex
    Case 0:       'ISOTHERM DB.
      Component(0).Source_KandOneOverN = KNSOURCE_ISOTHERMDB
      Component(0).Use_K = Component(0).IsothermDB_K
      Component(0).Use_OneOverN = Component(0).IsothermDB_OneOverN
      KandOneOverN_Enabled = False
    Case 1:       'IP ESTIMATION.
      Component(0).Source_KandOneOverN = KNSOURCE_IPES
      Component(0).Use_K = Component(0).IPESResult_K
      Component(0).Use_OneOverN = Component(0).IPESResult_OneOverN
      KandOneOverN_Enabled = False
    Case 2:       'USER INPUT.
      Component(0).Source_KandOneOverN = KNSOURCE_USERINPUT
      Component(0).Use_K = Component(0).UserEntered_K
      Component(0).Use_OneOverN = Component(0).UserEntered_OneOverN
      KandOneOverN_Enabled = True
  End Select
  'UPDATE WINDOW.
  Call txtPropUnits_Click(5)
  'txtDataComponentProperty(5) = Format$(Component(0).Use_K, "0.000")
  'txtDataComponentProperty(6) = Format$(Component(0).Use_OneOverN, "0.000")
  txtDataComponentProperty(5).Locked = Not KandOneOverN_Enabled
  txtDataComponentProperty(6).Locked = Not KandOneOverN_Enabled
  Call frmCompoProp_Refresh     'THIS CALL REDISPLAYS FREUNDLICH 1/N.
End Sub


Private Sub cmdCancelOK_Click(Index As Integer)
Dim i As Integer
  Select Case Index
    Case 0:     'CANCEL.
      If (frmCompoProp_Query_Unload() = False) Then
        'THE CANCEL WAS CANCELLED.
        Exit Sub
      End If
      USER_HIT_CANCEL = True
      USER_HIT_OK = False
      Unload Me
      Exit Sub
    Case 1:     'OK.
      'UPDATE KINETIC COEFFICIENTS.
      If Component(0).Corr(1) Then
        Component(0).kf = kf(0)
      End If
      If Component(0).Corr(2) Then
        Component(0).Ds = Ds(0)
      End If
      If Component(0).Corr(3) Then
        Component(0).Dp = Dp(0)
      End If
      If (FORM_MODE = FORM_MODE_ADDNEW) Then
        '/////////////////// ADD NEW COMPONENT CODE //////////////////////////////////////////////////////////////////////////////////////
        'ADD COMPONENT TO ROOM PROPERTIES DATA AREA.
        RoomParams.COUNT_CONTAMINANT = RoomParams.COUNT_CONTAMINANT + 1
        RoomParams.ROOM_C0(RoomParams.COUNT_CONTAMINANT) = 0#
        RoomParams.ROOM_EMIT(RoomParams.COUNT_CONTAMINANT) = 1.7
        RoomParams.ROOM_SS_VALUE(RoomParams.COUNT_CONTAMINANT) = 0#
        RoomParams.INITIAL_ROOM_CONC(RoomParams.COUNT_CONTAMINANT) = 0#
        RoomParams.RXN_RATE_CONSTANT(RoomParams.COUNT_CONTAMINANT) = 0#
        RoomParams.RXN_PRODUCT(RoomParams.COUNT_CONTAMINANT) = 0
        RoomParams.RXN_RATIO(RoomParams.COUNT_CONTAMINANT) = 0#
    
        'ADD COMPONENT TO MAIN DATA AREA.
        Number_Component = Number_Component + 1
        Component(Number_Component) = Component(0)
        'FRMMAIN.cmdViewDimensionless.Enabled = True
        'FRMMAIN.cmdEditComponent.Enabled = True
        'FRMMAIN.cmdDeleteComponent.Enabled = True
        'FRMMAIN.lstComponents.AddItem txtDataComponentProperty(0)
        'frmMain.cboSelectCompo.Enabled = True
        'frmMain.cboSelectCompo.AddItem txtDataComponentProperty(0)
        'If (Number_Component = Number_Compo_Max) Then
        '  frmMain.cmdAddComponent.Enabled = False
        'End If
        'frmMain.cboSelectCompo.ListIndex = frmMain.cboSelectCompo.ListCount - 1
        ''Update the corresponding kinetic data displayed
        'Call Update_Display_Kinetic
      Else
        '/////////////////// EDIT EXISTING COMPONENT(S) CODE //////////////////////////////////////////////////////////////////////////////////////
        TempComponents(CurrentCompNumber) = Component(0)
        For i = 1 To Number_Component
          Component(i) = TempComponents(i)
        Next i
        ''Update the display of the name
        'For N = 1 To frmMain.cboSelectCompo.ListCount
        '  frmMain.lstComponents.List(N - 1) = Component(N).name
        '  frmMain.cboSelectCompo.List(N - 1) = Component(N).name
        'Next
        'frmMain.cboSelectCompo.ListIndex = cboChemName.ListIndex
      End If
      'If (Number_Component > 0) Then
      '  frmMain.mnuRunItem(0).Enabled = True
      '  frmMain.mnuRunItem(1).Enabled = True
      '  frmMain.mnuRunItem(2).Enabled = True
      '  frmMain.mnuOptionsItem(0).Enabled = True
      '  frmMain.mnuOptionsItem(1).Enabled = True  'Variable Influent concentration
      '  frmMain.mnuOptionsItem(2).Enabled = True  'Variable Effluent concentration
      'End If
      'STORE ALL UNIT SETTINGS.
      Call Store_Unit_Settings
      'REFRESH MAIN WINDOW.
      Call frmMain_Refresh
      'EXIT OUT OF HERE.
      USER_HIT_CANCEL = False
      USER_HIT_OK = True
      Unload Me
      Exit Sub
  End Select
End Sub


Private Sub cmdFreundlich_Click()
Dim Raise_Dirty_Flag As Boolean
  Screen.MousePointer = 11
  Call frmFreundlich.frmFreundlich_Run(Raise_Dirty_Flag)
  If (Raise_Dirty_Flag) Then
    'UPDATE SOURCE OF K AND 1/N SCROLLBOX.
    Call Update_cboSource
    'THROW DIRTY FLAG.
    Call frmCompoProp_DirtyStatus_Throw
  End If
  'REFRESH VALUES, ESPECIALLY K AND 1/n.
  Call frmCompoProp_Refresh
End Sub


Private Sub cmdImportClipboard_Click()
Dim Was_Aborted As Boolean
  Call Do_ImportClipboard(Was_Aborted)
  If (Was_Aborted) Then
    Exit Sub
  Else
    'STORE ALL UNIT SETTINGS.
    Call Store_Unit_Settings
    'REFRESH MAIN WINDOW.
    Call frmMain_Refresh
    'EXIT OUT OF HERE.
    USER_HIT_CANCEL = False
    USER_HIT_OK = True
    Unload Me
    Exit Sub
  End If
End Sub
Private Sub cmdImportFromFile_Click()
Dim f As Integer
Dim LineCount As Integer
Dim ThisLine As String
Dim AllLines As String
Dim InvalidFile As Boolean
Const MAX_LINE_COUNT = 1000     'SOMEWHAT ARBITRARY.
  On Error GoTo err_cmdImportFromFile_Click
  frmMain.CommonDialog1.CancelError = True
  frmMain.CommonDialog1.DialogTitle = "Load StEPP Export File"
  frmMain.CommonDialog1.Filter = "All Files (*.*)|*.*|StEPP Export Files (*.exp)|*.exp"
  frmMain.CommonDialog1.FilterIndex = 2
  frmMain.CommonDialog1.flags = _
      cdlOFNFileMustExist + _
      cdlOFNPathMustExist
  frmMain.CommonDialog1.Action = 1
  If (frmMain.CommonDialog1.Filename = "") Then
    Exit Sub
  End If
  f = FreeFile
  LineCount = 0
  Open frmMain.CommonDialog1.Filename For Input As #f
  InvalidFile = False
  Do While (1 = 1)
    If (EOF(f)) Then Exit Do
    Line Input #f, ThisLine
    AllLines = AllLines & ThisLine & Chr$(13) & Chr$(10)
    LineCount = LineCount + 1
    If (LineCount > MAX_LINE_COUNT) Then
      InvalidFile = True
      Exit Do
    End If
  Loop
  Close #f
  If (InvalidFile) Then
    Call Show_Error("This is not a valid StEPP export file.")
    Exit Sub
  End If
  'DO THE IMPORT.
  Clipboard.SetText AllLines
  Call cmdImportClipboard_Click
  Exit Sub
exit_err_cmdImportFromFile_Click:
  Exit Sub
err_cmdImportFromFile_Click:
  If (Err.number = cdlCancel) Then
    'DO NOTHING.
  Else
    Call Show_Trapped_Error("cmdImportFromFile_Click")
  End If
  Resume exit_err_cmdImportFromFile_Click
End Sub


Private Sub cmdKinetics_Click()
Dim Raise_Dirty_Flag As Boolean
  Call frmKinetic.frmKinetic_Run(Raise_Dirty_Flag)
  If (Raise_Dirty_Flag) Then
    'THROW DIRTY FLAG.
    Call frmCompoProp_DirtyStatus_Throw
  End If
End Sub


Private Sub Command4_Click()
    Set Picture1.Picture = CaptureActiveWindow()
    PrintPictureToFitPage Printer, Picture1.Picture
    Printer.EndDoc
    ' Set focus back to form.
    Me.SetFocus
End Sub

Private Sub Form_Load()
Dim i As Integer
  'MISC INITS.
  Me.Height = 6885
  Me.Width = 8940
  Call CenterOnForm(Me, frmMain)
  'ADD/EDIT MODE RELATED.
  If (FORM_MODE = FORM_MODE_ADDNEW) Then
    'ADD MODE RELATED INITIALIZATION.
    '---- CREATE DEFAULT COMPONENT.
    '---- NOTE: GLOBAL COMPONENT(0) IS USED FOR THE CURRENT COMPONENT.
    Call SetComponentDefaults(Component(0), 0)
    '---- VARIOUS VISIBILITY SETTINGS.
    cboChemName.Visible = False
    ssframe_StEPP.Visible = True
    ''---- DATA ALREADY CONSIDERED CHANGED (DUE TO ADD MODE).
    'Call frmCompoProp_DirtyStatus_Throw
    '---- DATA UNCHANGED AS YET.
    Call frmCompoProp_DirtyStatus_Clear
  Else
    'EDIT MODE RELATED INITIALIZATION.
    '---- TRANSFER COMPONENTS TO LOCAL STORAGE.
    '---- NOTE: GLOBAL COMPONENT(0) IS USED FOR THE CURRENT COMPONENT.
    For i = 1 To Number_Component
      TempComponents(i) = Component(i)
    Next i
    Component(0) = Component(START_AT_COMPNUMBER)
    CurrentCompNumber = START_AT_COMPNUMBER
    '---- POPULATE COMPONENT NAME SCROLLBOX.
    HALT_CBOCHEMNAME = True
    cboChemName.Clear
    For i = 1 To Number_Component
      cboChemName.AddItem Trim$(Component(i).Name)
    Next i
    cboChemName.ListIndex = START_AT_COMPNUMBER - 1
    HALT_CBOCHEMNAME = False
    '---- VARIOUS VISIBILITY SETTINGS.
    cboChemName.Visible = True
    ssframe_StEPP.Visible = False
    '---- DATA UNCHANGED AS YET.
    Call frmCompoProp_DirtyStatus_Clear
  End If
  'POPULATE UNIT CONTROLS.
  Call frmCompoProp_PopulateUnits
  'REFRESH DISPLAY.
  Call frmCompoProp_Refresh
  'POPULATE SOURCE OF K AND 1/N SCROLLBOX.
  Call Update_cboSource
  'DEMO SETTINGS.
  Call LOCAL___Reset_DemoVersionDisablings
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Call unitsys_unregister_all_on_form(Me)
End Sub








Sub frmCompoProp_GenericStatus_Set(fn_Text As String)
  Me.sspanel_Status = fn_Text
End Sub
Sub frmCompoProp_DirtyStatus_Set(newVal As Boolean)
  If (newVal) Then
    frmCompoProp.sspanel_Dirty = "Data Changed"
    frmCompoProp.sspanel_Dirty.ForeColor = QBColor(12)
  Else
    frmCompoProp.sspanel_Dirty = "Unchanged"
    frmCompoProp.sspanel_Dirty.ForeColor = QBColor(0)
  End If
End Sub
Sub frmCompoProp_DirtyStatus_Set_Current()
  Call frmCompoProp_DirtyStatus_Set(frmCompoProp_Is_Dirty)
End Sub
Sub frmCompoProp_DirtyStatus_Throw()
  frmCompoProp_Is_Dirty = True
  Call frmCompoProp_DirtyStatus_Set_Current
End Sub
Sub frmCompoProp_DirtyStatus_Clear()
  frmCompoProp_Is_Dirty = False
  Call frmCompoProp_DirtyStatus_Set_Current
End Sub


Private Sub txtDataComponentProperty_GotFocus(Index As Integer)
Dim Ctl As Control
Set Ctl = txtDataComponentProperty(Index)
Dim StatusMessagePanel As String
  If (Index = 0) Then
    Call Global_GotFocus(Ctl)
  Else
    Call unitsys_control_txtx_gotfocus(Ctl)
  End If
  Select Case Index
    Case 0:
      StatusMessagePanel = "Type in the component name"
    Case 1:
      StatusMessagePanel = "Type in the molecular weight"
    Case 2:
      StatusMessagePanel = "Type in the molar volume at the normal boiling point"
    Case 3:
      StatusMessagePanel = "Type in the boiling point temperature"
    Case 4:
      StatusMessagePanel = "Type in the inlet concentration"
    Case 10:
      StatusMessagePanel = "Type in the liquid density"
    Case 9:
      StatusMessagePanel = "Type in the aqueous solubility"
    Case 7:
      StatusMessagePanel = "Type in the vapor pressure"
    Case 8:
      StatusMessagePanel = "Type in the refractive index"
    Case 11:
      StatusMessagePanel = "Type in the CAS number, with no hyphen characters"
    Case 5:
      StatusMessagePanel = "Type in the Freundlich K value"
    Case 6:
      StatusMessagePanel = "Type in the Freundlich 1/n value"
  End Select
  Call frmCompoProp_GenericStatus_Set(StatusMessagePanel)
End Sub
Private Sub txtDataComponentProperty_KeyPress(Index As Integer, KeyAscii As Integer)
  If (Index = 0) Then
    KeyAscii = Global_TextKeyPress(KeyAscii)
  Else
    KeyAscii = Global_NumericKeyPress(KeyAscii)
  End If
End Sub
Private Sub txtDataComponentProperty_LostFocus(Index As Integer)
Dim NewValue_Okay As Integer
Dim NewValue As Double
Dim Ctl As Control
Set Ctl = txtDataComponentProperty(Index)
Dim Val_Low As Double
Dim Val_High As Double
Dim Raise_Dirty_Flag As Boolean
Dim Too_Small As Integer
  'HANDLE THE COMPONENT NAME TEXTBOX.
  If (Index = 0) Then
    If (Trim$(Ctl.Text) = "") Then
      Ctl.Text = Component(0).Name
      'Call Show_Error("You must enter a non-blank string for the component name.")
      'NOTE: SHOWING THIS ERROR MESSAGE MESSES UP THE
      'SUBSEQUENT GOTFOCUS IF THE USER HITS <Enter> OR <Tab>.
    Else
      If (Trim$(Component(0).Name) <> Trim$(Ctl.Text)) Then
        Component(0).Name = Trim$(Ctl.Text)
        'THROW DIRTY FLAG.
        Call frmCompoProp_DirtyStatus_Throw
      End If
    End If
    Call Global_LostFocus(Ctl)
    Call frmCompoProp_GenericStatus_Set("")
    Exit Sub
  End If
  'NOTE: LOW AND HIGH VALUES IN BASE UNITS.
  Select Case Index
    Case 1: Val_Low = 2#: Val_High = 10000000000#
    'Case 2: Val_Low = 0.01 / 1000#: Val_High = 100000# / 1000#
    Case 2: Val_Low = 0.01: Val_High = 100000#
    Case 3: Val_Low = -273#: Val_High = 1000#
    Case 4: Val_Low = 1E-20: Val_High = 1000#
    Case 10: Val_Low = 0.001: Val_High = 100#
    Case 9: Val_Low = 0.0001: Val_High = 10000000#
    ''''Case 7: Val_Low = 0.01: Val_High = 1000000#
    Case 7: Val_Low = 0.0000000001: Val_High = 1000000#
    Case 8: Val_Low = 0.01: Val_High = 100000#
    Case 11: Val_Low = 0#: Val_High = 2000000000#
    Case 5: Val_Low = 0.0001: Val_High = 1000000#
    Case 6: Val_Low = 0.00001: Val_High = 10#
  End Select
  NewValue_Okay = False
  If (unitsys_control_txtx_lostfocus_validate(Ctl, Val_Low, Val_High, NewValue, Raise_Dirty_Flag)) Then
    NewValue_Okay = True
  End If
  Call unitsys_control_txtx_lostfocus(Ctl, NewValue)
  Call frmCompoProp_GenericStatus_Set("")
  If (NewValue_Okay) Then
    If (Raise_Dirty_Flag) Then
      'STORE TO MEMORY.
      Select Case Index
        Case 1:         'MOLECULAR WEIGHT.
          Component(0).MW = NewValue
        Case 2:         'MOLAR VOLUME.
          Component(0).MolarVolume = NewValue
        Case 3:         'BOILING POINT TEMPERATURE.
          Component(0).BP = NewValue
        Case 4:         'INLET CONCENTRATION.
          Component(0).InitialConcentration = NewValue
        Case 10:        'LIQUID DENSITY.
          Component(0).Liquid_Density = NewValue
        Case 9:         'AQUEOUS SOLUBILITY.
          Component(0).Aqueous_Solubility = NewValue
        Case 7:         'VAPOR PRESSURE.
          Component(0).Vapor_Pressure = NewValue
        Case 8:         'REFRACTIVE INDEX.
          Component(0).Refractive_Index = NewValue
        Case 11:        'CAS NUMBER.
          Component(0).CAS = NewValue
        Case 5:         'FREUNDLICH K.
          Component(0).UserEntered_K = NewValue
          Component(0).Use_K = NewValue
        Case 6:         'FREUNDLICH 1/N.
          Component(0).UserEntered_OneOverN = NewValue
          Component(0).Use_OneOverN = NewValue
      End Select
      'RAISE DIRTY FLAG IF NECESSARY.
      If (Raise_Dirty_Flag) Then
        'THROW DIRTY FLAG.
        Call frmCompoProp_DirtyStatus_Throw
        'BASED ON THIS CHANGE, UPDATE FREUNDLICH K IF NECESSARY.
        Call Update_Display_of_KFreundlich(Index)
      End If
      'REFRESH WINDOW.
      Call frmCompoProp_Refresh
    End If
  End If
End Sub


Private Sub txtPropUnits_Click(Index As Integer)
Dim Ctl As Control
Set Ctl = txtPropUnits(Index)
  Call unitsys_control_cbox_click(Ctl)
End Sub
Private Sub txtPropUnits_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = Global_TextKeyPress(KeyAscii)
End Sub




