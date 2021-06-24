VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmKinetic 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Kinetic Parameters"
   ClientHeight    =   6045
   ClientLeft      =   450
   ClientTop       =   2310
   ClientWidth     =   8415
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
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
      Left            =   6480
      TabIndex        =   45
      Top             =   5080
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   8520
      ScaleHeight     =   795
      ScaleWidth      =   1275
      TabIndex        =   44
      Top             =   4680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin Threed.SSCheck chkTortuosity_Corr 
      Height          =   285
      Left            =   1560
      TabIndex        =   43
      Top             =   4590
      Width           =   4875
      _Version        =   65536
      _ExtentX        =   8599
      _ExtentY        =   503
      _StockProps     =   78
      Caption         =   "Use Pore Diffusion Correlation for &Tortuosity"
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
   Begin Threed.SSFrame SSFrame1 
      Height          =   2325
      Left            =   8550
      TabIndex        =   24
      Top             =   1320
      Visible         =   0   'False
      Width           =   3495
      _Version        =   65536
      _ExtentX        =   6165
      _ExtentY        =   4101
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
      Begin Threed.SSOption optKF_old 
         Height          =   255
         Index           =   0
         Left            =   1680
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   1560
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
      End
      Begin Threed.SSOption optKF_old 
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   1920
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
      End
      Begin Threed.SSOption optDS_old 
         Height          =   255
         Index           =   0
         Left            =   2220
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   1500
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
      End
      Begin Threed.SSOption optDS_old 
         Height          =   255
         Index           =   1
         Left            =   2220
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   1860
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
      End
      Begin Threed.SSOption optDP_old 
         Height          =   255
         Index           =   0
         Left            =   2670
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   1530
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
      End
      Begin Threed.SSOption optDP_old 
         Height          =   255
         Index           =   1
         Left            =   2670
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   1890
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
      End
      Begin VB.Label lblDP_OLD 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   540
         TabIndex        =   27
         Top             =   810
         Width           =   1095
      End
      Begin VB.Label lblDS_OLD 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   600
         TabIndex        =   26
         Top             =   690
         Width           =   1095
      End
      Begin VB.Label lblKF_OLD 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   360
         TabIndex        =   25
         Top             =   420
         Width           =   1215
      End
   End
   Begin Threed.SSFrame fraKP 
      Height          =   2925
      Index           =   2
      Left            =   5640
      TabIndex        =   14
      Top             =   330
      Width           =   2055
      _Version        =   65536
      _ExtentX        =   3625
      _ExtentY        =   5159
      _StockProps     =   14
      Caption         =   "Pore Diffusion"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Begin VB.OptionButton optDP 
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   210
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   930
         Width           =   255
      End
      Begin VB.OptionButton optDP 
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   210
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   570
         Width           =   255
      End
      Begin VB.TextBox lblDP 
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
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   930
         Width           =   1212
      End
      Begin VB.TextBox txtDP 
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
         Left            =   600
         TabIndex        =   2
         Top             =   570
         Width           =   1215
      End
      Begin VB.Label lblUnit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "cm2/s"
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
         Left            =   480
         TabIndex        =   20
         Top             =   210
         Width           =   1095
      End
      Begin VB.Label lblCorrelationDP 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Wilke-Lee Modification of the Hirschfelder - Bird - Spotz method"
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
         Height          =   1455
         Left            =   120
         TabIndex        =   19
         Top             =   1320
         Width           =   1815
      End
   End
   Begin Threed.SSFrame fraKP 
      Height          =   2925
      Index           =   1
      Left            =   3600
      TabIndex        =   13
      Top             =   330
      Width           =   2055
      _Version        =   65536
      _ExtentX        =   3625
      _ExtentY        =   5159
      _StockProps     =   14
      Caption         =   "Surface Diffusion"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Begin VB.OptionButton optDS 
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   210
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   930
         Width           =   255
      End
      Begin VB.OptionButton optDS 
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   210
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   570
         Width           =   255
      End
      Begin VB.TextBox lblDS 
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
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   930
         Width           =   1212
      End
      Begin VB.TextBox txtDS 
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
         Left            =   600
         TabIndex        =   1
         Top             =   570
         Width           =   1215
      End
      Begin VB.Label lblUnit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "cm2/s"
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
         Left            =   480
         TabIndex        =   18
         Top             =   210
         Width           =   1095
      End
      Begin VB.Label lblCorrelationDS 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Wilke-Lee Modification of the Hirschfelder - Bird - Spotz method"
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
         Height          =   1455
         Left            =   120
         TabIndex        =   17
         Top             =   1320
         Width           =   1815
      End
   End
   Begin Threed.SSFrame fraKP 
      Height          =   2925
      Index           =   0
      Left            =   1560
      TabIndex        =   12
      Top             =   330
      Width           =   2055
      _Version        =   65536
      _ExtentX        =   3625
      _ExtentY        =   5159
      _StockProps     =   14
      Caption         =   "Film Diffusion"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Begin VB.OptionButton optKF 
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   210
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   930
         Width           =   255
      End
      Begin VB.OptionButton optKF 
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   210
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   570
         Width           =   255
      End
      Begin VB.TextBox txtKF 
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
         Left            =   600
         TabIndex        =   0
         Top             =   570
         Width           =   1212
      End
      Begin VB.TextBox lblKF 
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
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   930
         Width           =   1212
      End
      Begin VB.Label lblUnit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "cm/s"
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
         Left            =   480
         TabIndex        =   16
         Top             =   210
         Width           =   1095
      End
      Begin VB.Label lblCorrelationKF 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Wilke-Lee Modification of the Hirschfelder - Bird - Spotz method"
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
         Height          =   1485
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   1815
      End
   End
   Begin VB.TextBox txtSPDFR 
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
      Left            =   1560
      TabIndex        =   3
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox txtTort 
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
      Left            =   1560
      TabIndex        =   4
      Text            =   "txtTort"
      Top             =   3840
      Width           =   1095
   End
   Begin Threed.SSCommand cmdCancelOK 
      Height          =   495
      Index           =   1
      Left            =   1530
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Click here to save the changes you have made to the kinetic parameters on this window"
      Top             =   5010
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
   Begin Threed.SSCommand cmdCancelOK 
      Height          =   495
      Index           =   0
      Left            =   3690
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Click here to abandon any changes you have made to the kinetic parameters on this window"
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
      TabIndex        =   28
      Top             =   5640
      Width           =   8415
      _Version        =   65536
      _ExtentX        =   14843
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
         TabIndex        =   29
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
         TabIndex        =   30
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
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Correlation"
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
      Left            =   360
      TabIndex        =   9
      Top             =   1290
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "User Input"
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
      Left            =   360
      TabIndex        =   8
      Top             =   930
      Width           =   1095
   End
   Begin VB.Label lblSPDFR 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Surface To Pore Diffusion Flux Ratio (SPDFR)"
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
      Left            =   2700
      TabIndex        =   7
      Top             =   3540
      Width           =   4215
   End
   Begin VB.Label lblTort 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Tortuosity"
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
      Left            =   2700
      TabIndex        =   6
      Top             =   3900
      Width           =   1095
   End
   Begin VB.Label lblTortCorrelation 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Leave this label alone!"
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
      Height          =   555
      Left            =   1560
      TabIndex        =   5
      Top             =   4140
      Width           =   5775
   End
End
Attribute VB_Name = "frmKinetic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim USER_HIT_CANCEL As Boolean
Dim USER_HIT_OK As Boolean
Dim frmKinetic_Is_Dirty As Boolean

Dim SaveOldComponent As ComponentPropertyType




Const frmKinetic_declarations_end = True


Sub frmKinetic_Run( _
    OUTPUT_Raise_Dirty_Flag As Boolean)
  frmKinetic.Show 1
  If (USER_HIT_OK) Then
    OUTPUT_Raise_Dirty_Flag = True
  Else
    OUTPUT_Raise_Dirty_Flag = False
  End If
End Sub
Sub LOCAL___Reset_DemoVersionDisablings()
  If (IsThisADemo() = True) Then
    cmdCancelOK(1).Enabled = False
  End If
End Sub


Sub frmKinetic_GenericStatus_Set(fn_Text As String)
  Me.sspanel_Status = fn_Text
End Sub
Sub frmKinetic_DirtyStatus_Set(newVal As Boolean)
  If (newVal) Then
    frmKinetic.sspanel_Dirty = "Data Changed"
    frmKinetic.sspanel_Dirty.ForeColor = QBColor(12)
  Else
    frmKinetic.sspanel_Dirty = "Unchanged"
    frmKinetic.sspanel_Dirty.ForeColor = QBColor(0)
  End If
End Sub
Sub frmKinetic_DirtyStatus_Set_Current()
  Call frmKinetic_DirtyStatus_Set(frmKinetic_Is_Dirty)
End Sub
Sub frmKinetic_DirtyStatus_Throw()
  frmKinetic_Is_Dirty = True
  Call frmKinetic_DirtyStatus_Set_Current
End Sub
Sub frmKinetic_DirtyStatus_Clear()
  frmKinetic_Is_Dirty = False
  Call frmKinetic_DirtyStatus_Set_Current
End Sub


Sub frmKinetic_PopulateUnits()
  Call unitsys_register(frmKinetic, lblSPDFR, _
      txtKF, Nothing, "", _
      "", "", "", "", 100#, False)
  Call unitsys_register(frmKinetic, lblSPDFR, _
      txtDS, Nothing, "", _
      "", "", "", "", 100#, False)
  Call unitsys_register(frmKinetic, lblSPDFR, _
      txtDP, Nothing, "", _
      "", "", "", "", 100#, False)
  Call unitsys_register(frmKinetic, lblSPDFR, _
      txtSPDFR, Nothing, "", _
      "", "", "", "", 100#, False)
  Call unitsys_register(frmKinetic, lblTort, _
      txtTort, Nothing, "", _
      "", "", "", "", 100#, False)
End Sub


Sub Update_Ds_and_Dp_Editability()
  If (Component(0).Use_Tortuosity_Correlation) Then
    '---- Tortuosity(time) correlation is ON!
    txtTort.Enabled = False
    '---- De-enable Ds user-input
    optDS(0).Enabled = False
    optDS(0).Value = False
    optDS(1).Value = True
    Component(0).Corr(2) = True
    '---- De-enable Dp user-input
    optDP(0).Enabled = False
    optDP(0).Value = False
    optDP(1).Value = True
    Component(0).Corr(3) = True
  Else
    '---- Tortuosity(time) correlation is OFF!
    txtTort.Enabled = True
    '---- Enable Ds user-input
    optDS(0).Enabled = True
    '---- Enable Dp user-input
    optDP(0).Enabled = True
  End If
End Sub
Sub Update_Tortuosity_Display()
Dim T As Double
  If (Component(0).Use_Tortuosity_Correlation) Then
    '---- Tortuosity(time) correlation is ON!
    T = Tortuosity(0)
    txtTort = Format_It(T, 3)
    ''''chkTortuosity_Corr.Value = True
    txtTort.Enabled = False
    'lblTortCorrelation.Caption = "For t<=70 days, tortuosity = 1;" & Chr$(13) & Chr$(10) & "For t>70 days, tortuosity = 0.334 + 0.00000661*(t,minutes)"
    lblTortCorrelation.Caption = "For t<=70 days, tortuosity = 1;" & Chr$(13) & Chr$(10) & "For t>70 days, tortuosity = 0.334 + 0.009518*(t,days)"
    lblTortCorrelation.Visible = True
    lblTortCorrelation.Left = txtTort.Left
    lblTortCorrelation.Top = txtTort.Top
    txtTort.Visible = False
    lblTort.Visible = False
    Call frmKinetic_Refresh
        'THIS REDISPLAYS txtTort.
  Else
    '---- Tortuosity(time) correlation is OFF!
    T = Component(0).Tortuosity
    Call frmKinetic_Repopulate_Values     'THIS REDISPLAYS txtTort.
    'txtTort = Format_It(T, 3)
    ''''chkTortuosity_Corr.Value = False
    txtTort.Enabled = True
    lblTortCorrelation.Visible = False
    txtTort.Visible = True
    lblTort.Visible = True
  End If
End Sub


Private Sub chkTortuosity_Corr_Click(Value As Integer)
  If (Value = True) Then
    '---- Turn tortuosity(time) correlation ON!
    Component(0).Use_Tortuosity_Correlation = True
    Component(0).Constant_Tortuosity = False
    'frmprint!chkSelect(4).Enabled = True
    '---- Update SPDFR to 1.000e-30!
    Component(0).SPDFR = 1E-30
    'txtSPDFR.Text = "1.000E-30"
    'txtSPDFR = Format_It(Component(0).SPDFR, 3)
    Call frmKinetic_Refresh
        'THIS REDISPLAYS txtSPDFR AND lblKF,lblDS,lblDP.
    '---- LOCK SPDFR.
    txtSPDFR.Locked = True
  Else
    '---- Turn tortuosity(time) correlation OFF!
    Component(0).Use_Tortuosity_Correlation = False
    Component(0).Constant_Tortuosity = False
    'frmprint!chkSelect(4).Enabled = False
    '---- UNLOCK SPDFR.
    txtSPDFR.Locked = False
  End If
  'lblDS = Format$(Ds(0), "0.00E+00")
  Call Update_Ds_and_Dp_Editability
  Call Update_Tortuosity_Display
  'THROW DIRTY FLAG.
  Call frmKinetic_DirtyStatus_Throw
  'REFRESH WINDOW.
  Call frmKinetic_Refresh
End Sub


Private Sub cmdCancelOK_Click(Index As Integer)
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
      'SAVE USER/CORRELATION SELECTION OPTIONBOXES.
      Component(0).Corr(1) = optKF(1).Value
      Component(0).Corr(2) = optDS(1).Value
      Component(0).Corr(3) = optDP(1).Value
      'SAVE USER INPUT FOR KF/DS/DP/SPDFR.
      Component(0).KP_User_Input(1) = CDbl(txtKF)
      Component(0).KP_User_Input(2) = CDbl(txtDS)
      Component(0).KP_User_Input(3) = CDbl(txtDP)
      Component(0).SPDFR = CDbl(txtSPDFR)
      'SAVE CURRENT VALUES FOR KF/DS/DP.
      If optKF(0).Value = True Then
        Component(0).kf = CDbl(txtKF)
      Else
        Component(0).kf = kf(0)
      End If
      If optDS(0).Value = True Then
        Component(0).Ds = CDbl(txtDS)
      Else
        Component(0).Ds = Ds(0)
      End If
      If optDP(0).Value = True Then
        Component(0).Dp = CDbl(txtDP)
      Else
        Component(0).Dp = Dp(0)
      End If
      'SAVE CURRENT VALUE FOR TORTUOSITY AND CORRELATION SETTINGS.
      If (chkTortuosity_Corr.Value) Then
        Component(0).Use_Tortuosity_Correlation = True
        Component(0).Constant_Tortuosity = False
        Component(0).Tortuosity = 1
      Else
        Component(0).Use_Tortuosity_Correlation = False
        Component(0).Constant_Tortuosity = False
        Component(0).Tortuosity = txtTort   'iffy conversion here!
      End If
      'EXIT OUT OF HERE.
      USER_HIT_CANCEL = False
      USER_HIT_OK = True
      Unload Me
      Exit Sub
  End Select
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
  Me.Height = 6450
  Me.Width = 8535
  Call CenterOnForm(Me, frmCompoProp)
  Me.Caption = "Kinetic Parameters for " & Trim$(Component(0).Name)
  lblUnit(1) = "cm" & Chr$(178) & "/s"
  lblUnit(2) = "cm" & Chr$(178) & "/s"
  'TORTUOSITY CORRELATION DISPLAY.
  lblTortCorrelation.Visible = False
  If (Component(0).Use_Tortuosity_Correlation) Then
    chkTortuosity_Corr.Value = True
  Else
    chkTortuosity_Corr.Value = False
  End If
  'DISPLAY USER/CORRELATION SELECTION OPTIONBOXES.
  optKF(0).Value = Not (Component(0).Corr(1))
  optKF(1).Value = Component(0).Corr(1)
  optDS(0).Value = Not (Component(0).Corr(2))
  optDS(1).Value = Component(0).Corr(2)
  optDP(0).Value = Not (Component(0).Corr(3))
  optDP(1).Value = Component(0).Corr(3)
  For i = 0 To 1
    optKF(i).Enabled = True
    optDS(i).Enabled = True
    optDP(i).Enabled = True
  Next i
  'SAVE OLD COMPONENT FOR CANCEL ROLLBACK.
  SaveOldComponent = Component(0)
  'POPULATE UNIT CONTROLS.
  Call frmKinetic_PopulateUnits
  'DATA UNCHANGED AS YET.
  Call frmKinetic_DirtyStatus_Clear
  Call frmKinetic_GenericStatus_Set("")
  'REFRESH DISPLAY.
  Call frmKinetic_Refresh
  'DEMO SETTINGS.
  Call LOCAL___Reset_DemoVersionDisablings
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Call unitsys_unregister_all_on_form(Me)
End Sub




Private Sub UCtl_GotFocus(Ctl As Control)
Dim StatusMessagePanel As String
  Call unitsys_control_txtx_gotfocus(Ctl)
  If (Trim$(UCase$(Ctl.Name)) = Trim$(UCase$("txtKF"))) Then
    StatusMessagePanel = "Type in the user-input film diffusion coefficient"
  ElseIf (Trim$(UCase$(Ctl.Name)) = Trim$(UCase$("txtDS"))) Then
    StatusMessagePanel = "Type in the user-input surface diffusion coefficient"
  ElseIf (Trim$(UCase$(Ctl.Name)) = Trim$(UCase$("txtDP"))) Then
    StatusMessagePanel = "Type in the user-input pore diffusion coefficient"
  ElseIf (Trim$(UCase$(Ctl.Name)) = Trim$(UCase$("txtSPDFR"))) Then
    StatusMessagePanel = "Type in the user-input surface-to-pore diffusion flux ratio"
  ElseIf (Trim$(UCase$(Ctl.Name)) = Trim$(UCase$("txtTort"))) Then
    StatusMessagePanel = "Type in the user-input tortuosity"
  Else
    'NOT RECOGNIZED -- DO NOTHING.
  End If
  Call frmKinetic_GenericStatus_Set(StatusMessagePanel)
End Sub
Sub UCtl_LostFocus(Ctl As Control)
Dim NewValue_Okay As Integer
Dim NewValue As Double
Dim Val_Low As Double
Dim Val_High As Double
Dim Raise_Dirty_Flag As Boolean
  'NOTE: LOW AND HIGH VALUES IN BASE UNITS.
  If (Trim$(UCase$(Ctl.Name)) = Trim$(UCase$("txtKF"))) Then
    Val_Low = 1E-40: Val_High = 1E+40
  ElseIf (Trim$(UCase$(Ctl.Name)) = Trim$(UCase$("txtDS"))) Then
    Val_Low = 1E-40: Val_High = 1E+40
  ElseIf (Trim$(UCase$(Ctl.Name)) = Trim$(UCase$("txtDP"))) Then
    Val_Low = 1E-40: Val_High = 1E+40
  ElseIf (Trim$(UCase$(Ctl.Name)) = Trim$(UCase$("txtSPDFR"))) Then
    Val_Low = 1E-40: Val_High = 1E+40
  ElseIf (Trim$(UCase$(Ctl.Name)) = Trim$(UCase$("txtTort"))) Then
    Val_Low = 1E-40: Val_High = 1E+40
  Else
    'NOT RECOGNIZED -- DO NOTHING.
  End If
  NewValue_Okay = False
  If (unitsys_control_txtx_lostfocus_validate(Ctl, Val_Low, Val_High, NewValue, Raise_Dirty_Flag)) Then
    NewValue_Okay = True
  End If
  Call unitsys_control_txtx_lostfocus(Ctl, NewValue)
  Call frmKinetic_GenericStatus_Set("")
  If (NewValue_Okay) Then
    If (Raise_Dirty_Flag) Then
      'STORE TO MEMORY.
      If (Trim$(UCase$(Ctl.Name)) = Trim$(UCase$("txtKF"))) Then
        Component(0).KP_User_Input(1) = NewValue
      ElseIf (Trim$(UCase$(Ctl.Name)) = Trim$(UCase$("txtDS"))) Then
        Component(0).KP_User_Input(2) = NewValue
      ElseIf (Trim$(UCase$(Ctl.Name)) = Trim$(UCase$("txtDP"))) Then
        Component(0).KP_User_Input(3) = NewValue
      ElseIf (Trim$(UCase$(Ctl.Name)) = Trim$(UCase$("txtSPDFR"))) Then
        Component(0).SPDFR = NewValue
      ElseIf (Trim$(UCase$(Ctl.Name)) = Trim$(UCase$("txtTort"))) Then
        Component(0).Tortuosity = NewValue
      Else
        'NOT RECOGNIZED -- DO NOTHING.
      End If
      'RAISE DIRTY FLAG IF NECESSARY.
      If (Raise_Dirty_Flag) Then
        'THROW DIRTY FLAG.
        Call frmKinetic_DirtyStatus_Throw
      End If
      'REFRESH WINDOW.
      Call frmKinetic_Refresh
    End If
  End If
End Sub


Private Sub optDP_Click(Index As Integer)
  'THROW DIRTY FLAG.
  Call frmKinetic_DirtyStatus_Throw
  'REFRESH WINDOW.
  Call frmKinetic_Refresh
End Sub
Private Sub optDS_Click(Index As Integer)
  'THROW DIRTY FLAG.
  Call frmKinetic_DirtyStatus_Throw
  'REFRESH WINDOW.
  Call frmKinetic_Refresh
End Sub
Private Sub optKF_Click(Index As Integer)
  'THROW DIRTY FLAG.
  Call frmKinetic_DirtyStatus_Throw
  'REFRESH WINDOW.
  Call frmKinetic_Refresh
End Sub


Private Sub txtDP_GotFocus()
  Dim Ctl As Control: Set Ctl = txtDP: Call UCtl_GotFocus(Ctl)
End Sub
Private Sub txtDP_KeyPress(KeyAscii As Integer)
  KeyAscii = Global_NumericKeyPress(KeyAscii)
End Sub
Private Sub txtDP_LostFocus()
  Dim Ctl As Control: Set Ctl = txtDP: Call UCtl_LostFocus(Ctl)
End Sub

Private Sub txtDS_GotFocus()
  Dim Ctl As Control: Set Ctl = txtDS: Call UCtl_GotFocus(Ctl)
End Sub
Private Sub txtDS_KeyPress(KeyAscii As Integer)
  KeyAscii = Global_NumericKeyPress(KeyAscii)
End Sub
Private Sub txtDS_LostFocus()
  Dim Ctl As Control: Set Ctl = txtDS: Call UCtl_LostFocus(Ctl)
End Sub

Private Sub txtKF_GotFocus()
  Dim Ctl As Control: Set Ctl = txtKF: Call UCtl_GotFocus(Ctl)
End Sub
Private Sub txtKF_KeyPress(KeyAscii As Integer)
  KeyAscii = Global_NumericKeyPress(KeyAscii)
End Sub
Private Sub txtKF_LostFocus()
  Dim Ctl As Control: Set Ctl = txtKF: Call UCtl_LostFocus(Ctl)
End Sub

Private Sub txtSPDFR_GotFocus()
  Dim Ctl As Control: Set Ctl = txtSPDFR: Call UCtl_GotFocus(Ctl)
End Sub
Private Sub txtSPDFR_KeyPress(KeyAscii As Integer)
  KeyAscii = Global_NumericKeyPress(KeyAscii)
End Sub
Private Sub txtSPDFR_LostFocus()
  Dim Ctl As Control: Set Ctl = txtSPDFR: Call UCtl_LostFocus(Ctl)
End Sub

Private Sub txtTort_GotFocus()
  Dim Ctl As Control: Set Ctl = txtTort: Call UCtl_GotFocus(Ctl)
End Sub
Private Sub txtTort_KeyPress(KeyAscii As Integer)
  KeyAscii = Global_NumericKeyPress(KeyAscii)
End Sub
Private Sub txtTort_LostFocus()
  Dim Ctl As Control: Set Ctl = txtTort: Call UCtl_LostFocus(Ctl)
End Sub



