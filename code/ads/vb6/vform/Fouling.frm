VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmFouling 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fouling of GAC"
   ClientHeight    =   6975
   ClientLeft      =   2505
   ClientTop       =   1725
   ClientWidth     =   6495
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   6495
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   6360
      ScaleHeight     =   795
      ScaleWidth      =   1275
      TabIndex        =   41
      Top             =   5760
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
      Left            =   2520
      TabIndex        =   40
      ToolTipText     =   "Click here to print current screen to selected printer"
      Top             =   6030
      Width           =   1455
   End
   Begin Threed.SSFrame fraWater 
      Height          =   1250
      Left            =   120
      TabIndex        =   0
      Top             =   50
      Width           =   6255
      _Version        =   65536
      _ExtentX        =   11033
      _ExtentY        =   2205
      _StockProps     =   14
      Caption         =   "Water Type:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ComboBox cboType 
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
         Left            =   930
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   360
         Width           =   4515
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   375
         Left            =   930
         TabIndex        =   4
         Top             =   750
         Width           =   4515
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "Edit &Water Type Correlations"
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
   Begin Threed.SSFrame fraCompo 
      Height          =   4665
      Left            =   120
      TabIndex        =   1
      Top             =   1300
      Width           =   6255
      _Version        =   65536
      _ExtentX        =   11033
      _ExtentY        =   8229
      _StockProps     =   14
      Caption         =   "Chemical Type:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ComboBox cboCorrel 
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
         Index           =   9
         Left            =   3750
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   3735
         Width           =   2295
      End
      Begin VB.ComboBox cboCorrel 
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
         Index           =   8
         Left            =   3750
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   3375
         Width           =   2295
      End
      Begin VB.ComboBox cboCorrel 
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
         Index           =   7
         Left            =   3750
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   3015
         Width           =   2295
      End
      Begin VB.ComboBox cboCorrel 
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
         Index           =   6
         Left            =   3750
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2655
         Width           =   2295
      End
      Begin VB.ComboBox cboCorrel 
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
         Index           =   5
         Left            =   3750
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2295
         Width           =   2295
      End
      Begin VB.ComboBox cboCorrel 
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
         Index           =   4
         Left            =   3750
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1935
         Width           =   2295
      End
      Begin VB.ComboBox cboCorrel 
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
         Index           =   3
         Left            =   3750
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1575
         Width           =   2295
      End
      Begin VB.ComboBox cboCorrel 
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
         Index           =   2
         Left            =   3750
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1215
         Width           =   2295
      End
      Begin VB.ComboBox cboCorrel 
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
         Index           =   1
         Left            =   3750
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   855
         Width           =   2295
      End
      Begin VB.ComboBox cboCorrel 
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
         Index           =   0
         Left            =   3750
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   495
         Width           =   2295
      End
      Begin Threed.SSCommand cmdEditCompo 
         Height          =   375
         Left            =   930
         TabIndex        =   6
         Top             =   4140
         Width           =   4515
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "&Edit Chemical Type Correlations"
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
      Begin Threed.SSCheck chkUse 
         Height          =   255
         Index           =   0
         Left            =   330
         TabIndex        =   17
         Top             =   540
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
      Begin Threed.SSCheck chkUse 
         Height          =   255
         Index           =   1
         Left            =   330
         TabIndex        =   18
         Top             =   900
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
      Begin Threed.SSCheck chkUse 
         Height          =   255
         Index           =   2
         Left            =   330
         TabIndex        =   19
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
      End
      Begin Threed.SSCheck chkUse 
         Height          =   255
         Index           =   3
         Left            =   330
         TabIndex        =   20
         Top             =   1620
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
      Begin Threed.SSCheck chkUse 
         Height          =   255
         Index           =   4
         Left            =   330
         TabIndex        =   21
         Top             =   1980
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
      Begin Threed.SSCheck chkUse 
         Height          =   255
         Index           =   5
         Left            =   330
         TabIndex        =   22
         Top             =   2340
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
      Begin Threed.SSCheck chkUse 
         Height          =   255
         Index           =   6
         Left            =   330
         TabIndex        =   23
         Top             =   2700
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
      Begin Threed.SSCheck chkUse 
         Height          =   255
         Index           =   7
         Left            =   330
         TabIndex        =   24
         Top             =   3060
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
      Begin Threed.SSCheck chkUse 
         Height          =   255
         Index           =   8
         Left            =   330
         TabIndex        =   25
         Top             =   3420
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
      Begin Threed.SSCheck chkUse 
         Height          =   255
         Index           =   9
         Left            =   330
         TabIndex        =   26
         Top             =   3780
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
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Apply"
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
         TabIndex        =   39
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Type of correlation used"
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
         Left            =   3750
         TabIndex        =   38
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
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
         Left            =   990
         TabIndex        =   37
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Left            =   690
         TabIndex        =   36
         Top             =   3780
         Width           =   2895
      End
      Begin VB.Label lblName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Left            =   690
         TabIndex        =   35
         Top             =   3420
         Width           =   2895
      End
      Begin VB.Label lblName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Left            =   690
         TabIndex        =   34
         Top             =   3060
         Width           =   2895
      End
      Begin VB.Label lblName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Left            =   690
         TabIndex        =   33
         Top             =   2700
         Width           =   2895
      End
      Begin VB.Label lblName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Left            =   690
         TabIndex        =   32
         Top             =   2340
         Width           =   2895
      End
      Begin VB.Label lblName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Left            =   690
         TabIndex        =   31
         Top             =   1980
         Width           =   2895
      End
      Begin VB.Label lblName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Left            =   690
         TabIndex        =   30
         Top             =   1620
         Width           =   2895
      End
      Begin VB.Label lblName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Left            =   690
         TabIndex        =   29
         Top             =   1260
         Width           =   2895
      End
      Begin VB.Label lblName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Left            =   690
         TabIndex        =   28
         Top             =   900
         Width           =   2895
      End
      Begin VB.Label lblName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Left            =   690
         TabIndex        =   27
         Top             =   540
         Width           =   2895
      End
   End
   Begin Threed.SSCommand cmdCancelOK 
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   6400
      Width           =   2295
      _Version        =   65536
      _ExtentX        =   4048
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
      Left            =   4080
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   6400
      Width           =   2295
      _Version        =   65536
      _ExtentX        =   4048
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
End
Attribute VB_Name = "frmFouling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Dim Raise_Dirty_Flag As Boolean






Const frmFouling_decl_end = True


Sub frmFouling_Go(OUT_Raise_Dirty_Flag As Boolean)
  Raise_Dirty_Flag = False
  frmFouling.Show 1
  OUT_Raise_Dirty_Flag = Raise_Dirty_Flag
End Sub


Sub LOCAL___Reset_DemoVersionDisablings()
  If (IsThisADemo() = True) Then
    cmdCancelOK(1).Enabled = False
  End If
End Sub


Sub Populate_cboType()
Dim i As Integer
  i = False
  Call Load_Correlations_Water(i)
  If i Then Number_Water_Correlations = 0
  cboType.Clear
  For i = 1 To Number_Water_Correlations
    cboType.AddItem Trim$(Correlations_For_Water(i).Name)
  Next i
  'If (DemoMode) Then
  '  cboType.ListIndex = 0
  'Else
    If Number_Water_Correlations > 0 Then cboType.ListIndex = Set_Number_Correlation_Water() - 1
  'End If
End Sub
Sub Populate_cboCorrel()
Dim i As Integer
Dim J As Integer
  i = False
  Call Load_Correlation_Compounds(i)
  If i Then Number_Correlations_Compounds = 0
  For i = 0 To Number_Component - 1
    cboCorrel(i).Visible = True
    cboCorrel(i).Clear
    For J = 1 To Number_Correlations_Compounds
      cboCorrel(i).AddItem Trim$(Correlations_For_Classes(J).Name)
    Next J
    'If (DemoMode) Then
    '  ' Set it to Halogenated Alkenes
    '  If (0 = StrComp(lblName(i).Caption, "trichloroethylene", 1)) Then cboCorrel(i).ListIndex = 1
    '  ' Set it to Aromatics
    '  If (0 = StrComp(lblName(i).Caption, "benzene", 1)) Then cboCorrel(i).ListIndex = 3
    '  ' Set it to Aromatics
    '  If (0 = StrComp(lblName(i).Caption, "1,2-dichlorobenzene", 1)) Then cboCorrel(i).ListIndex = 3
    'Else
      cboCorrel(i).ListIndex = (Set_Number_Correlation(i + 1) - 1)
    'End If
    chkUse(i).Visible = True
    chkUse(i) = Component(i + 1).K_Reduction
    lblName(i).Visible = True
    lblName(i) = Trim$(Component(i + 1).Name)
  Next i
End Sub


Private Sub cboCorrel_Click(Index As Integer)
'If (DemoMode) Then
'    ' Set it to Halogenated Alkenes
'    If (0 = StrComp(Trim$(lblName(index).Caption), "trichloroethylene", 1)) Then cboCorrel(index).ListIndex = 1
'    ' Set it to Aromatics
'    If (0 = StrComp(Trim$(lblName(index).Caption), "benzene", 1)) Then cboCorrel(index).ListIndex = 3
'    ' Set it to Aromatics
'    If (0 = StrComp(Trim$(lblName(index).Caption), "1,2-dichlorobenzene", 1)) Then cboCorrel(index).ListIndex = 3
'End If
End Sub


Private Sub cboType_Click()
Dim msg$
Static old_index%
'  If (DemoMode) Then
'    If (0 = StrComp(Trim$(cboType.Text), "Organic Free Water", 1)) Then
'      old_index% = 0
'      Exit Sub
'    End If
'    If (0 = StrComp(Trim$(cboType.Text), "Groundwater from the city of Karlsruhe, Germany", 1)) Then
'      old_index% = 3
'      Exit Sub
'    End If
'    msg$ = "            " + cboType.Text + NL + NL
'    msg$ = msg$ + "Is not a valid Water Type in the Demonstration version."
'    MsgBox msg$
'    cboType.ListIndex = old_index%
'  End If
End Sub


Private Sub chkUse_Click(Index As Integer, Value As Integer)
Dim Is_Invalid As Boolean
  'DE-APPLY CORRELATION IF USER HAS NOT PROPERLY
  'SELECTED A CORRELATION.
  If (chkUse(Index) = True) Then
    Is_Invalid = False
    If (cboCorrel(Index).ListIndex < 0) Then
      Is_Invalid = True
    Else
      If (cboCorrel(Index).List(cboCorrel(Index).ListIndex) = "") Then
        Is_Invalid = True
      End If
    End If
    If (Is_Invalid) Then
      Call Show_Error("You must select a correlation " & _
          "type before you can apply fouling for this chemical.")
      chkUse(Index) = False
      Exit Sub
    End If
  End If
End Sub


Private Sub cmdCancelOK_Click(Index As Integer)
Dim i As Integer
Dim msg$
Dim IsInvalid As Boolean
  Select Case Index
    Case 0:   'CANCEL.
      Raise_Dirty_Flag = False
      Unload Me
    Case 1:   'OK.
      IsInvalid = True
      If (cboType.ListIndex >= 0) Then
        If (cboType.ListCount >= 1) Then
          If (Trim$(cboType.List(cboType.ListIndex)) <> "") Then
            IsInvalid = False
          End If
        End If
      End If
      If (IsInvalid) Then
        Call Show_Error("You must first select a water correlation type.")
        Exit Sub
      End If
'      If (DemoMode) Then
'        If (0 = StrComp(Trim$(cboType.Text), "Organic Free Water")) Then GoTo DEMO_00_CONTINUE
'        If (0 = StrComp(Trim$(cboType.Text), "Groundwater from the city of Karlsruhe, Germany", 1)) Then GoTo DEMO_00_CONTINUE
'        msg$ = "In Demonstration version you can only use two types of water:" + NL + NL
'        msg$ = msg$ + Chr$(9) + "- Organic Free Water" + NL
'        msg$ = msg$ + Chr$(9) + "- Groundwater from the city of Karlsruhe, Germany" + NL
'        MsgBox msg$
'        Exit Sub
'      End If
'DEMO_00_CONTINUE:
      For i = 1 To Number_Component
        If cboCorrel(i - 1).ListIndex > -1 Then
          Component(i).Correlation.Name = Trim$(cboCorrel(i - 1).List(cboCorrel(i - 1).ListIndex))
          Component(i).K_Reduction = chkUse(i - 1)
          Component(i).Correlation.Coeff(1) = Correlations_For_Classes(cboCorrel(i - 1).ListIndex + 1).Coeff(1)
          Component(i).Correlation.Coeff(2) = Correlations_For_Classes(cboCorrel(i - 1).ListIndex + 1).Coeff(2)
        Else
          Component(i).K_Reduction = False
        End If
      Next i
      If cboType.ListIndex = -1 Then cboType.ListIndex = 0
      Bed.Water_Correlation.Name = Correlations_For_Water(cboType.ListIndex + 1).Name
      For i = 1 To 4
        Bed.Water_Correlation.Coeff(i) = Correlations_For_Water(cboType.ListIndex + 1).Coeff(i)
      Next i
      '
      ' STORE SIGNAL TO RAISE DIRTY FLAG AND THEN EXIT.
      Raise_Dirty_Flag = True
      Unload Me
  End Select
End Sub


Private Sub cmdEdit_Click()
  Call frmFoulingWaterDatabase.frmFoulingWaterDatabase_Edit
  Call Populate_cboType
End Sub
Private Sub cmdEditCompo_Click()
  Call frmFoulingCompoundDatabase.frmFoulingCompoundDatabase_Edit
  Call Populate_cboCorrel
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
Dim J As Integer
  'If (DemoMode) Then
  '  cmdEdit.Enabled = False
  '  cmdEditCompo.Enabled = False
  'End If
  'Me.HelpContextID = Hlp_Fouling_of
  Call Populate_cboType
  
  Call Populate_cboCorrel
  
  For i = Number_Component To Number_Compo_Max - 1
    chkUse(i).Visible = False
    lblName(i).Visible = False
    cboCorrel(i).Visible = False
  Next i
  
  cmdEditCompo.Top = lblName(Number_Component - 1).Top + lblName(Number_Component - 1).Height + Screen.TwipsPerPixelY * 10
  fraCompo.Height = cmdEditCompo.Top + cmdEditCompo.Height + Screen.TwipsPerPixelY * 10
  cmdCancelOK(1).Top = fraCompo.Top + fraCompo.Height + Screen.TwipsPerPixelY * 10
  cmdCancelOK(0).Top = fraCompo.Top + fraCompo.Height + Screen.TwipsPerPixelY * 10
  Height = cmdCancelOK(1).Top + cmdCancelOK(1).Height + Screen.TwipsPerPixelY * 35
  cmdEditCompo.Left = (fraCompo.Width - cmdEditCompo.Width) / 2
  cmdEdit.Left = (fraWater.Width - cmdEdit.Width) / 2
  cboType.Left = (fraWater.Width - cboType.Width) / 2
  Me.Top = Screen.Height / 2 - Height / 2
  Me.Left = Screen.Width / 2 - Width / 2
  Call CenterOnForm(Me, frmMain)
  '
  ' DEMO SETTINGS.
  '
  Call LOCAL___Reset_DemoVersionDisablings
End Sub


Private Sub Load_Correlation_Compounds(flag As Integer)
Dim f As Integer, N As Integer, i As Integer
  On Error GoTo Error_In_Reading_Corr
  f = FreeFile
  Open Database_Path & "\corr_com.txt" For Input As f
  Input #f, N
  If N > Max_Number_Correlation_Compo Then
    flag = True
    Close (f)
    Call Show_Error("Too many correlations in the file.")
    Exit Sub
  End If
  For i = 1 To N
  Input #f, Correlations_For_Classes(i).Name, Correlations_For_Classes(i).Coeff(1), Correlations_For_Classes(i).Coeff(2)
  Next i
  Close (f)
  Number_Correlations_Compounds = N
  flag = False
  Exit Sub
Error_In_Reading_Corr:
  Call Show_Error("Error while reading the file containing correlations.")
  flag = True
  Resume Exit_Corr_Compound
Exit_Corr_Compound:
End Sub
Private Sub Load_Correlations_Water(flag As Integer)
Dim f As Integer, N As Integer, i As Integer
  On Error GoTo Error_In_Reading_WCorr
  f = FreeFile
  Open Database_Path & "\water_co.txt" For Input As f
  Input #f, N
  If N > Max_Number_Water_Correlations Then
    flag = True
    Close (f)
    Call Show_Error("Too many correlations in the file.")
    Exit Sub
  End If
  For i = 1 To N
    Input #f, Correlations_For_Water(i).Name, Correlations_For_Water(i).Coeff(1), Correlations_For_Water(i).Coeff(2), Correlations_For_Water(i).Coeff(3), Correlations_For_Water(i).Coeff(4)
  Next i
  Close (f)
  Number_Water_Correlations = N
  flag = False
  Exit Sub
Error_In_Reading_WCorr:
  Call Show_Error("Error while reading the file containing correlations.")
  flag = True
  Close (f)
  Resume Exit_Corr_Water
Exit_Corr_Water:
End Sub
Private Function Set_Number_Correlation(i As Integer) As Integer
Dim ST As String, J As Integer
  ST = Component(i).Correlation.Name
  For J = 1 To Number_Correlations_Compounds
    If Trim$(ST) = Trim$(Correlations_For_Classes(J).Name) Then
      Set_Number_Correlation = J
      Exit Function
    Else
      Set_Number_Correlation = 0
    End If
  Next J
End Function
Private Function Set_Number_Correlation_Water() As Integer
Dim ST As String, J As Integer
  ST = Bed.Water_Correlation.Name
  For J = 1 To Number_Water_Correlations
    If Trim$(ST) = Trim$(Correlations_For_Water(J).Name) Then
      Set_Number_Correlation_Water = J
      Exit Function
    Else
      Set_Number_Correlation_Water = 0
    End If
  Next J
End Function


