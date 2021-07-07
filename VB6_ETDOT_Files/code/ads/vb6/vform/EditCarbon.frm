VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmEditCarbon 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Adsorbent Database"
   ClientHeight    =   6600
   ClientLeft      =   420
   ClientTop       =   1245
   ClientWidth     =   9480
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   9480
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   9360
      ScaleHeight     =   795
      ScaleWidth      =   1275
      TabIndex        =   33
      Top             =   4200
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
      Left            =   7800
      TabIndex        =   32
      ToolTipText     =   "Click here to print current screen to selected printer"
      Top             =   5160
      Width           =   1455
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   2565
      Left            =   180
      TabIndex        =   0
      Top             =   240
      Width           =   3735
      _Version        =   65536
      _ExtentX        =   6588
      _ExtentY        =   4524
      _StockProps     =   14
      Caption         =   "Select a Manufacturer:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ListBox lstManu 
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
         Height          =   1980
         Left            =   120
         TabIndex        =   3
         Top             =   390
         Width           =   3495
      End
      Begin VB.Label lblEmpty_lstManu 
         Alignment       =   2  'Center
         Caption         =   "No Manufacturers Available"
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
         TabIndex        =   31
         Top             =   240
         Visible         =   0   'False
         Width           =   3495
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   4305
      Left            =   4140
      TabIndex        =   1
      Top             =   240
      Width           =   5145
      _Version        =   65536
      _ExtentX        =   9075
      _ExtentY        =   7594
      _StockProps     =   14
      Caption         =   "Adsorbent Properties:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Apparent Density"
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
         Left            =   450
         TabIndex        =   29
         Top             =   540
         Width           =   1755
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Particle Radius"
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
         Left            =   450
         TabIndex        =   28
         Top             =   900
         Width           =   1755
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Particle Porosity"
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
         Left            =   450
         TabIndex        =   27
         Top             =   1260
         Width           =   1755
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblData(0)"
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
         Left            =   2370
         TabIndex        =   26
         Top             =   525
         Width           =   1095
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblData(1)"
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
         Left            =   2370
         TabIndex        =   25
         Top             =   885
         Width           =   1095
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblData(2)"
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
         Left            =   2370
         TabIndex        =   24
         Top             =   1245
         Width           =   1095
      End
      Begin VB.Label lblUnit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "g/cm3"
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
         Left            =   3570
         TabIndex        =   23
         Top             =   540
         Width           =   975
      End
      Begin VB.Label lblUnit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "cm"
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
         Left            =   3570
         TabIndex        =   22
         Top             =   900
         Width           =   975
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
         Index           =   2
         Left            =   3570
         TabIndex        =   21
         Top             =   1260
         Width           =   975
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblData(3)"
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
         Left            =   2370
         TabIndex        =   20
         Top             =   1605
         Width           =   2175
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Adsorbent Type"
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
         Left            =   450
         TabIndex        =   19
         Top             =   1620
         Width           =   1755
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblData(4)"
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
         Left            =   2370
         TabIndex        =   18
         Top             =   1965
         Width           =   1095
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblData(5)"
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
         Left            =   2370
         TabIndex        =   17
         Top             =   2325
         Width           =   1095
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblData(6)"
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
         Left            =   2370
         TabIndex        =   16
         Top             =   2685
         Width           =   1095
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Polanyi W0"
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
         Left            =   450
         TabIndex        =   15
         Top             =   1980
         Width           =   1755
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Polanyi B"
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
         Left            =   450
         TabIndex        =   14
         Top             =   2340
         Width           =   1755
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Polanyi Exponent"
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
         Left            =   450
         TabIndex        =   13
         Top             =   2700
         Width           =   1755
      End
      Begin VB.Label lblUnit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "cm3/g"
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
         Left            =   3570
         TabIndex        =   12
         Top             =   1980
         Width           =   975
      End
      Begin VB.Label lblUnit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "*"
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
         Left            =   3570
         TabIndex        =   11
         Top             =   2340
         Width           =   975
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
         Index           =   6
         Left            =   3570
         TabIndex        =   10
         Top             =   2700
         Width           =   975
      End
      Begin VB.Label lblUnitB 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "* in (mol/cal) ^(Polanyi Exponent)"
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
         Left            =   1590
         TabIndex        =   9
         Top             =   3180
         Width           =   2925
      End
   End
   Begin Threed.SSFrame SSFrame3 
      Height          =   3285
      Left            =   180
      TabIndex        =   2
      Top             =   3030
      Width           =   3735
      _Version        =   65536
      _ExtentX        =   6588
      _ExtentY        =   5794
      _StockProps     =   14
      Caption         =   "Select an Adsorbent:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ListBox lstName 
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
         TabIndex        =   6
         Top             =   720
         Width           =   3495
      End
      Begin Threed.SSOption optPhase 
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "&Liquid Phase"
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption optPhase 
         Height          =   375
         Index           =   1
         Left            =   2040
         TabIndex        =   5
         Top             =   240
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "&Gas Phase"
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblEmpty_lstName 
         Alignment       =   2  'Center
         Caption         =   "No Adsorbents Available"
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
         TabIndex        =   30
         Top             =   570
         Visible         =   0   'False
         Width           =   3495
      End
   End
   Begin Threed.SSCommand cmdOK 
      Height          =   615
      Left            =   4140
      TabIndex        =   7
      Top             =   5730
      Width           =   4035
      _Version        =   65536
      _ExtentX        =   7117
      _ExtentY        =   1085
      _StockProps     =   78
      Caption         =   "&Use These Adsorbent Specifications"
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
      Height          =   615
      Left            =   8160
      TabIndex        =   8
      Top             =   5730
      Width           =   1125
      _Version        =   65536
      _ExtentX        =   1984
      _ExtentY        =   1085
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
   Begin VB.Menu mnuManufacturer 
      Caption         =   "&Manufacturer"
      Begin VB.Menu mnuManufacturerItem 
         Caption         =   "&New"
         Index           =   1
      End
      Begin VB.Menu mnuManufacturerItem 
         Caption         =   "&Edit Current"
         Index           =   2
      End
      Begin VB.Menu mnuManufacturerItem 
         Caption         =   "&Delete Current"
         Index           =   3
      End
   End
   Begin VB.Menu mnuAdsorbent 
      Caption         =   "&Adsorbent"
      Begin VB.Menu mnuAdsorbentItem 
         Caption         =   "&New"
         Index           =   1
      End
      Begin VB.Menu mnuAdsorbentItem 
         Caption         =   "&Edit Current"
         Index           =   2
      End
      Begin VB.Menu mnuAdsorbentItem 
         Caption         =   "&Delete Current"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmEditCarbon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim FORM_MODE As Integer
Const FORM_MODE_QUERY_DATABASE = 1
Const FORM_MODE_EDIT_DATABASE = 2

Dim USER_HIT_CANCEL As Boolean
Dim USER_HIT_USE_THESE As Boolean

Dim DB_Carbon As Database

Private LocalCarbon_Record As frmEditCarbonData_Record_Type




Const frmEditCarbon_declarations_end = True


Sub frmEditCarbon_QueryDatabase( _
    OUTPUT_User_Transferred_Data As Boolean)
  
  On Error GoTo err_frmEditCarbon_QueryDatabase
  'NOTE: THE FOLLOWING OpenDatabase() COMMAND MUST BE
  'SPECIFIED EXACTLY AS-IS, OR ELSE IT WILL FAIL.
  Set DB_Carbon = _
      Ws1.OpenDatabase(fn_DB_Carbon, True, False, _
      ";pwd=" & decrypt_string(Encrypted_User_Password))
  'Set DB_Carbon = ws1.OpenDatabase(fn_DB_Carbon)
  FORM_MODE = FORM_MODE_QUERY_DATABASE
  frmEditCarbon.Show 1
  If (USER_HIT_USE_THESE) Then
    OUTPUT_User_Transferred_Data = True
  Else
    OUTPUT_User_Transferred_Data = False
  End If
  DB_Carbon.Close
  Exit Sub
exit_err_frmEditCarbon_QueryDatabase:
  Exit Sub
err_frmEditCarbon_QueryDatabase:
  Call Show_Trapped_Error("frmEditCarbon_QueryDatabase")
  OUTPUT_User_Transferred_Data = False
  Resume exit_err_frmEditCarbon_QueryDatabase
End Sub
Sub frmEditCarbon_EditDatabase()
  On Error GoTo err_frmEditCarbon_EditDatabase
  'NOTE: THE FOLLOWING OpenDatabase() COMMAND MUST BE
  'SPECIFIED EXACTLY AS-IS, OR ELSE IT WILL FAIL.
  Set DB_Carbon = _
      Ws1.OpenDatabase(fn_DB_Carbon, True, False, _
      ";pwd=" & decrypt_string(Encrypted_User_Password))
  'Set DB_Carbon = ws1.OpenDatabase(fn_DB_Carbon)
  FORM_MODE = FORM_MODE_EDIT_DATABASE
  frmEditCarbon.Show 1
  DB_Carbon.Close
  Exit Sub
exit_err_frmEditCarbon_EditDatabase:
  Exit Sub
err_frmEditCarbon_EditDatabase:
  Call Show_Trapped_Error("frmEditCarbon_EditDatabase")
  Resume exit_err_frmEditCarbon_EditDatabase
End Sub
Sub LOCAL___Reset_DemoVersionDisablings()
  If (IsThisADemo() = True) Then
    cmdOK.Enabled = False
    mnuManufacturerItem(1).Enabled = False
    mnuManufacturerItem(2).Enabled = False
    mnuManufacturerItem(3).Enabled = False
    mnuAdsorbentItem(1).Enabled = False
    mnuAdsorbentItem(2).Enabled = False
    mnuAdsorbentItem(3).Enabled = False
  End If
End Sub


Sub populate_lstManu()
Dim Rs1 As Recordset
Dim Current_Criteria As String
Dim SAVE_CURRENT_POSITION As Long
Dim NEW_LISTINDEX As Integer
Dim This_ID As Long
Dim NumRecords As Long
  On Error GoTo err_populate_lstManu
  'SAVE CURRENT POSITION.
  If (lstManu.ListCount > 0) And (lstManu.ListIndex >= 0) Then
    SAVE_CURRENT_POSITION = lstManu.ItemData(lstManu.ListIndex)
  Else
    SAVE_CURRENT_POSITION = -1
  End If
  'SET UP SEARCH CRITERIA.
  Current_Criteria = "select * from [Manufacturers] " & _
      "order by [Name]"
  'START SEARCH.
  Set Rs1 = _
      DB_Carbon.OpenRecordset(Current_Criteria)
  On Error Resume Next
  Rs1.MoveFirst
  Rs1.MoveLast
  Rs1.MoveFirst
  NumRecords = Rs1.RecordCount
  On Error GoTo err_populate_lstManu
  'POPULATE LISTBOX.
  lstManu.Clear
  If (NumRecords = 0) Then
    'NO RECORDS AVAILABLE.
    lstManu.Visible = False
    lblEmpty_lstManu.Move lstManu.Left, lstManu.Top
    lblEmpty_lstManu.Visible = True
  Else
    'DISPLAY RECORDS.
    lstManu.Visible = True
    lblEmpty_lstManu.Visible = False
    NEW_LISTINDEX = 0
    Do Until Rs1.EOF
      lstManu.AddItem Database_Get_String(Rs1, "Name")
      This_ID = Database_Get_Long(Rs1, "Manufacturer ID")
      lstManu.ItemData(lstManu.NewIndex) = This_ID
      If (SAVE_CURRENT_POSITION <> -1) Then
        If (SAVE_CURRENT_POSITION = This_ID) Then
          NEW_LISTINDEX = lstManu.NewIndex
        End If
      End If
      Rs1.MoveNext
    Loop
    If (lstManu.ListCount > 0) Then
      lstManu.ListIndex = NEW_LISTINDEX
    End If
  End If
  'CLOSE DATABASE AND EXIT.
  Rs1.Close
  Exit Sub
exit_err_populate_lstManu:
  Exit Sub
err_populate_lstManu:
  Call Show_Trapped_Error("populate_lstManu")
  Resume exit_err_populate_lstManu
End Sub
Sub populate_lstName(THIS_ITEMDATA As Long)
Dim PHASE_CODE As Integer
Dim Rs1 As Recordset
Dim Current_Criteria As String
Dim SAVE_CURRENT_POSITION As Long
Dim This_ID As Long
Dim NEW_LISTINDEX As Long
Dim NumRecords As Long
  On Error GoTo err_populate_lstName
  'GET PHASE CODE.
  If (optPhase(0).Value) Then PHASE_CODE = 1
  If (optPhase(1).Value) Then PHASE_CODE = 2
  'SAVE CURRENT POSITION.
  If (lstName.ListCount > 0) And (lstName.ListIndex >= 0) Then
    SAVE_CURRENT_POSITION = lstName.ItemData(lstName.ListIndex)
  Else
    SAVE_CURRENT_POSITION = -1
  End If
  'SET UP SEARCH CRITERIA.
  Current_Criteria = "select * from [Carbon Data] " & _
      "where [Manufacturer ID]=" & _
      Trim$(Str$(THIS_ITEMDATA)) & " and " & _
      "[Phase ID]=" & Trim$(Str$(PHASE_CODE)) & " " & _
      "order by [Name]"
  'START SEARCH.
  Set Rs1 = _
      DB_Carbon.OpenRecordset(Current_Criteria)
  On Error Resume Next
  Rs1.MoveFirst
  Rs1.MoveLast
  Rs1.MoveFirst
  NumRecords = Rs1.RecordCount
  On Error GoTo err_populate_lstName
  'POPULATE LISTBOX.
  lstName.Clear
  If (NumRecords = 0) Then
    'NO RECORDS AVAILABLE.
    lstName.Visible = False
    lblEmpty_lstName.Move lstName.Left, lstName.Top
    lblEmpty_lstName.Visible = True
  Else
    'DISPLAY RECORDS.
    lstName.Visible = True
    lblEmpty_lstName.Visible = False
    NEW_LISTINDEX = 0
    Do Until Rs1.EOF
      lstName.AddItem Database_Get_String(Rs1, "Name")
      This_ID = Database_Get_Long(Rs1, "ID")
      lstName.ItemData(lstName.NewIndex) = This_ID
      If (SAVE_CURRENT_POSITION <> -1) Then
        If (SAVE_CURRENT_POSITION = This_ID) Then
          NEW_LISTINDEX = lstName.NewIndex
        End If
      End If
      Rs1.MoveNext
    Loop
    If (lstName.ListCount > 0) Then
      lstName.ListIndex = NEW_LISTINDEX
    End If
  End If
  'CLOSE DATABASE AND EXIT.
  Rs1.Close
  Exit Sub
exit_err_populate_lstName:
  Exit Sub
err_populate_lstName:
  Call Show_Trapped_Error("populate_lstName")
  Resume exit_err_populate_lstName
End Sub
Sub populate_lblData(THIS_ITEMDATA As Long)
Dim Rs1 As Recordset
Dim NumRecords As Long
Dim Current_Criteria As String
Dim TempDbl As Double
  On Error GoTo err_populate_lblData
  'SET UP SEARCH CRITERIA.
  Current_Criteria = "select * from [Carbon Data] " & _
      "where [ID]=" & _
      Trim$(Str$(THIS_ITEMDATA)) & " " & _
      "order by [Name]"
  'START SEARCH.
  Set Rs1 = _
      DB_Carbon.OpenRecordset(Current_Criteria)
  On Error Resume Next
  Rs1.MoveFirst
  Rs1.MoveLast
  Rs1.MoveFirst
  NumRecords = Rs1.RecordCount
  On Error GoTo err_populate_lblData
  'POPULATE LABEL CONTROLS.
  If (NumRecords = 0) Then
    'NO RECORD(S) AVAILABLE; WEIRD PROBLEM.
    'DO NOTHING.
  Else
    'DISPLAY (FIRST) RECORD (THERE SHOULD ONLY BE ONE).
    LocalCarbon_Record.Name = Database_Get_String(Rs1, "Name")
    LocalCarbon_Record.Manufacturer = ""
    On Error Resume Next
    LocalCarbon_Record.Manufacturer = lstManu.List(lstManu.ListIndex)
    On Error GoTo err_populate_lblData
    LocalCarbon_Record.AppDen = Database_Get_Double(Rs1, "Apparent Density")
        'NOTE: RADIUS IN DATABASE IS STORED IN millimeters;
        'DIVISION BY 10 CONVERTS THIS TO centimers.
    LocalCarbon_Record.ParticleRadius = _
        Database_Get_Double(Rs1, "Average Particle Radius") / 10#
    LocalCarbon_Record.ParticlePorosity = Database_Get_Double(Rs1, "Porosity")
    LocalCarbon_Record.AdsType = Database_Get_String(Rs1, "Type")
    Call AssignCaptionAndTag(lblData(0), LocalCarbon_Record.AppDen)
    Call AssignCaptionAndTag(lblData(1), LocalCarbon_Record.ParticleRadius)
    Call AssignCaptionAndTag(lblData(2), LocalCarbon_Record.ParticlePorosity)
    Call AssignCaptionAndTag(lblData(3), LocalCarbon_Record.AdsType)
    TempDbl = Database_Get_Double(Rs1, "W0")
    LocalCarbon_Record.W0 = TempDbl
    If (TempDbl = 0#) Then
      Call AssignCaptionAndTag(lblData(4), "Unavailable")
    Else
      Call AssignCaptionAndTag(lblData(4), TempDbl)
    End If
    TempDbl = Database_Get_Double(Rs1, "BB")
    LocalCarbon_Record.BB = TempDbl
    If (TempDbl = 0#) Then
      Call AssignCaptionAndTag(lblData(5), "Unavailable")
    Else
      Call AssignCaptionAndTag(lblData(5), TempDbl)
    End If
    TempDbl = Database_Get_Double(Rs1, "Polanyi Exponent")
    LocalCarbon_Record.PolanyiExponent = TempDbl
    If (TempDbl = 0#) Then
      Call AssignCaptionAndTag(lblData(6), "Unavailable")
    Else
      Call AssignCaptionAndTag(lblData(6), TempDbl)
    End If
  End If
  'CLOSE DATABASE AND EXIT.
  Rs1.Close
  Exit Sub
exit_err_populate_lblData:
  Exit Sub
err_populate_lblData:
  Call Show_Trapped_Error("populate_lblData")
  Resume exit_err_populate_lblData
End Sub


Private Sub cmdCancel_Click()
  'frmEditAdsorber_Cancelled = True
  USER_HIT_CANCEL = True
  USER_HIT_USE_THESE = False
  Unload Me
End Sub
Private Sub cmdOK_Click()
  If (lstManu.ListIndex < 0) Or (lstManu.ListCount = 0) Then
    Call Show_Error("You must first select a manufacturer.")
    Exit Sub
  End If
  If (lstName.ListIndex < 0) Or (lstName.ListCount = 0) Then
    Call Show_Error("You must first select an adsorbent.")
    Exit Sub
  End If
  If (LocalCarbon_Record.AppDen = 0#) Or _
      (LocalCarbon_Record.ParticlePorosity = 0#) Or _
      (LocalCarbon_Record.ParticleRadius = 0#) Then
    Call Show_Error("You must select an adsorbent with non-zero values " & _
        "for apparent density, particle porosity, and particle radius.")
    Exit Sub
  End If
  '
  ' TRANSFER DATA TO CURRENT CARBON RECORD.
  '
  Carbon.Name = Trim$(LocalCarbon_Record.Manufacturer)
  If (Carbon.Name <> "") Then
    Carbon.Name = Carbon.Name & " "
  End If
  Carbon.Name = Carbon.Name & Trim$(LocalCarbon_Record.Name)
  Carbon.Density = LocalCarbon_Record.AppDen
  Carbon.ParticleRadius = LocalCarbon_Record.ParticleRadius / 100#
  Carbon.Porosity = LocalCarbon_Record.ParticlePorosity
  Carbon.ShapeFactor = 1#
  Carbon.W0 = LocalCarbon_Record.W0
  Carbon.BB = LocalCarbon_Record.BB
  Carbon.PolanyiExponent = LocalCarbon_Record.PolanyiExponent
  '
  ' EXIT OUT OF HERE.
  '
  USER_HIT_CANCEL = False
  USER_HIT_USE_THESE = True
  Unload Me
End Sub


Private Sub Command4_Click()
    Set Picture1.Picture = CaptureActiveWindow()
    PrintPictureToFitPage Printer, Picture1.Picture
    Printer.EndDoc
    ' Set focus back to form.
    Me.SetFocus
End Sub

Private Sub Form_Load()
  'MISC INITS.
  Me.Height = 7290
  Me.Width = 9600
  Call CenterOnForm(Me, frmMain)
  lblUnit(0).Caption = "g/cm³"
  lblUnit(4).Caption = "cm³/g"
  If (FORM_MODE = FORM_MODE_QUERY_DATABASE) Then
    'QUERY DATABASE MODE.
    cmdOK.Visible = True
    cmdCancel.Visible = True
    cmdCancel.Caption = "&Cancel"
  Else
    'EDIT DATABASE MODE.
    cmdOK.Visible = False
    cmdCancel.Visible = True
    cmdCancel.Caption = "E&xit"
  End If
  If (Bed.Phase = 0) Then
    optPhase(0).Value = True
    optPhase(1).Value = False
  Else
    optPhase(0).Value = False
    optPhase(1).Value = True
  End If
  'RE-POPULATE MANUFACTURER LIST.
  Call populate_lstManu
  'DEMO SETTINGS.
  Call LOCAL___Reset_DemoVersionDisablings
End Sub


Private Sub lstManu_Click()
Dim THIS_ITEMDATA As Long
  If (lstManu.ListIndex < 0) Or (lstManu.ListCount <= 0) Then
    Exit Sub
  End If
  THIS_ITEMDATA = lstManu.ItemData(lstManu.ListIndex)
  Call populate_lstName(THIS_ITEMDATA)
End Sub
Private Sub lstManu_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If ((Button And 2) = 2) Then
    Me.PopupMenu mnuManufacturer
  End If
End Sub
Private Sub lstName_Click()
Dim THIS_ITEMDATA As Long
  If (lstName.ListIndex < 0) Or (lstName.ListCount <= 0) Then
    Exit Sub
  End If
  THIS_ITEMDATA = lstName.ItemData(lstName.ListIndex)
  Call populate_lblData(THIS_ITEMDATA)
End Sub
Private Sub lstName_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If ((Button And 2) = 2) Then
    Me.PopupMenu mnuAdsorbent
  End If
End Sub


Private Sub mnuAdsorbentItem_Click(Index As Integer)
Dim USER_HIT_CANCEL As Boolean
Dim USER_HIT_SAVE As Boolean
Dim USER_HIT_SAVEASNEW As Boolean
Dim THIS_MANU_ID As Long
Dim THIS_ADS_ID As Long
Dim Current_Criteria As String
Dim Rs1 As Recordset
Dim NumRecords As Long
Dim DEFAULT_PHASE_IS_LIQUID As Boolean
Const INDEX_NEW = 1
Const INDEX_EDIT = 2
Const INDEX_DELETE = 3
Dim NewName As String
Dim msg As String
Dim RetVal As Integer
Dim Select_Index As Integer
Dim i As Integer
  On Error GoTo err_mnuAdsorbentItem_Click
  If (lstManu.ListIndex < 0) Or (lstManu.ListCount = 0) Then
    Call Show_Error("You must first select a manufacturer.")
    Exit Sub
  End If
  THIS_MANU_ID = lstManu.ItemData(lstManu.ListIndex)
  If (Index = INDEX_EDIT) Or (Index = INDEX_DELETE) Then
    If (lstName.ListIndex < 0) Or (lstName.ListCount = 0) Then
      Call Show_Error("You must first select an adsorbent.")
      Exit Sub
    End If
    THIS_ADS_ID = lstName.ItemData(lstName.ListIndex)
    'SET UP SEARCH CRITERIA.
    Current_Criteria = "select * from [Carbon Data] " & _
        "where [ID]=" & _
        Trim$(Str$(THIS_ADS_ID)) & " " & _
        "order by [Name]"
    'START SEARCH.
    Set Rs1 = _
        DB_Carbon.OpenRecordset(Current_Criteria)
    On Error Resume Next
    Rs1.MoveFirst
    Rs1.MoveLast
    Rs1.MoveFirst
    NumRecords = Rs1.RecordCount
    On Error GoTo err_mnuAdsorbentItem_Click
    'POPULATE LABEL CONTROLS.
    If (NumRecords = 0) Then
      'NO RECORD(S) AVAILABLE; WEIRD PROBLEM.
      'EXIT SUBROUTINE.
      Exit Sub
    End If
      'Call AssignCaptionAndTag(lblData(0), Database_Get_Double(RS1, "Apparent Density"))
      'Call AssignCaptionAndTag(lblData(1), Database_Get_Double(RS1, "Average Particle Radius"))
      'Call AssignCaptionAndTag(lblData(2), Database_Get_Double(RS1, "Porosity"))
      'Call AssignCaptionAndTag(lblData(3), Database_Get_String(RS1, "Type"))
  End If
  Select Case Index
    Case INDEX_NEW:     'NEW ADSORBENT.
      'DETERMINE DEFAULT PHASE.
      If (optPhase(0).Value) Then
        DEFAULT_PHASE_IS_LIQUID = True
      Else
        DEFAULT_PHASE_IS_LIQUID = False
      End If
      'ALLOW USER TO ADD NEW RECORD.
      Call frmEditCarbonData.frmEditCarbonData_AddNew( _
          DEFAULT_PHASE_IS_LIQUID, _
          USER_HIT_CANCEL, _
          USER_HIT_SAVE)
      If (USER_HIT_CANCEL) Then Exit Sub
      'ADD THE NEW ADSORBENT RECORD.
      Current_Criteria = "select * from [Carbon Data]"
      Set Rs1 = _
          DB_Carbon.OpenRecordset(Current_Criteria)
      Rs1.AddNew
      'SET THE MANUFACTURER-ID AND PHASE FIELDS.
      Rs1("Manufacturer ID") = THIS_MANU_ID
      If (frmEditCarbonData_Record.PhaseIsLiquid) Then
        Rs1("Phase ID") = 1
      Else
        Rs1("Phase ID") = 2
      End If
      'THE FIELD [ID] IS AUTOMATICALLY UPDATED.
      Rs1("Name") = frmEditCarbonData_Record.Name
      Rs1("Type") = frmEditCarbonData_Record.AdsType
      Rs1("Apparent Density") = frmEditCarbonData_Record.AppDen
      ' NEXT LINE CONVERTS centimeters TO millimeters.
      Rs1("Average Particle Radius") = frmEditCarbonData_Record.ParticleRadius * 10#
      Rs1("Porosity") = frmEditCarbonData_Record.ParticlePorosity
      Rs1("W0") = frmEditCarbonData_Record.W0
      Rs1("BB") = frmEditCarbonData_Record.BB
      Rs1("Polanyi Exponent") = frmEditCarbonData_Record.PolanyiExponent
      THIS_ADS_ID = Database_Get_Long(Rs1, "ID")
      Rs1.Update
      'CLOSE THE DATABASE.
      Rs1.Close
      'UPDATE WINDOW.
      Call lstManu_Click
      'SELECT THE NEW ADSORBENT.
      Select_Index = 0
      For i = 0 To lstName.ListCount - 1
        If (lstName.ItemData(i) = THIS_ADS_ID) Then
          Select_Index = i
          Exit For
        End If
      Next i
      If (lstName.ListCount > 0) Then
        lstName.ListIndex = Select_Index
      End If
    Case INDEX_EDIT:     'EDIT ADSORBENT.
      'TRANSFER DATABASE RECORD FIELDS TO LOCAL MEMORY.
      If (Database_Get_Long(Rs1, "Phase ID") = 1) Then
        frmEditCarbonData_Record.PhaseIsLiquid = True
      Else
        frmEditCarbonData_Record.PhaseIsLiquid = False
      End If
      frmEditCarbonData_Record.AppDen = Database_Get_Double(Rs1, "Apparent Density")
      ' NEXT LINE CONVERTS millimeters TO centimeters.
      frmEditCarbonData_Record.ParticleRadius = _
          Database_Get_Double(Rs1, "Average Particle Radius") / 10#
      frmEditCarbonData_Record.ParticlePorosity = Database_Get_Double(Rs1, "Porosity")
      frmEditCarbonData_Record.W0 = Database_Get_Double(Rs1, "W0")
      frmEditCarbonData_Record.BB = Database_Get_Double(Rs1, "BB")
      frmEditCarbonData_Record.PolanyiExponent = Database_Get_Double(Rs1, "Polanyi Exponent")
      frmEditCarbonData_Record.Name = Database_Get_String(Rs1, "Name")
      frmEditCarbonData_Record.AdsType = Database_Get_String(Rs1, "Type")
      'ALLOW USER TO EDIT THIS RECORD.
      Call frmEditCarbonData.frmEditCarbonData_Edit( _
          USER_HIT_CANCEL, _
          USER_HIT_SAVE, _
          USER_HIT_SAVEASNEW)
      If (USER_HIT_CANCEL) Then Exit Sub
      'SAVE THE EDITED ADSORBENT RECORD.
      Current_Criteria = "select * from [Carbon Data] " & _
          "where [ID]=" & Trim$(Str$(THIS_ADS_ID))
      Set Rs1 = _
          DB_Carbon.OpenRecordset(Current_Criteria)
      'SET THE MANUFACTURER-ID AND PHASE FIELDS.
      If (USER_HIT_SAVE) Then
        'MODIFY EXISTING RECORD.
        Rs1.Edit
        'KEEP ORIGINAL [Manufacturer ID] FIELD INTACT.
        'KEEP ORIGINAL [ID] FIELD INTACT.
      End If
      If (USER_HIT_SAVEASNEW) Then
        'GENERATE NEW RECORD.
        Rs1.AddNew
        'SAVE [Manufacturer ID] FIELD.
        Rs1("Manufacturer ID") = THIS_MANU_ID
        'THE FIELD [ID] IS AUTOMATICALLY CREATED DURING THE .Update COMMAND.
      End If
      If (frmEditCarbonData_Record.PhaseIsLiquid) Then
        Rs1("Phase ID") = 1
      Else
        Rs1("Phase ID") = 2
      End If
      Rs1("Name") = frmEditCarbonData_Record.Name
      Rs1("Type") = frmEditCarbonData_Record.AdsType
      Rs1("Apparent Density") = frmEditCarbonData_Record.AppDen
      ' NEXT LINE CONVERTS centimeters TO millimeters.
      Rs1("Average Particle Radius") = frmEditCarbonData_Record.ParticleRadius * 10#
      Rs1("Porosity") = frmEditCarbonData_Record.ParticlePorosity
      Rs1("W0") = frmEditCarbonData_Record.W0
      Rs1("BB") = frmEditCarbonData_Record.BB
      Rs1("Polanyi Exponent") = frmEditCarbonData_Record.PolanyiExponent
      Rs1.Update
      'CLOSE THE DATABASE.
      Rs1.Close
      'UPDATE WINDOW.
      Call lstManu_Click
    Case INDEX_DELETE:     'DELETE ADSORBENT.
      NewName = Trim$(lstName.List(lstName.ListIndex))
      msg = "Do you really want to delete adsorbent '" & _
          NewName & "' ?"
      RetVal = MsgBox(msg, vbCritical + vbYesNo, AppName_For_Display_Long)
      If RetVal = vbNo Then Exit Sub
      'PERFORM DELETION.
      Current_Criteria = "select * from [Carbon Data] where " & _
          "[ID] = " & Trim$(Str$(THIS_ADS_ID))
      Set Rs1 = _
          DB_Carbon.OpenRecordset(Current_Criteria)
      Rs1.Delete
      'CLOSE THE DATABASE.
      Rs1.Close
      'UPDATE WINDOW.
      Call lstManu_Click
  End Select
  Exit Sub
exit_err_mnuAdsorbentItem_Click:
  Exit Sub
err_mnuAdsorbentItem_Click:
  Call Show_Trapped_Error("mnuAdsorbentItem_Click")
  Resume exit_err_mnuAdsorbentItem_Click
End Sub
Private Sub mnuManufacturerItem_Click(Index As Integer)
Dim THIS_MANU_ID As Long
Dim NewName As String
'Dim USER_HIT_CANCEL As Boolean
Dim Current_Criteria As String
Dim Rs1 As Recordset
Dim msg As String
Dim RetVal As Integer
Dim i As Integer
Dim Select_Index As Integer
Dim NumRecords As Integer
  On Error GoTo err_mnuManufacturerItem_Click
  If (Index = 2) Or (Index = 3) Then
    If (lstManu.ListIndex < 0) Or (lstManu.ListCount = 0) Then
      Call Show_Error("You must first select a manufacturer.")
      Exit Sub
    End If
    THIS_MANU_ID = lstManu.ItemData(lstManu.ListIndex)
  End If
  Select Case Index
    Case 1:     'new
      NewName = "New Manufacturer"
      Do While (1 = 1)
        NewName = frmNewName.frmNewName_GetName( _
            "Creating New Manufacturer", _
            "Each manufacturer record should have a unique name.", _
            NewName, _
            USER_HIT_CANCEL)
        If (USER_HIT_CANCEL) Then Exit Sub
        NewName = Trim$(NewName)
        If (NewName <> "") Then Exit Do
        Call Show_Error("Manufacturer name must be a non-blank string.")
      Loop
      'ADD THE NEW MANUFACTURER RECORD.
      Current_Criteria = "select * from [Manufacturers]"
      Set Rs1 = _
          DB_Carbon.OpenRecordset(Current_Criteria)
      Rs1.AddNew
      'THE FIELD [Manufacturer ID] IS AUTOMATICALLY UPDATED.
      Rs1("Name") = NewName
      THIS_MANU_ID = Database_Get_Long(Rs1, "Manufacturer ID")
      Rs1.Update
      Rs1.Close
      'REDISPLAY WINDOW.
      Call populate_lstManu
      'SELECT THE NEW MANUFACTURER.
      Select_Index = 0
      For i = 0 To lstManu.ListCount - 1
        If (lstManu.ItemData(i) = THIS_MANU_ID) Then
          Select_Index = i
          Exit For
        End If
      Next i
      If (lstManu.ListCount > 0) Then
        lstManu.ListIndex = Select_Index
      End If
    Case 2:     'edit current
      NewName = Trim$(lstManu.List(lstManu.ListIndex))
      Do While (1 = 1)
        NewName = frmNewName.frmNewName_GetName( _
            "Editing Existing Manufacturer Name", _
            "Each manufacturer record should have a unique name.", _
            NewName, _
            USER_HIT_CANCEL)
        If (USER_HIT_CANCEL) Then Exit Sub
        NewName = Trim$(NewName)
        If (NewName <> "") Then Exit Do
        Call Show_Error("Manufacturer name must be a non-blank string.")
      Loop
      'EDIT THE MANUFACTURER RECORD.
      Current_Criteria = "select * from [Manufacturers] where " & _
          "[Manufacturer ID] = " & Trim$(Str$(THIS_MANU_ID))
      Set Rs1 = _
          DB_Carbon.OpenRecordset(Current_Criteria)
      Rs1.Edit
      Rs1("Name") = NewName
      Rs1.Update
      Rs1.Close
      'REDISPLAY WINDOW.
      Call populate_lstManu
    Case 3:     'delete current
      NewName = Trim$(lstManu.List(lstManu.ListIndex))
      msg = "Do you really want to delete manufacturer '" & _
          NewName & "' and all of the corresponding adsorbent " & _
          "records from the database ?"
      RetVal = MsgBox(msg, vbCritical + vbYesNo, AppName_For_Display_Long)
      If RetVal = vbNo Then Exit Sub
      'PERFORM DELETION OF MANUFACTURER RECORD.
      Current_Criteria = "select * from [Manufacturers] where " & _
          "[Manufacturer ID] = " & Trim$(Str$(THIS_MANU_ID))
      Set Rs1 = _
          DB_Carbon.OpenRecordset(Current_Criteria)
      Rs1.Delete
      Rs1.Close
      'PERFORM DELETION OF ADSORBENT RECORDS.
      Current_Criteria = "select * from [Carbon Data] where " & _
          "[Manufacturer ID] = " & Trim$(Str$(THIS_MANU_ID))
      Set Rs1 = _
          DB_Carbon.OpenRecordset(Current_Criteria)
      On Error Resume Next
      Rs1.MoveFirst
      Rs1.MoveLast
      Rs1.MoveFirst
      NumRecords = Rs1.RecordCount
      On Error GoTo err_mnuManufacturerItem_Click
      If (NumRecords > 0) Then
        Do Until Rs1.EOF
          Rs1.Delete
          Rs1.MoveNext
        Loop
      End If
      Rs1.Close
      'REDISPLAY WINDOW.
      Call populate_lstManu
      'DISPLAY TOTAL.
      Call Show_Message("A total of " & Trim$(Str$(NumRecords)) & _
          " adsorbent records were deleted.")
  End Select
  Exit Sub
exit_err_mnuManufacturerItem_Click:
  Exit Sub
err_mnuManufacturerItem_Click:
  Call Show_Trapped_Error("mnuManufacturerItem_Click")
  Resume exit_err_mnuManufacturerItem_Click
End Sub


Private Sub optPhase_Click(Index As Integer, Value As Integer)
  Call lstManu_Click
End Sub





