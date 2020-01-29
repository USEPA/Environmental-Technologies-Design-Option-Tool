VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmEditAdsorber 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Adsorber Database"
   ClientHeight    =   6600
   ClientLeft      =   840
   ClientTop       =   1095
   ClientWidth     =   9135
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   9135
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   8640
      ScaleHeight     =   795
      ScaleWidth      =   1275
      TabIndex        =   33
      Top             =   4320
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
      Left            =   7560
      TabIndex        =   32
      ToolTipText     =   "Click here to print current screen to selected printer"
      Top             =   5280
      Width           =   1455
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   2415
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3735
      _Version        =   65536
      _ExtentX        =   6588
      _ExtentY        =   4260
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
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   330
         Width           =   3495
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   3735
      Left            =   60
      TabIndex        =   1
      Top             =   2670
      Width           =   3735
      _Version        =   65536
      _ExtentX        =   6588
      _ExtentY        =   6588
      _StockProps     =   14
      Caption         =   "Select an Adsorber:"
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
         Height          =   2955
         Left            =   120
         TabIndex        =   3
         Top             =   660
         Width           =   3495
      End
      Begin Threed.SSOption optPhase 
         Height          =   375
         Index           =   2
         Left            =   2040
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   210
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
      Begin Threed.SSOption optPhase 
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   210
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
   End
   Begin Threed.SSFrame SSFrame4 
      Height          =   4215
      Left            =   3900
      TabIndex        =   2
      Top             =   60
      Width           =   5145
      _Version        =   65536
      _ExtentX        =   9075
      _ExtentY        =   7435
      _StockProps     =   14
      Caption         =   "Adsorber Properties:"
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
         Caption         =   "Internal Area:"
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
         Left            =   60
         TabIndex        =   29
         Top             =   285
         Width           =   2055
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   2160
         TabIndex        =   28
         Top             =   270
         Width           =   1395
      End
      Begin VB.Label lblUnits 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "(ft²)"
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
         Left            =   3600
         TabIndex        =   27
         Top             =   285
         Width           =   915
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Maximum Capacity:"
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
         Left            =   60
         TabIndex        =   26
         Top             =   585
         Width           =   2055
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   2160
         TabIndex        =   25
         Top             =   570
         Width           =   1395
      End
      Begin VB.Label lblUnits 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "(lbs)"
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
         Left            =   3600
         TabIndex        =   24
         Top             =   585
         Width           =   915
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Outside Diameter:"
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
         Left            =   60
         TabIndex        =   23
         Top             =   885
         Width           =   2055
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   2160
         TabIndex        =   22
         Top             =   870
         Width           =   1395
      End
      Begin VB.Label lblUnits 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "(feet)"
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
         Left            =   3600
         TabIndex        =   21
         Top             =   885
         Width           =   915
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Design Pressure:"
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
         Left            =   60
         TabIndex        =   20
         Top             =   1185
         Width           =   2055
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   2160
         TabIndex        =   19
         Top             =   1170
         Width           =   1395
      End
      Begin VB.Label lblUnits 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "(psig)"
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
         Left            =   3600
         TabIndex        =   18
         Top             =   1185
         Width           =   915
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Design Flow Range:"
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
         Left            =   60
         TabIndex        =   17
         Top             =   1485
         Width           =   2055
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   2160
         TabIndex        =   16
         Top             =   1470
         Width           =   1395
      End
      Begin VB.Label lblUnits 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "(see code)"
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
         Left            =   3600
         TabIndex        =   15
         Top             =   1485
         Width           =   915
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Default Flow Rate:"
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
         Left            =   60
         TabIndex        =   14
         Top             =   1785
         Width           =   2055
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   2160
         TabIndex        =   13
         Top             =   1770
         Width           =   1395
      End
      Begin VB.Label lblUnits 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "(see code)"
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
         Left            =   3600
         TabIndex        =   12
         Top             =   1785
         Width           =   915
      End
      Begin VB.Label lblDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Notes:"
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
         Left            =   600
         TabIndex        =   11
         Top             =   2100
         Width           =   1035
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   600
         TabIndex        =   10
         Top             =   2370
         Width           =   3675
      End
      Begin VB.Label lblUnits 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "(lb/ft³)"
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
         Left            =   3600
         TabIndex        =   9
         Top             =   2865
         Width           =   915
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblData(7)"
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
         Left            =   2160
         TabIndex        =   8
         Top             =   2850
         Width           =   1395
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Default Bulk Density:"
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
         Left            =   60
         TabIndex        =   7
         Top             =   2865
         Width           =   2055
      End
   End
   Begin Threed.SSCommand cmdOK 
      Height          =   615
      Left            =   3900
      TabIndex        =   30
      Top             =   5790
      Width           =   4035
      _Version        =   65536
      _ExtentX        =   7117
      _ExtentY        =   1085
      _StockProps     =   78
      Caption         =   "&Use These Adsorber Specifications"
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
      Left            =   7920
      TabIndex        =   31
      Top             =   5790
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
   Begin VB.Menu mnuAdsorber 
      Caption         =   "&Adsorber"
      Begin VB.Menu mnuAdsorberItem 
         Caption         =   "&New"
         Index           =   1
      End
      Begin VB.Menu mnuAdsorberItem 
         Caption         =   "&Edit Current"
         Index           =   2
      End
      Begin VB.Menu mnuAdsorberItem 
         Caption         =   "&Delete Current"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmEditAdsorber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim frmEditAdsorber_Cancelled As Integer
Dim frmEditAdsorber_RunMode As Integer
Const frmEditAdsorber_RunMode_QUERY_DATABASE = 1
Const frmEditAdsorber_RunMode_EDIT_DATABASE = 2

Dim USER_HIT_CANCEL As Boolean
Dim USER_HIT_USE_THESE As Boolean




Const frmEditAdsorber_declarations_end = True


Sub frmEditAdsorber_QueryDatabase( _
    OUTPUT_User_Transferred_Data As Boolean)
  frmEditAdsorber_RunMode = frmEditAdsorber_RunMode_QUERY_DATABASE
  frmEditAdsorber.Show 1
  If (USER_HIT_USE_THESE) Then
    OUTPUT_User_Transferred_Data = True
  Else
    OUTPUT_User_Transferred_Data = False
  End If
End Sub
Sub frmEditAdsorber_EditDatabase()
  frmEditAdsorber_RunMode = frmEditAdsorber_RunMode_EDIT_DATABASE
  frmEditAdsorber.Show 1
End Sub
Sub LOCAL___Reset_DemoVersionDisablings()
  If (IsThisADemo() = True) Then
    cmdOK.Enabled = False
    mnuManufacturerItem(1).Enabled = False
    mnuManufacturerItem(2).Enabled = False
    mnuManufacturerItem(3).Enabled = False
    mnuAdsorberItem(1).Enabled = False
    mnuAdsorberItem(2).Enabled = False
    mnuAdsorberItem(3).Enabled = False
  End If
End Sub


Private Function adsorber_db_AssignUniqueID() As Integer
Dim this_try As Integer
Dim i As Integer
Dim Found As Integer
  Do While (1 = 1)
    this_try = Int(Rnd(1) * 32000) + 1
    Found = False
    For i = 1 To adsorber_db_num_manufacturers
      If (CInt(adsorber_db_manufacturers(i).UniqueID) = this_try) Then
        Found = True
        Exit For
      End If
    Next i
    If (Not Found) Then
      adsorber_db_AssignUniqueID = this_try
      Exit Function
    End If
  Loop
End Function


Private Sub adsorber_db_displayall()
Dim i As Integer
Dim N As Integer
  'POPULATE LISTBOX lstManu:
  lstManu.Clear
  For i = 1 To adsorber_db_num_manufacturers
    lstManu.AddItem adsorber_db_manufacturers(i).Name
    N = lstManu.NewIndex
    lstManu.ItemData(N) = CInt(adsorber_db_manufacturers(i).UniqueID)
  Next i
End Sub


'Load all adsorber DB entries.
Private Sub adsorber_db_loadall()
Dim f As Integer
Dim i As Integer
Dim fn As String
  
  On Error GoTo err_adsorber_db_loadall
  
  'LOAD MANUFACTURERS.
  fn = Database_Path & "\beds2.txt"
  f = FreeFile
  Open fn For Input As #f
  Input #f, adsorber_db_num_manufacturers
  If (adsorber_db_num_manufacturers <> 0) Then
    ReDim adsorber_db_manufacturers(1 To adsorber_db_num_manufacturers)
    For i = 1 To adsorber_db_num_manufacturers
      Input #f, adsorber_db_manufacturers(i).UniqueID, adsorber_db_manufacturers(i).Name
    Next i
  End If
  Close #f

  'LOAD ADSORBERS.
  fn = Database_Path & "\beds1.txt"
  f = FreeFile
  Open fn For Input As #f
  Input #f, adsorber_db_num_adsorbers
  If (adsorber_db_num_adsorbers <> 0) Then
    ReDim adsorber_db_adsorbers(1 To adsorber_db_num_adsorbers)
    For i = 1 To adsorber_db_num_adsorbers
      Input #f, adsorber_db_adsorbers(i).UniqueID_Manufacturer, adsorber_db_adsorbers(i).Phase, adsorber_db_adsorbers(i).PartNumber, adsorber_db_adsorbers(i).InternalArea, adsorber_db_adsorbers(i).MaxCapacity, adsorber_db_adsorbers(i).OutsideDiameter, adsorber_db_adsorbers(i).DesignPressure, adsorber_db_adsorbers(i).DesignFlowRange, adsorber_db_adsorbers(i).DefaultFlowRate, adsorber_db_adsorbers(i).Note
      'MsgBox CStr(adsorber_db_adsorbers(i).PartNumber)
    Next i
  End If
  Close #f
exit_err_adsorber_db_loadall:
  Exit Sub
err_adsorber_db_loadall:
  Call Show_Trapped_Error("Load Adsorber Database")
  Resume exit_err_adsorber_db_loadall
End Sub


'RETURNS:
'  -1 if not found
'  index if found
Private Function adsorber_db_lookup_UniqueID_Manufacturer(search_for As Integer) As Integer
Dim i As Integer
Dim Found As Integer
  Found = False
  For i = 1 To adsorber_db_num_manufacturers
    If (CInt(adsorber_db_manufacturers(i).UniqueID) = search_for) Then
      Found = True
      Exit For
    End If
  Next i
  If (Found) Then
    adsorber_db_lookup_UniqueID_Manufacturer = i
  Else
    adsorber_db_lookup_UniqueID_Manufacturer = -1
  End If
End Function


Private Sub adsorber_db_saveall()
Dim f As Integer
Dim i As Integer
Dim fn As String

  On Error GoTo err_adsorber_db_saveall

  'SAVE MANUFACTURERS.
  fn = Database_Path & "\beds2.txt"
  f = FreeFile
  Open fn For Output As #f
  Write #f, adsorber_db_num_manufacturers
  'ReDim adsorber_db_manufacturers(1 To adsorber_db_num_manufacturers)
  For i = 1 To adsorber_db_num_manufacturers
    Write #f, adsorber_db_manufacturers(i).UniqueID, adsorber_db_manufacturers(i).Name
  Next i
  Close #f

  'SAVE ADSORBERS.
  fn = Database_Path & "\beds1.txt"
  f = FreeFile
  Open fn For Output As #f
  Write #f, adsorber_db_num_adsorbers
  'ReDim adsorber_db_adsorbers(1 To adsorber_db_num_adsorbers)
  For i = 1 To adsorber_db_num_adsorbers
    Write #f, adsorber_db_adsorbers(i).UniqueID_Manufacturer, _
        adsorber_db_adsorbers(i).Phase, _
        Trim$(adsorber_db_adsorbers(i).PartNumber), _
        Trim$(adsorber_db_adsorbers(i).InternalArea), _
        Trim$(adsorber_db_adsorbers(i).MaxCapacity), _
        Trim$(adsorber_db_adsorbers(i).OutsideDiameter), _
        Trim$(adsorber_db_adsorbers(i).DesignPressure), _
        Trim$(adsorber_db_adsorbers(i).DesignFlowRange), _
        Trim$(adsorber_db_adsorbers(i).DefaultFlowRate), _
        Trim$(adsorber_db_adsorbers(i).Note)
    'MsgBox CStr(adsorber_db_adsorbers(i).PartNumber)
  Next i
  Close #f

  Call clear_this_record
  lstName.Clear
  Call adsorber_db_loadall
  Call adsorber_db_displayall
exit_err_adsorber_db_saveall:
  Exit Sub
err_adsorber_db_saveall:
  Call Show_Trapped_Error("Save Adsorber Database")
  Resume exit_err_adsorber_db_saveall
End Sub


Private Sub clear_this_record()
Dim i As Integer
  For i = 0 To 6
    lblData(i) = ""
  Next i
End Sub


Private Sub cmdCancel_Click()
  'frmEditAdsorber_Cancelled = True
  USER_HIT_CANCEL = True
  USER_HIT_USE_THESE = False
  Unload Me
End Sub
Private Sub cmdOK_Click()
Dim N As Integer
Dim this_A As Double
Dim this_M As Double
Dim this_rhoB As Double
Dim New_D As Double
Dim new_V As Double
Dim New_L As Double
Dim this_Q As Double
Dim now_phase As Integer
  If (lstName.ListIndex < 0) Or (lstName.ListCount = 0) Then
    Call Show_Error("You must first select an adsorber.")
    Exit Sub
  End If
  N = lstName.ItemData(lstName.ListIndex)
  'CONVERT AREA FROM FT^2 TO M^2
  this_A = CDbl(adsorber_db_adsorbers(N).InternalArea)
  this_A = this_A / 10.7639104167
  'CONVERT MASS FROM LBS TO KG
  this_M = CDbl(adsorber_db_adsorbers(N).MaxCapacity)
  this_M = this_M / 2.20462262185
  'CONVERT BULK DENSITY FROM LBM/FT^3 TO KG/M^3
  this_rhoB = 28#
  this_rhoB = this_rhoB * 0.45359237 / 0.028316846592
  'CALCULATE INTERNAL DIAMETER IN M
  New_D = (4# * this_A / 3.14159) ^ 0.5
  'CALCULATE VOLUME IN M^3
  new_V = this_M / this_rhoB
  'CALCULATE LENGTH IN M
  New_L = new_V / this_A
  'CONVERT DEFAULT FLOW RATE TO M^3/S
  this_Q = CDbl(adsorber_db_adsorbers(N).DefaultFlowRate)
  If (optPhase(1).Value) Then
    now_phase = 1     'LIQUID PHASE
  Else
    now_phase = 2     'GAS PHASE
  End If
  If (now_phase = 1) Then
    'CONVERT FROM GAL/MIN TO M^3/S
    this_Q = this_Q * 0.003785411784 / 60#
  End If
  If (now_phase = 2) Then
    'CONVERT FROM FT^3/MIN TO M^3/S
    this_Q = this_Q * 0.028316846592 / 60#
  End If
  'TRANSFER PARAMETERS BACK TO MAIN SCREEN
  frmEditAdsorber_ReturnParameters.D = New_D
  frmEditAdsorber_ReturnParameters.L = New_L
  frmEditAdsorber_ReturnParameters.Q = this_Q
  frmEditAdsorber_ReturnParameters.M = this_M
  'frmEditAdsorber_Cancelled = False
  USER_HIT_CANCEL = False
  USER_HIT_USE_THESE = True
  Unload Me
End Sub


Private Sub Form_Load()
  'MISC INITS.
  Me.Height = 7290
  Me.Width = 9255
  Call CenterOnForm(Me, frmMain)
  If (frmEditAdsorber_RunMode = frmEditAdsorber_RunMode_QUERY_DATABASE) Then
    cmdOK.Visible = True
    cmdCancel.Visible = True
    lblData(7).Caption = "28.0"
    lblData(7).Visible = True
  Else
    cmdOK.Visible = False
    cmdCancel.Visible = True
    cmdCancel.Caption = "E&xit"
    lblData(7).Visible = False
    lblDesc(7).Visible = False
    lblUnits(7).Visible = False
  End If
  If (Bed.Phase = 0) Then
    optPhase(1).Value = True
    optPhase(2).Value = False
  Else
    optPhase(1).Value = False
    optPhase(2).Value = True
  End If
  Call adsorber_db_loadall
  Call clear_this_record
  Call adsorber_db_displayall
  ' DEMO SETTINGS.
  Call LOCAL___Reset_DemoVersionDisablings
End Sub


Private Sub lstManu_Click()
Dim thismanu As Integer
Dim i As Integer
Dim N As Integer
Dim now_phase As Integer
  Call clear_this_record
  'DISPLAY NAMES FOR THIS MANUFACTURER (IF ANY)
  lstName.Clear
  If (lstManu.ListIndex < 0) Then Exit Sub
  thismanu = lstManu.ItemData(lstManu.ListIndex)
  If (optPhase(1).Value) Then
    now_phase = 1     'LIQUID PHASE
  Else
    now_phase = 2     'GAS PHASE
  End If
  For i = 1 To adsorber_db_num_adsorbers
    If (thismanu = adsorber_db_adsorbers(i).UniqueID_Manufacturer) Then
      If (now_phase = adsorber_db_adsorbers(i).Phase) Then
        lstName.AddItem adsorber_db_adsorbers(i).PartNumber
        N = lstName.NewIndex
        lstName.ItemData(N) = i
      End If
    End If
  Next i
End Sub
Private Sub lstManu_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If ((Button And 2) = 2) Then
    Me.PopupMenu mnuManufacturer
  End If
End Sub
Private Sub lstName_Click()
Dim N As Integer
  'DISPLAY ADSORBER PROPERTIES:
  N = lstName.ItemData(lstName.ListIndex)
  lblData(0) = Trim$(adsorber_db_adsorbers(N).InternalArea)
  lblData(1) = Trim$(adsorber_db_adsorbers(N).MaxCapacity)
  lblData(2) = Trim$(adsorber_db_adsorbers(N).OutsideDiameter)
  lblData(3) = Trim$(adsorber_db_adsorbers(N).DesignPressure)
  lblData(4) = Trim$(adsorber_db_adsorbers(N).DesignFlowRange)
  lblData(5) = Trim$(adsorber_db_adsorbers(N).DefaultFlowRate)
  lblData(6) = Trim$(adsorber_db_adsorbers(N).Note)
End Sub
Private Sub lstName_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If ((Button And 2) = 2) Then
    Me.PopupMenu mnuAdsorber
  End If
End Sub


Private Sub mnuAdsorberItem_Click(Index As Integer)
Dim N As Integer
Dim i As Integer
Dim response As Integer
Dim msg As String
Dim n_manu As Integer
Dim now_phase As Integer
Dim USER_HIT_CANCEL As Boolean
  If (lstManu.ListIndex < 0) Or (lstManu.ListCount = 0) Then
    Call Show_Error("You must first select a manufacturer.")
    Exit Sub
  End If
  n_manu = lstManu.ItemData(lstManu.ListIndex)
  n_manu = adsorber_db_lookup_UniqueID_Manufacturer(n_manu)
  If (Index = 2) Or (Index = 3) Then
    If (lstName.ListIndex < 0) Or (lstName.ListCount = 0) Then
      Call Show_Error("You must first select an adsorber.")
      Exit Sub
    End If
    N = lstName.ItemData(lstName.ListIndex)
  End If
  If (optPhase(1).Value) Then
    now_phase = 1     'LIQUID PHASE
  Else
    now_phase = 2     'GAS PHASE
  End If
  Select Case Index
    Case 1:     'new
      Call frmEditAdsorberData.frmEditAdsorberData_AddNew( _
          now_phase, _
          USER_HIT_CANCEL)
      If (Not USER_HIT_CANCEL) Then
        frmEditAdsorberData_Record.UniqueID_Manufacturer = CInt(adsorber_db_manufacturers(n_manu).UniqueID)
        adsorber_db_num_adsorbers = adsorber_db_num_adsorbers + 1
        N = adsorber_db_num_adsorbers
        ReDim Preserve adsorber_db_adsorbers(1 To N)
        adsorber_db_adsorbers(N) = frmEditAdsorberData_Record
        Call adsorber_db_saveall
      End If
    Case 2:     'edit current
      frmEditAdsorberData_Record = adsorber_db_adsorbers(N)
      Call frmEditAdsorberData.frmEditAdsorberData_Edit( _
          adsorber_db_adsorbers(N).Phase, _
          USER_HIT_CANCEL)
      If (Not USER_HIT_CANCEL) Then
        adsorber_db_adsorbers(N) = frmEditAdsorberData_Record
        Call adsorber_db_saveall
      End If
    Case 3:     'delete current
      msg = "Do you really want to delete adsorber '" & Trim$(adsorber_db_adsorbers(N).PartNumber) & "' ?"
      response = MsgBox(msg, vbCritical + vbYesNo, AppName_For_Display_Short)
      If response = vbNo Then Exit Sub
      'PERFORM DELETION
      For i = N To adsorber_db_num_adsorbers - 1
        adsorber_db_adsorbers(i) = adsorber_db_adsorbers(i + 1)
      Next i
      adsorber_db_num_adsorbers = adsorber_db_num_adsorbers - 1
      'SAVE MANUFACTURER FILE
      Call adsorber_db_saveall
  End Select
End Sub


Private Sub mnuManufacturerItem_Click(Index As Integer)
Dim N As Integer
Dim i As Integer
Dim response As Integer
Dim msg As String
Dim new_UniqueID As Integer
Dim NewName As String
Dim USER_HIT_CANCEL As Boolean
  If (Index = 2) Or (Index = 3) Then
    If (lstManu.ListIndex < 0) Or (lstManu.ListCount = 0) Then
      Call Show_Error("You must first select a manufacturer.")
      Exit Sub
    End If
    N = lstManu.ItemData(lstManu.ListIndex)
    N = adsorber_db_lookup_UniqueID_Manufacturer(N)
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
      new_UniqueID = adsorber_db_AssignUniqueID()
      adsorber_db_num_manufacturers = adsorber_db_num_manufacturers + 1
      N = adsorber_db_num_manufacturers
      ReDim Preserve adsorber_db_manufacturers(1 To N)
      adsorber_db_manufacturers(N).Name = Trim$(CStr(NewName))
      adsorber_db_manufacturers(N).UniqueID = Trim$(Str$(new_UniqueID))
      Call adsorber_db_saveall
    Case 2:     'edit current
      'If Number_Of_Manufacturers = 0 Then
      '  MsgBox "There is no manufacturer name to edit.", MB_ICONEXCLAMATION, AppName_For_Display_long
      '  Exit Sub
      'End If
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
      adsorber_db_manufacturers(N).Name = NewName
      Call adsorber_db_saveall
    Case 3:     'delete current
      msg = "Do you really want to delete manufacturer '" & Trim$(adsorber_db_manufacturers(N).Name) & "' ?"
      response = MsgBox(msg, vbCritical + vbYesNo, AppName_For_Display_Long)
      If response = vbNo Then Exit Sub
      'PERFORM DELETION
      For i = N To adsorber_db_num_manufacturers - 1
        adsorber_db_manufacturers(i) = adsorber_db_manufacturers(i + 1)
      Next i
      adsorber_db_num_manufacturers = adsorber_db_num_manufacturers - 1
      'SAVE MANUFACTURER FILE
      Call adsorber_db_saveall
  End Select
End Sub


Private Sub optPhase_Click(Index As Integer, Value As Integer)
  If (optPhase(1).Value) Then
    'LIQUID PHASE
    lblUnits(4).Caption = "(gal/min)"
    lblUnits(5).Caption = "(gal/min)"
  Else
    'GAS PHASE
    lblUnits(4).Caption = "(ft³/min)"
    lblUnits(5).Caption = "(ft³/min)"
  End If
  If (lstManu.ListCount > 0) Then
    'UPDATE LIST OF NAMES:
    If (lstManu.ListIndex < 0) Then
      lstManu.ListIndex = 0
    End If
    Call lstManu_Click
  End If
End Sub


