VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{B16553C3-06DB-101B-85B2-0000C009BE81}#1.0#0"; "Spin32.ocx"
Begin VB.Form frmD3_PrimaryClarifier 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enter Parameters [Primary Clarifier]"
   ClientHeight    =   7110
   ClientLeft      =   3285
   ClientTop       =   1545
   ClientWidth     =   7500
   ControlBox      =   0   'False
   HelpContextID   =   4000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   Begin Threed.SSFrame SSFrame3 
      Height          =   2955
      Left            =   6030
      TabIndex        =   26
      Top             =   6840
      Visible         =   0   'False
      Width           =   3135
      _Version        =   65536
      _ExtentX        =   5530
      _ExtentY        =   5212
      _StockProps     =   14
      Caption         =   "Unused -- Invisible"
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ComboBox cboUnits 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   450
         Width           =   1500
      End
      Begin VB.ComboBox cboUnits 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   390
         Style           =   2  'Dropdown List
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1050
         Width           =   1500
      End
      Begin VB.Label lblData 
         Caption         =   "lblData(0).caption"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   27
         Top             =   630
         Width           =   2805
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   3045
      Left            =   240
      TabIndex        =   7
      Top             =   2040
      Width           =   6975
      _Version        =   65536
      _ExtentX        =   12303
      _ExtentY        =   5371
      _StockProps     =   14
      Caption         =   "Clarifier Specifications:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   5
         Left            =   2955
         TabIndex        =   34
         Text            =   "txtData(5)"
         Top             =   2580
         Width           =   1995
      End
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   2955
         TabIndex        =   24
         Text            =   "txtData(4)"
         Top             =   2190
         Width           =   1995
      End
      Begin VB.ComboBox cboUnits 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   5025
         Style           =   2  'Dropdown List
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   2160
         Width           =   1500
      End
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   2955
         TabIndex        =   21
         Text            =   "txtData(3)"
         Top             =   1800
         Width           =   1995
      End
      Begin VB.ComboBox cboUnits 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   5025
         Style           =   2  'Dropdown List
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1770
         Width           =   1500
      End
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   2955
         TabIndex        =   18
         Text            =   "txtData(2)"
         Top             =   1410
         Width           =   1995
      End
      Begin VB.ComboBox cboUnits 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   5025
         Style           =   2  'Dropdown List
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1380
         Width           =   1500
      End
      Begin Threed.SSFrame SSFrame2 
         Height          =   1005
         Left            =   90
         TabIndex        =   11
         Top             =   270
         Width           =   6795
         _Version        =   65536
         _ExtentX        =   11986
         _ExtentY        =   1773
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.ComboBox cboUnits 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   1
            Left            =   4935
            Style           =   2  'Dropdown List
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   480
            Width           =   1500
         End
         Begin VB.TextBox txtData 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   2865
            TabIndex        =   14
            Text            =   "txtData(1)"
            Top             =   510
            Width           =   1995
         End
         Begin Threed.SSOption opt_IsCovered 
            Height          =   345
            Index           =   0
            Left            =   1080
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   180
            Width           =   1500
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   609
            _StockProps     =   78
            Caption         =   "Uncovered"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
         End
         Begin Threed.SSOption opt_IsCovered 
            Height          =   345
            Index           =   1
            Left            =   1080
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   495
            Width           =   1500
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   609
            _StockProps     =   78
            Caption         =   "Covered"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
         End
         Begin VB.Label lblData 
            Caption         =   "Ventilation Rate:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   2880
            TabIndex        =   16
            Top             =   225
            Width           =   2805
         End
      End
      Begin VB.Label lblDataUnits 
         Alignment       =   2  'Center
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   5070
         TabIndex        =   37
         Top             =   2610
         Width           =   1365
      End
      Begin VB.Label lblData 
         Alignment       =   1  'Right Justify
         Caption         =   "Percentage Removal:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   30
         TabIndex        =   35
         Top             =   2610
         Width           =   2805
      End
      Begin VB.Label lblData 
         Alignment       =   1  'Right Justify
         Caption         =   "Wastage Flow Rate:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   30
         TabIndex        =   25
         Top             =   2220
         Width           =   2805
      End
      Begin VB.Label lblData 
         Alignment       =   1  'Right Justify
         Caption         =   "Volume:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   30
         TabIndex        =   22
         Top             =   1830
         Width           =   2805
      End
      Begin VB.Label lblData 
         Alignment       =   1  'Right Justify
         Caption         =   "Basin Depth:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   30
         TabIndex        =   19
         Top             =   1440
         Width           =   2805
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   6705
      Width           =   7500
      _Version        =   65536
      _ExtentX        =   13229
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
         TabIndex        =   1
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
         TabIndex        =   2
         Top             =   60
         Width           =   5000
         _Version        =   65536
         _ExtentX        =   8819
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
   Begin Threed.SSCommand cmdCancelOK 
      Height          =   495
      Index           =   1
      Left            =   5940
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Click here to save the changes to this window"
      Top             =   630
      Width           =   1305
      _Version        =   65536
      _ExtentX        =   2302
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
      Left            =   5940
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Click here to abandon any changes on this window"
      Top             =   150
      Width           =   1305
      _Version        =   65536
      _ExtentX        =   2302
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
   Begin Threed.SSCommand cmdCancelOK 
      Height          =   495
      Index           =   2
      Left            =   5940
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Click here for help"
      Top             =   1290
      Width           =   1305
      _Version        =   65536
      _ExtentX        =   2302
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "&Help"
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
   Begin Threed.SSFrame SSFrame6 
      Height          =   915
      Left            =   240
      TabIndex        =   6
      Top             =   1050
      Width           =   2685
      _Version        =   65536
      _ExtentX        =   4736
      _ExtentY        =   1614
      _StockProps     =   14
      Caption         =   "Number of Clarifiers:"
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   180
         TabIndex        =   10
         Text            =   "txtData(0)"
         Top             =   390
         Width           =   1425
      End
      Begin Spin.SpinButton spnData 
         Height          =   300
         Index           =   10
         Left            =   1620
         TabIndex        =   9
         Top             =   390
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   529
         _StockProps     =   73
      End
   End
   Begin Threed.SSFrame SSFrame4 
      Height          =   1275
      Left            =   240
      TabIndex        =   29
      Top             =   5190
      Width           =   6975
      _Version        =   65536
      _ExtentX        =   12303
      _ExtentY        =   2249
      _StockProps     =   14
      Caption         =   "Removal Mechanisms:"
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ComboBox cbo_RemovalMechanism 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   780
         Width           =   3705
      End
      Begin VB.ComboBox cbo_RemovalMechanism 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   360
         Width           =   3705
      End
      Begin VB.Label lbl_cbo_RemovalMechanism 
         Alignment       =   1  'Right Justify
         Caption         =   "Volatilization:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   960
         TabIndex        =   33
         Top             =   840
         Width           =   1605
      End
      Begin VB.Label lbl_cbo_RemovalMechanism 
         Alignment       =   1  'Right Justify
         Caption         =   "Sorption:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   960
         TabIndex        =   31
         Top             =   420
         Width           =   1605
      End
   End
   Begin VB.Label Label1 
      Caption         =   $"D3_PrimaryClarifier.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   300
      TabIndex        =   8
      Top             =   150
      Width           =   5445
   End
End
Attribute VB_Name = "frmD3_PrimaryClarifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim USER_HIT_CANCEL As Boolean
Dim USER_HIT_OK As Boolean
Dim frmD3_PrimaryClarifier_Is_Dirty As Boolean

Dim Temp_Plant As TYPE_PlantDiagram

Public HALT_opt_IsCovered As Boolean
Public HALT_cbo_RemovalMechanism As Boolean





Const frmD3_PrimaryClarifier_declarations_end = True


Sub frmD3_PrimaryClarifier_Edit( _
    OUTPUT_Raise_Dirty_Flag As Boolean)
  Temp_Plant = NowProj.Plant
  frmD3_PrimaryClarifier.Show 1
  If (USER_HIT_OK) Then
    OUTPUT_Raise_Dirty_Flag = True
    NowProj.Plant = Temp_Plant
  Else
    OUTPUT_Raise_Dirty_Flag = False
  End If
End Sub


Sub frmD3_PrimaryClarifier_PopulateUnits()
Dim Frm As Form
Set Frm = frmD3_PrimaryClarifier
  '
  ' MAIN DATA BLOCK.
  '
  With Temp_Plant.PrimaryClarifier
    Call unitsys_register(Frm, lblData(0), txtData(0), Nothing, "", _
        "", "", "0", "0", 100#, False)
    Call unitsys_register(Frm, lblData(1), txtData(1), cboUnits(1), "flow_volumetric", _
        .UnitsOfDisplay(1), "L/min", "", "", 100#, True)
    Call unitsys_register(Frm, lblData(2), txtData(2), cboUnits(2), "length", _
        .UnitsOfDisplay(2), "m", "", "", 100#, True)
    Call unitsys_register(Frm, lblData(3), txtData(3), cboUnits(3), "volume", _
        .UnitsOfDisplay(3), "liter", "", "", 100#, True)
    Call unitsys_register(Frm, lblData(4), txtData(4), cboUnits(4), "flow_volumetric", _
        .UnitsOfDisplay(4), "L/d", "", "", 100#, True)
    Call unitsys_register(Frm, lblData(5), txtData(5), Nothing, "", _
        "", "", "", "", 100#, False)
  End With
End Sub
Sub Store_Unit_Settings()
Dim i As Integer
  With Temp_Plant.PrimaryClarifier
    For i = 1 To 4
      .UnitsOfDisplay(i) = unitsys_get_units(cboUnits(i))
    Next i
  End With
End Sub


Sub Local_DirtyStatus_Set(DirtyFlag As Boolean, NewSetting As Boolean)
  Call Global_DirtyStatus_Set(Me, DirtyFlag, NewSetting)
End Sub
Sub Local_GenericStatus_Set(NewString As String)
  Call Global_GenericStatus_Set(Me, NewString)
End Sub


Sub Populate_cbo_RemovalMechanism()
Dim Ctl As Control
  HALT_cbo_RemovalMechanism = True
  Set Ctl = cbo_RemovalMechanism(0)
  Ctl.Clear
  Ctl.AddItem "Dobbs": Ctl.ItemData(Ctl.NewIndex) = _
      PRIMCLARIF_SORPTION_REMOVAL_DOBBS
  Ctl.AddItem "Matter-Muller": Ctl.ItemData(Ctl.NewIndex) = _
      PRIMCLARIF_SORPTION_REMOVAL_MATTER_MULLER
  Set Ctl = cbo_RemovalMechanism(1)
  Ctl.Clear
  Ctl.AddItem "Mackay & Yeun": Ctl.ItemData(Ctl.NewIndex) = _
      PRIMCLARIF_VOLATILIZATION_REMOVAL_MACKAY_YEUN
  Ctl.AddItem "KLA": Ctl.ItemData(Ctl.NewIndex) = _
      PRIMCLARIF_VOLATILIZATION_REMOVAL_KLA
  HALT_cbo_RemovalMechanism = False
End Sub


Private Sub cbo_RemovalMechanism_Click(Index As Integer)
Dim Ctl As Control
Set Ctl = cbo_RemovalMechanism(Index)
  If (HALT_cbo_RemovalMechanism = True) Then Exit Sub
  If (Val(Ctl.Tag) = Ctl.ListIndex) Then Exit Sub
  With Temp_Plant.PrimaryClarifier
    Select Case Index
      Case 0: .SorptionRemovalMethod = Ctl.ItemData(Ctl.ListIndex)
      Case 1: .VolatilizationRemovalMechanism = Ctl.ItemData(Ctl.ListIndex)
    End Select
  End With
  'RAISE DIRTY FLAG AND REFRESH WINDOW.
  Call Local_DirtyStatus_Set(frmD3_PrimaryClarifier_Is_Dirty, True)
  Call frmD3_PrimaryClarifier_Refresh(Temp_Plant)
End Sub


Private Sub cboUnits_Click(Index As Integer)
Dim Ctl As Control
Set Ctl = cboUnits(Index)
  Call unitsys_control_cbox_click(Ctl)
End Sub
Private Sub cboUnits_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = Global_TextKeyPress(KeyAscii)
End Sub


Private Sub cmdCancelOK_Click(Index As Integer)
Dim i As Integer
  Select Case Index
    Case 0:     'CANCEL.
      USER_HIT_CANCEL = True
      USER_HIT_OK = False
      Unload Me
      Exit Sub
    Case 1:     'OK.
      '
      ' STORE ALL UNIT SETTINGS.
      '
      Call Store_Unit_Settings
      '
      ' EXIT OUT OF HERE.
      '
      USER_HIT_CANCEL = False
      USER_HIT_OK = True
      Unload Me
      Exit Sub
    Case 2:     'HELP.
      SendKeys "{F1}"
  End Select
End Sub


Private Sub Form_Load()
  '
  ' MISC INITS.
  '
  Call CenterOnForm(Me, frmMain)
  Call Local_DirtyStatus_Set(frmD3_PrimaryClarifier_Is_Dirty, False)
  Call Local_GenericStatus_Set("")
  HALT_opt_IsCovered = False
  HALT_cbo_RemovalMechanism = False
  Call Populate_cbo_RemovalMechanism
  '
  ' POPULATE UNIT CONTROLS.
  '
  Call frmD3_PrimaryClarifier_PopulateUnits
  '
  ' REFRESH DISPLAY.
  '
  Call frmD3_PrimaryClarifier_Refresh(Temp_Plant)
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Call unitsys_unregister_all_on_form(Me)
End Sub


Private Sub opt_IsCovered_Click(Index As Integer, Value As Integer)
Dim Ctl0 As Control
Dim Ctl1 As Control
Set Ctl0 = opt_IsCovered(0)
Set Ctl1 = opt_IsCovered(1)
Dim NewTag As Integer
Dim NewSetting As Integer
  If (HALT_opt_IsCovered) Then Exit Sub
  NewTag = Index
  If (CInt(Val(Ctl0.Tag)) <> NewTag) Then
    NewSetting = IIf(NewTag = 0, False, True)
    With Temp_Plant.PrimaryClarifier
      .IsCovered = NewSetting
    End With
    'RAISE DIRTY FLAG AND REFRESH WINDOW.
    Call Local_DirtyStatus_Set(frmD3_PrimaryClarifier_Is_Dirty, True)
    Call frmD3_PrimaryClarifier_Refresh(Temp_Plant)
  End If
End Sub


Private Sub spnData_SpinDown(Index As Integer)
Dim Made_Dirty As Boolean
  Made_Dirty = False
  With Temp_Plant.PrimaryClarifier
    If (.Count > 1) Then
      .Count = .Count - 1
      Made_Dirty = True
    End If
  End With
  If (Made_Dirty) Then
    '
    ' THROW DIRTY FLAG AND REFRESH WINDOW.
    '
    Call Local_DirtyStatus_Set(frmD3_PrimaryClarifier_Is_Dirty, True)
    Call frmD3_PrimaryClarifier_Refresh(Temp_Plant)
  End If
End Sub
Private Sub spnData_SpinUp(Index As Integer)
Dim Made_Dirty As Boolean
  Made_Dirty = False
  With Temp_Plant.PrimaryClarifier
    If (.Count < PRIMCLARIF_MAX_CLARIFIERS) Then
      .Count = .Count + 1
      Made_Dirty = True
    End If
  End With
  If (Made_Dirty) Then
    '
    ' THROW DIRTY FLAG AND REFRESH WINDOW.
    '
    Call Local_DirtyStatus_Set(frmD3_PrimaryClarifier_Is_Dirty, True)
    Call frmD3_PrimaryClarifier_Refresh(Temp_Plant)
  End If
End Sub


Private Sub txtData_GotFocus(Index As Integer)
Dim Ctl As Control
Set Ctl = txtData(Index)
Dim StatusMessagePanel As String
  Call unitsys_control_txtx_gotfocus(Ctl)
  Select Case Index
    '
    ' MAIN DATA BLOCK.
    '
    Case 0:
      StatusMessagePanel = ""
  End Select
  Call Local_GenericStatus_Set(StatusMessagePanel)
End Sub
Private Sub txtData_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = Global_NumericKeyPress(KeyAscii)
End Sub
Private Sub txtData_LostFocus(Index As Integer)
Dim NewValue_Okay As Integer
Dim NewValue As Double
Dim Ctl As Control
Set Ctl = txtData(Index)
Dim Val_Low As Double
Dim Val_High As Double
Dim Raise_Dirty_Flag As Boolean
Dim Too_Small As Integer
  'NOTE: LOW AND HIGH VALUES IN BASE UNITS
  Select Case Index
    ' MAIN DATA BLOCK.
    Case 0: Val_Low = CDbl(1): Val_High = CDbl(PRIMCLARIF_MAX_CLARIFIERS)
    Case 1: Val_Low = 1E-20: Val_High = 1E+20
    Case 2: Val_Low = 1E-20: Val_High = 1E+20
    Case 3: Val_Low = 1E-20: Val_High = 1E+20
    Case 4: Val_Low = 1E-20: Val_High = 1E+20
    Case 5: Val_Low = 0#: Val_High = 100#
  End Select
  NewValue_Okay = False
  If (unitsys_control_txtx_lostfocus_validate(Ctl, Val_Low, Val_High, NewValue, Raise_Dirty_Flag)) Then
    NewValue_Okay = True
  End If
  Call unitsys_control_txtx_lostfocus(Ctl, NewValue)
  Call Local_GenericStatus_Set("")
  If (NewValue_Okay) Then
    If (Raise_Dirty_Flag) Then
      'STORE TO MEMORY.
      With Temp_Plant.PrimaryClarifier
        Select Case Index
          '
          ' MAIN DATA BLOCK.
          '
          Case 0: .Count = CInt(NewValue)
          Case 1: .VentilationRate = NewValue
          Case 2: .Depth = NewValue
          Case 3: .Volume = NewValue
          Case 4: .WastageFlow = NewValue
          Case 5: .PercentageRemoval = NewValue
        End Select
      End With
      If (Raise_Dirty_Flag) Then
        'THROW DIRTY FLAG.
        Call Local_DirtyStatus_Set(frmD3_PrimaryClarifier_Is_Dirty, True)
      End If
      'REFRESH WINDOW.
      Call frmD3_PrimaryClarifier_Refresh(Temp_Plant)
    End If
  End If
End Sub

