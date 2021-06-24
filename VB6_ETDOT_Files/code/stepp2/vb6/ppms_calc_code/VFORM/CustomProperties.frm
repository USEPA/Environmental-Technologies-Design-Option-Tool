VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmCustomProperties 
   Caption         =   "Property Hierarchy Customization"
   ClientHeight    =   7065
   ClientLeft      =   4275
   ClientTop       =   795
   ClientWidth     =   10545
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   10545
   Begin Threed.SSPanel sspMainButtons 
      Height          =   525
      Left            =   5190
      TabIndex        =   40
      Top             =   6480
      Width           =   5055
      _Version        =   65536
      _ExtentX        =   8916
      _ExtentY        =   926
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
      Begin VB.CommandButton cmdCancelOK 
         Caption         =   "&Cancel"
         Height          =   345
         Index           =   0
         Left            =   1950
         TabIndex        =   43
         TabStop         =   0   'False
         ToolTipText     =   "Cancel changes (if any) and return to main window"
         Top             =   90
         Width           =   1485
      End
      Begin VB.CommandButton cmdCancelOK 
         Caption         =   "&Accept"
         Height          =   345
         Index           =   1
         Left            =   3450
         TabIndex        =   42
         TabStop         =   0   'False
         ToolTipText     =   "Accept changes (if any) and return to main window"
         Top             =   90
         Width           =   1485
      End
      Begin VB.CommandButton cmdRestoreDefaults 
         Caption         =   "Restore Defaults"
         Height          =   345
         Left            =   90
         TabIndex        =   41
         TabStop         =   0   'False
         ToolTipText     =   "Restore all default settings"
         Top             =   90
         Width           =   1485
      End
   End
   Begin Threed.SSPanel sspAll 
      Height          =   6315
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   10185
      _Version        =   65536
      _ExtentX        =   17965
      _ExtentY        =   11139
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
      Begin Threed.SSPanel sspTop 
         Height          =   2775
         Left            =   60
         TabIndex        =   1
         Top             =   90
         Width           =   9645
         _Version        =   65536
         _ExtentX        =   17013
         _ExtentY        =   4895
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
         Begin Threed.SSFrame ssfTop 
            Height          =   2565
            Left            =   90
            TabIndex        =   2
            Top             =   90
            Width           =   8505
            _Version        =   65536
            _ExtentX        =   15002
            _ExtentY        =   4524
            _StockProps     =   14
            Caption         =   "Selection and Ordering of Properties for Each Property Sheet:"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin Threed.SSFrame ssfTopLeft 
               Height          =   1935
               Left            =   90
               TabIndex        =   3
               Top             =   270
               Width           =   1425
               _Version        =   65536
               _ExtentX        =   2514
               _ExtentY        =   3413
               _StockProps     =   14
               Caption         =   "Property Sheets:"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Begin VB.ListBox lstTop_PropSheets 
                  Height          =   1230
                  ItemData        =   "CustomProperties.frx":0000
                  Left            =   90
                  List            =   "CustomProperties.frx":0002
                  TabIndex        =   4
                  Top             =   270
                  Width           =   1155
               End
            End
            Begin Threed.SSFrame ssfTopLeftCmds 
               Height          =   1935
               Left            =   1500
               TabIndex        =   5
               Top             =   270
               Width           =   1455
               _Version        =   65536
               _ExtentX        =   2566
               _ExtentY        =   3413
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
               Begin VB.CommandButton cmdTopLeftCmds 
                  Caption         =   "Add"
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   10
                  TabStop         =   0   'False
                  ToolTipText     =   "Add a new property sheet"
                  Top             =   330
                  Width           =   1200
               End
               Begin VB.CommandButton cmdTopLeftCmds 
                  Caption         =   "Delete"
                  Height          =   255
                  Index           =   1
                  Left            =   120
                  TabIndex        =   9
                  TabStop         =   0   'False
                  ToolTipText     =   "Delete a property sheet"
                  Top             =   600
                  Width           =   1200
               End
               Begin VB.CommandButton cmdTopLeftCmds 
                  Caption         =   "Move Up"
                  Height          =   255
                  Index           =   3
                  Left            =   120
                  TabIndex        =   8
                  TabStop         =   0   'False
                  ToolTipText     =   "Move a property sheet up the list"
                  Top             =   1140
                  Width           =   1200
               End
               Begin VB.CommandButton cmdTopLeftCmds 
                  Caption         =   "Move Down"
                  Height          =   255
                  Index           =   4
                  Left            =   120
                  TabIndex        =   7
                  TabStop         =   0   'False
                  ToolTipText     =   "Move a property sheet down the list"
                  Top             =   1410
                  Width           =   1200
               End
               Begin VB.CommandButton cmdTopLeftCmds 
                  Caption         =   "Rename"
                  Height          =   255
                  Index           =   2
                  Left            =   120
                  TabIndex        =   6
                  TabStop         =   0   'False
                  ToolTipText     =   "Rename a property sheet"
                  Top             =   870
                  Width           =   1200
               End
            End
            Begin Threed.SSFrame ssfTopRight 
               Height          =   2145
               Left            =   2940
               TabIndex        =   11
               Top             =   270
               Width           =   5055
               _Version        =   65536
               _ExtentX        =   8916
               _ExtentY        =   3784
               _StockProps     =   14
               Caption         =   "Properties for 'Sheet 1':"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Begin Threed.SSFrame ssfTopRight1 
                  Height          =   1695
                  Left            =   120
                  TabIndex        =   12
                  Top             =   270
                  Width           =   1245
                  _Version        =   65536
                  _ExtentX        =   2196
                  _ExtentY        =   2990
                  _StockProps     =   14
                  Caption         =   "Selected:"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Begin VB.ListBox lstTop_Props 
                     Height          =   1230
                     Left            =   90
                     MultiSelect     =   2  'Extended
                     TabIndex        =   13
                     Top             =   270
                     Width           =   1065
                  End
               End
               Begin Threed.SSFrame ssfTopRight2 
                  Height          =   1695
                  Left            =   2790
                  TabIndex        =   14
                  Top             =   270
                  Width           =   1905
                  _Version        =   65536
                  _ExtentX        =   3360
                  _ExtentY        =   2990
                  _StockProps     =   14
                  Caption         =   "Not Selected:"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Begin VB.ListBox lstTop_PropsAll 
                     Height          =   1035
                     Left            =   90
                     MultiSelect     =   2  'Extended
                     Sorted          =   -1  'True
                     TabIndex        =   15
                     Top             =   270
                     Width           =   1695
                  End
               End
               Begin Threed.SSFrame ssfTopRight1Cmds 
                  Height          =   1695
                  Left            =   1350
                  TabIndex        =   16
                  Top             =   270
                  Width           =   1455
                  _Version        =   65536
                  _ExtentX        =   2566
                  _ExtentY        =   2990
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
                  Begin VB.CommandButton cmdTopRight1Cmds 
                     Caption         =   "<<"
                     Height          =   255
                     Index           =   1
                     Left            =   735
                     TabIndex        =   22
                     TabStop         =   0   'False
                     ToolTipText     =   "Select all properties into property sheet"
                     Top             =   330
                     Width           =   585
                  End
                  Begin VB.CommandButton cmdTopRight1Cmds 
                     Caption         =   "<"
                     Height          =   255
                     Index           =   0
                     Left            =   120
                     TabIndex        =   21
                     TabStop         =   0   'False
                     ToolTipText     =   "Select property/properties into property sheet"
                     Top             =   330
                     Width           =   585
                  End
                  Begin VB.CommandButton cmdTopRight1Cmds 
                     Caption         =   ">>"
                     Height          =   255
                     Index           =   3
                     Left            =   735
                     TabIndex        =   20
                     TabStop         =   0   'False
                     ToolTipText     =   "Deselect all properties from property sheet"
                     Top             =   600
                     Width           =   585
                  End
                  Begin VB.CommandButton cmdTopRight1Cmds 
                     Caption         =   ">"
                     Height          =   255
                     Index           =   2
                     Left            =   120
                     TabIndex        =   19
                     TabStop         =   0   'False
                     ToolTipText     =   "Deselect property/properties from property sheet"
                     Top             =   600
                     Width           =   585
                  End
                  Begin VB.CommandButton cmdTopRight1Cmds 
                     Caption         =   "Move Up"
                     Height          =   255
                     Index           =   4
                     Left            =   120
                     TabIndex        =   18
                     TabStop         =   0   'False
                     ToolTipText     =   "Move a property up the list"
                     Top             =   870
                     Width           =   1200
                  End
                  Begin VB.CommandButton cmdTopRight1Cmds 
                     Caption         =   "Move Down"
                     Height          =   255
                     Index           =   5
                     Left            =   120
                     TabIndex        =   17
                     TabStop         =   0   'False
                     ToolTipText     =   "Move a property down the list"
                     Top             =   1140
                     Width           =   1200
                  End
               End
            End
         End
      End
      Begin Threed.SSPanel sspBottom 
         Height          =   2775
         Left            =   60
         TabIndex        =   23
         Top             =   2910
         Width           =   9645
         _Version        =   65536
         _ExtentX        =   17013
         _ExtentY        =   4895
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
         Begin Threed.SSFrame ssfBottom 
            Height          =   2565
            Left            =   90
            TabIndex        =   24
            Top             =   90
            Width           =   8505
            _Version        =   65536
            _ExtentX        =   15002
            _ExtentY        =   4524
            _StockProps     =   14
            Caption         =   "Selection and Ordering of Techniques for Each Property:"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin Threed.SSFrame ssfBottomLeft 
               Height          =   1935
               Left            =   90
               TabIndex        =   25
               Top             =   270
               Width           =   1425
               _Version        =   65536
               _ExtentX        =   2514
               _ExtentY        =   3413
               _StockProps     =   14
               Caption         =   "Properties:"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Begin VB.ListBox lstBottom_Props 
                  Height          =   1230
                  Left            =   90
                  TabIndex        =   26
                  Top             =   270
                  Width           =   1155
               End
            End
            Begin Threed.SSFrame ssfBottomLeftCmds 
               Height          =   1935
               Left            =   1500
               TabIndex        =   27
               Top             =   270
               Width           =   1455
               _Version        =   65536
               _ExtentX        =   2566
               _ExtentY        =   3413
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
            End
            Begin Threed.SSFrame ssfBottomRight 
               Height          =   2145
               Left            =   2940
               TabIndex        =   28
               Top             =   270
               Width           =   5055
               _Version        =   65536
               _ExtentX        =   8916
               _ExtentY        =   3784
               _StockProps     =   14
               Caption         =   "Techniques for 'Property 1':"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Begin Threed.SSFrame ssfBottomRight1 
                  Height          =   1695
                  Left            =   120
                  TabIndex        =   29
                  Top             =   270
                  Width           =   1245
                  _Version        =   65536
                  _ExtentX        =   2196
                  _ExtentY        =   2990
                  _StockProps     =   14
                  Caption         =   "Selected:"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Begin VB.ListBox lstTechniques 
                     Height          =   1230
                     Left            =   90
                     MultiSelect     =   2  'Extended
                     TabIndex        =   30
                     Top             =   270
                     Width           =   1065
                  End
               End
               Begin Threed.SSFrame ssfBottomRight2 
                  Height          =   1695
                  Left            =   2790
                  TabIndex        =   31
                  Top             =   270
                  Width           =   1905
                  _Version        =   65536
                  _ExtentX        =   3360
                  _ExtentY        =   2990
                  _StockProps     =   14
                  Caption         =   "Not Selected:"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Begin VB.ListBox lstTechniquesAll 
                     Height          =   1230
                     Left            =   90
                     MultiSelect     =   2  'Extended
                     TabIndex        =   32
                     Top             =   270
                     Width           =   1695
                  End
               End
               Begin Threed.SSFrame ssfBottomRight1Cmds 
                  Height          =   1695
                  Left            =   1350
                  TabIndex        =   33
                  Top             =   270
                  Width           =   1455
                  _Version        =   65536
                  _ExtentX        =   2566
                  _ExtentY        =   2990
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
                  Begin VB.CommandButton cmdBottomRight1Cmds 
                     Caption         =   "Move Down"
                     Height          =   255
                     Index           =   5
                     Left            =   120
                     TabIndex        =   39
                     TabStop         =   0   'False
                     ToolTipText     =   "Move a property down the list"
                     Top             =   1140
                     Width           =   1200
                  End
                  Begin VB.CommandButton cmdBottomRight1Cmds 
                     Caption         =   "Move Up"
                     Height          =   255
                     Index           =   4
                     Left            =   120
                     TabIndex        =   38
                     TabStop         =   0   'False
                     ToolTipText     =   "Move a property up the list"
                     Top             =   870
                     Width           =   1200
                  End
                  Begin VB.CommandButton cmdBottomRight1Cmds 
                     Caption         =   ">"
                     Height          =   255
                     Index           =   2
                     Left            =   120
                     TabIndex        =   37
                     TabStop         =   0   'False
                     ToolTipText     =   "Deselect technique(s) from property"
                     Top             =   600
                     Width           =   585
                  End
                  Begin VB.CommandButton cmdBottomRight1Cmds 
                     Caption         =   ">>"
                     Height          =   255
                     Index           =   3
                     Left            =   735
                     TabIndex        =   36
                     TabStop         =   0   'False
                     ToolTipText     =   "Deselect all techniques from property"
                     Top             =   600
                     Width           =   585
                  End
                  Begin VB.CommandButton cmdBottomRight1Cmds 
                     Caption         =   "<"
                     Height          =   255
                     Index           =   0
                     Left            =   120
                     TabIndex        =   35
                     TabStop         =   0   'False
                     ToolTipText     =   "Select technique(s) into property"
                     Top             =   330
                     Width           =   585
                  End
                  Begin VB.CommandButton cmdBottomRight1Cmds 
                     Caption         =   "<<"
                     Height          =   255
                     Index           =   1
                     Left            =   735
                     TabIndex        =   34
                     TabStop         =   0   'False
                     ToolTipText     =   "Select all techniques into property"
                     Top             =   330
                     Width           =   585
                  End
               End
            End
         End
      End
   End
End
Attribute VB_Name = "frmCustomProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Save_TempCopy_NowProj As Project_Type

Dim USER_HIT_CANCEL As Boolean
Dim USER_HIT_OK As Boolean
Public HALT_ALL_CONTROLS As Boolean





Const frmCustomProperties_decl_end = True


Function frmCustomProperties_Go( _
    out_HitCancel As Boolean) _
    As Boolean
On Error GoTo err_ThisFunc
Dim Frm As Form
Set Frm = frmCustomProperties
  Save_TempCopy_NowProj = NowProj
  Frm.Show 1
  out_HitCancel = IIf(USER_HIT_CANCEL = True, True, False)
  If (out_HitCancel = True) Then
    NowProj = Save_TempCopy_NowProj
  Else
    'Call PrefEnvironment_SaveToINI
  End If
exit_normally_ThisFunc:
  frmCustomProperties_Go = True
  Exit Function
exit_err_ThisFunc:
  frmCustomProperties_Go = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("frmCustomProperties_Go")
  Resume exit_err_ThisFunc
End Function


Function frmCustomProperties_Resize( _
    ) _
    As Boolean
On Error GoTo err_ThisFunc
Const USE_MARGIN_sspSep = 50
Const USE_MARGIN_CONCENTRIC_FRAMES_X = 90
Const USE_MARGIN_CONCENTRIC_FRAMES_Y = 270
Dim XX As Double
Dim Left_of_This As Double
Dim Top_of_This As Double
Dim Width_of_This As Double
Dim Height_of_This As Double
Dim Left_Major_Separation As Double
Dim Left_Minor_Separation As Double
Dim Frm As Form
Set Frm = Me
  '
  '////////// START OF MAIN RESIZING CODE. ///////////////////////////////////////
  '
  '
  ' RESIZE sspMainButtons AND ALL CONTAINED CONTROLS.
  '
  XX = Frm.ScaleWidth - sspMainButtons.Width
  Left_of_This = IIf(XX > 10, XX, 10)
  XX = Frm.ScaleHeight - sspMainButtons.Height
  Top_of_This = IIf(XX > 10, XX, 10)
  sspMainButtons.Move _
      Left_of_This, _
      Top_of_This, _
      sspMainButtons.Width, _
      sspMainButtons.Height
  '
  ' RESIZE sspAll AND ALL CONTAINED CONTROLS.
  '
  Width_of_This = Frm.ScaleWidth
  XX = Frm.ScaleHeight - sspMainButtons.Height
  Height_of_This = IIf(XX > 10, XX, 10)
  sspAll.Move _
      0, _
      0, _
      Width_of_This, _
      Height_of_This
    '
    ' MISCELLANEOUS.
    '
    Left_Major_Separation = CDbl(sspAll.Width) * 1# / 3#
    '
    ' RESIZE sspTop AND ALL CONTAINED CONTROLS.
    '
    Width_of_This = sspAll.Width
    Height_of_This = CDbl(sspAll.Height) / 2#
    sspTop.Move _
        0, _
        0, _
        Width_of_This, _
        Height_of_This
      '
      ' RESIZE ssfTop AND ALL CONTAINED CONTROLS.
      '
      Width_of_This = sspTop.Width
      Height_of_This = sspTop.Height
      ssfTop.Move _
          0, _
          0, _
          Width_of_This, _
          Height_of_This
        '
        ' RESIZE ssfTopLeft AND ALL CONTAINED CONTROLS.
        '
        XX = Left_Major_Separation - ssfTopLeftCmds.Width - USE_MARGIN_CONCENTRIC_FRAMES_X
        Width_of_This = IIf(XX > 10, XX, 10)
        XX = ssfTop.Height - USE_MARGIN_CONCENTRIC_FRAMES_Y * 2
        Height_of_This = IIf(XX > 10, XX, 10)
        ssfTopLeft.Move _
            USE_MARGIN_CONCENTRIC_FRAMES_X, _
            USE_MARGIN_CONCENTRIC_FRAMES_Y, _
            Width_of_This, _
            Height_of_This
          '
          ' RESIZE lstTop_PropSheets AND ALL CONTAINED CONTROLS.
          '
          XX = ssfTopLeft.Width - USE_MARGIN_CONCENTRIC_FRAMES_X * 2#
          Width_of_This = IIf(XX > 10, XX, 10)
          XX = ssfTopLeft.Height - USE_MARGIN_CONCENTRIC_FRAMES_Y * 2#
          Height_of_This = IIf(XX > 10, XX, 10)
          lstTop_PropSheets.Move _
              USE_MARGIN_CONCENTRIC_FRAMES_X, _
              USE_MARGIN_CONCENTRIC_FRAMES_Y, _
              Width_of_This, _
              Height_of_This
        '
        ' RESIZE ssfTopLeftCmds AND ALL CONTAINED CONTROLS.
        '
        Width_of_This = ssfTopLeftCmds.Width
        Height_of_This = ssfTopLeft.Height
        ssfTopLeftCmds.Move _
            ssfTopLeft.Left + ssfTopLeft.Width, _
            USE_MARGIN_CONCENTRIC_FRAMES_Y, _
            Width_of_This, _
            Height_of_This
        '
        ' RESIZE ssfTopRight AND ALL CONTAINED CONTROLS.
        '
        XX = ssfTop.Width - Left_Major_Separation - USE_MARGIN_CONCENTRIC_FRAMES_X
        Width_of_This = IIf(XX > 10, XX, 10)
        XX = ssfTop.Height - USE_MARGIN_CONCENTRIC_FRAMES_Y * 2
        Height_of_This = IIf(XX > 10, XX, 10)
        ssfTopRight.Move _
            Left_Major_Separation, _
            USE_MARGIN_CONCENTRIC_FRAMES_Y, _
            Width_of_This, _
            Height_of_This
          '
          ' MISCELLANEOUS.
          '
          Left_Minor_Separation = _
              USE_MARGIN_CONCENTRIC_FRAMES_X + _
              (CDbl(ssfTopRight.Width) - 2# * USE_MARGIN_CONCENTRIC_FRAMES_X - _
              CDbl(ssfTopRight1Cmds.Width)) / 2#
          '
          ' RESIZE ssfTopRight1 AND ALL CONTAINED CONTROLS.
          '
          XX = Left_Minor_Separation - USE_MARGIN_CONCENTRIC_FRAMES_X
          Width_of_This = IIf(XX > 10, XX, 10)
          XX = ssfTopRight.Height - USE_MARGIN_CONCENTRIC_FRAMES_Y * 2
          Height_of_This = IIf(XX > 10, XX, 10)
          ssfTopRight1.Move _
              USE_MARGIN_CONCENTRIC_FRAMES_X, _
              USE_MARGIN_CONCENTRIC_FRAMES_Y, _
              Width_of_This, _
              Height_of_This
            '
            ' RESIZE lstTop_Props AND ALL CONTAINED CONTROLS.
            '
            XX = ssfTopRight1.Width - USE_MARGIN_CONCENTRIC_FRAMES_X * 2#
            Width_of_This = IIf(XX > 10, XX, 10)
            XX = ssfTopRight1.Height - USE_MARGIN_CONCENTRIC_FRAMES_Y * 2#
            Height_of_This = IIf(XX > 10, XX, 10)
            lstTop_Props.Move _
                USE_MARGIN_CONCENTRIC_FRAMES_X, _
                USE_MARGIN_CONCENTRIC_FRAMES_Y, _
                Width_of_This, _
                Height_of_This
          '
          ' RESIZE ssfTopRight1Cmds AND ALL CONTAINED CONTROLS.
          '
          Width_of_This = ssfTopRight1Cmds.Width
          Height_of_This = ssfTopRight1.Height
          ssfTopRight1Cmds.Move _
              ssfTopRight1.Left + ssfTopRight1.Width, _
              USE_MARGIN_CONCENTRIC_FRAMES_Y, _
              Width_of_This, _
              Height_of_This
          '
          ' RESIZE ssfTopRight2 AND ALL CONTAINED CONTROLS.
          '
          XX = ssfTopRight.Width - USE_MARGIN_CONCENTRIC_FRAMES_X - _
              (ssfTopRight1Cmds.Left + ssfTopRight1Cmds.Width)
          Width_of_This = IIf(XX > 10, XX, 10)
          XX = ssfTopRight.Height - USE_MARGIN_CONCENTRIC_FRAMES_Y * 2
          Height_of_This = IIf(XX > 10, XX, 10)
          ssfTopRight2.Move _
              ssfTopRight1Cmds.Left + ssfTopRight1Cmds.Width, _
              USE_MARGIN_CONCENTRIC_FRAMES_Y, _
              Width_of_This, _
              Height_of_This
            '
            ' RESIZE lstTop_PropsAll AND ALL CONTAINED CONTROLS.
            '
            XX = ssfTopRight2.Width - USE_MARGIN_CONCENTRIC_FRAMES_X * 2#
            Width_of_This = IIf(XX > 10, XX, 10)
            XX = ssfTopRight2.Height - USE_MARGIN_CONCENTRIC_FRAMES_Y * 2#
            Height_of_This = IIf(XX > 10, XX, 10)
            lstTop_PropsAll.Move _
                USE_MARGIN_CONCENTRIC_FRAMES_X, _
                USE_MARGIN_CONCENTRIC_FRAMES_Y, _
                Width_of_This, _
                Height_of_This
    '
    ' RESIZE sspBottom AND ALL CONTAINED CONTROLS.
    '
    Width_of_This = sspAll.Width
    Height_of_This = CDbl(sspAll.Height) / 2#
    sspBottom.Move _
        0, _
        Height_of_This, _
        Width_of_This, _
        Height_of_This
      '
      ' RESIZE ssfBottom AND ALL CONTAINED CONTROLS.
      '
      Width_of_This = sspBottom.Width
      Height_of_This = sspBottom.Height
      ssfBottom.Move _
          0, _
          0, _
          Width_of_This, _
          Height_of_This
        '
        ' RESIZE ssfBottomLeft AND ALL CONTAINED CONTROLS.
        '
        XX = Left_Major_Separation - ssfBottomLeftCmds.Width - USE_MARGIN_CONCENTRIC_FRAMES_X
        Width_of_This = IIf(XX > 10, XX, 10)
        XX = ssfBottom.Height - USE_MARGIN_CONCENTRIC_FRAMES_Y * 2
        Height_of_This = IIf(XX > 10, XX, 10)
        ssfBottomLeft.Move _
            USE_MARGIN_CONCENTRIC_FRAMES_X, _
            USE_MARGIN_CONCENTRIC_FRAMES_Y, _
            Width_of_This, _
            Height_of_This
          '
          ' RESIZE lstBottom_Props AND ALL CONTAINED CONTROLS.
          '
          XX = ssfBottomLeft.Width - USE_MARGIN_CONCENTRIC_FRAMES_X * 2#
          Width_of_This = IIf(XX > 10, XX, 10)
          XX = ssfBottomLeft.Height - USE_MARGIN_CONCENTRIC_FRAMES_Y * 2#
          Height_of_This = IIf(XX > 10, XX, 10)
          lstBottom_Props.Move _
              USE_MARGIN_CONCENTRIC_FRAMES_X, _
              USE_MARGIN_CONCENTRIC_FRAMES_Y, _
              Width_of_This, _
              Height_of_This
        '
        ' RESIZE ssfBottomLeftCmds AND ALL CONTAINED CONTROLS.
        '
        Width_of_This = ssfBottomLeftCmds.Width
        Height_of_This = ssfBottomLeft.Height
        ssfBottomLeftCmds.Move _
            ssfBottomLeft.Left + ssfBottomLeft.Width, _
            USE_MARGIN_CONCENTRIC_FRAMES_Y, _
            Width_of_This, _
            Height_of_This
        '
        ' RESIZE ssfBottomRight AND ALL CONTAINED CONTROLS.
        '
        XX = ssfBottom.Width - Left_Major_Separation - USE_MARGIN_CONCENTRIC_FRAMES_X
        Width_of_This = IIf(XX > 10, XX, 10)
        XX = ssfBottom.Height - USE_MARGIN_CONCENTRIC_FRAMES_Y * 2
        Height_of_This = IIf(XX > 10, XX, 10)
        ssfBottomRight.Move _
            Left_Major_Separation, _
            USE_MARGIN_CONCENTRIC_FRAMES_Y, _
            Width_of_This, _
            Height_of_This
          '
          ' MISCELLANEOUS.
          '
          Left_Minor_Separation = _
              USE_MARGIN_CONCENTRIC_FRAMES_X + _
              (CDbl(ssfBottomRight.Width) - 2# * USE_MARGIN_CONCENTRIC_FRAMES_X - _
              CDbl(ssfBottomRight1Cmds.Width)) / 2#
          '
          ' RESIZE ssfBottomRight1 AND ALL CONTAINED CONTROLS.
          '
          XX = Left_Minor_Separation - USE_MARGIN_CONCENTRIC_FRAMES_X
          Width_of_This = IIf(XX > 10, XX, 10)
          XX = ssfBottomRight.Height - USE_MARGIN_CONCENTRIC_FRAMES_Y * 2
          Height_of_This = IIf(XX > 10, XX, 10)
          ssfBottomRight1.Move _
              USE_MARGIN_CONCENTRIC_FRAMES_X, _
              USE_MARGIN_CONCENTRIC_FRAMES_Y, _
              Width_of_This, _
              Height_of_This
            '
            ' RESIZE lstTechniques AND ALL CONTAINED CONTROLS.
            '
            XX = ssfBottomRight1.Width - USE_MARGIN_CONCENTRIC_FRAMES_X * 2#
            Width_of_This = IIf(XX > 10, XX, 10)
            XX = ssfBottomRight1.Height - USE_MARGIN_CONCENTRIC_FRAMES_Y * 2#
            Height_of_This = IIf(XX > 10, XX, 10)
            lstTechniques.Move _
                USE_MARGIN_CONCENTRIC_FRAMES_X, _
                USE_MARGIN_CONCENTRIC_FRAMES_Y, _
                Width_of_This, _
                Height_of_This
          '
          ' RESIZE ssfBottomRight1Cmds AND ALL CONTAINED CONTROLS.
          '
          Width_of_This = ssfBottomRight1Cmds.Width
          Height_of_This = ssfBottomRight1.Height
          ssfBottomRight1Cmds.Move _
              ssfBottomRight1.Left + ssfBottomRight1.Width, _
              USE_MARGIN_CONCENTRIC_FRAMES_Y, _
              Width_of_This, _
              Height_of_This
          '
          ' RESIZE ssfBottomRight2 AND ALL CONTAINED CONTROLS.
          '
          XX = ssfBottomRight.Width - USE_MARGIN_CONCENTRIC_FRAMES_X - _
              (ssfBottomRight1Cmds.Left + ssfBottomRight1Cmds.Width)
          Width_of_This = IIf(XX > 10, XX, 10)
          XX = ssfBottomRight.Height - USE_MARGIN_CONCENTRIC_FRAMES_Y * 2
          Height_of_This = IIf(XX > 10, XX, 10)
          ssfBottomRight2.Move _
              ssfBottomRight1Cmds.Left + ssfBottomRight1Cmds.Width, _
              USE_MARGIN_CONCENTRIC_FRAMES_Y, _
              Width_of_This, _
              Height_of_This
            '
            ' RESIZE lstTechniquesAll AND ALL CONTAINED CONTROLS.
            '
            XX = ssfBottomRight2.Width - USE_MARGIN_CONCENTRIC_FRAMES_X * 2#
            Width_of_This = IIf(XX > 10, XX, 10)
            XX = ssfBottomRight2.Height - USE_MARGIN_CONCENTRIC_FRAMES_Y * 2#
            Height_of_This = IIf(XX > 10, XX, 10)
            lstTechniquesAll.Move _
                USE_MARGIN_CONCENTRIC_FRAMES_X, _
                USE_MARGIN_CONCENTRIC_FRAMES_Y, _
                Width_of_This, _
                Height_of_This
    
    
    
''''    '
''''    ' RESIZE sspTop AND ALL CONTAINED CONTROLS.
''''    '
''''    XX = sspSep.Top - USE_MARGIN_sspSep
''''    XX = IIf(XX > 10, XX, 10)
''''    sspTop.Move _
''''        0, _
''''        0, _
''''        sspAll.Width, _
''''        XX
''''      '
''''      ' SOME CONTAINED CONTROLS ...
''''      '
''''Dim Width_of_Each_List As Double
''''Dim Height_of_Each_List As Double
''''Dim Width_of_This As Double
''''Dim Height_of_This As Double
''''      XX = CDbl(sspTop.Width - ssfButtons.Width - 60 - 60) / 2#
''''      Width_of_Each_List = IIf(XX > 100#, XX, 100#)
''''      XX = sspTop.Height - 60 - 60
''''      Height_of_Each_List = IIf(XX > 100#, XX, 100#)
''''      '
''''      ' RESIZE ssfMasterList AND ALL CONTAINED CONTROLS.
''''      '
''''      ssfMasterList.Move _
''''          60, _
''''          60, _
''''          Width_of_Each_List, _
''''          Height_of_Each_List
''''        '
''''        ' RESIZE dblstMaster AND ALL CONTAINED CONTROLS.
''''        '
''''        XX = ssfMasterList.Width - 120 - 120
''''        Width_of_This = IIf(XX > 100#, XX, 100#)
''''        XX = ssfMasterList.Height - 300 - 120
''''        Height_of_This = IIf(XX > 100#, XX, 100#)
''''        dblstMaster.Move _
''''            120, _
''''            300, _
''''            Width_of_This, _
''''            Height_of_This
''''      '
''''      ' RESIZE ssfButtons AND ALL CONTAINED CONTROLS.
''''      '
''''      ssfButtons.Move _
''''          ssfMasterList.Left + ssfMasterList.Width, _
''''          60, _
''''          ssfButtons.Width, _
''''          ssfButtons.Height
''''      '
''''      ' RESIZE ssfUserList AND ALL CONTAINED CONTROLS.
''''      '
''''      ssfUserList.Move _
''''          ssfButtons.Left + ssfButtons.Width, _
''''          60, _
''''          Width_of_Each_List, _
''''          Height_of_Each_List
''''        '
''''        ' RESIZE lstUser AND ALL CONTAINED CONTROLS.
''''        '
''''        XX = ssfUserList.Width - 120 - 120
''''        Width_of_This = IIf(XX > 100#, XX, 100#)
''''        XX = ssfUserList.Height - 300 - 120
''''        Height_of_This = IIf(XX > 100#, XX, 100#)
''''        lstUser.Move _
''''            120, _
''''            300, _
''''            Width_of_This, _
''''            Height_of_This
''''    '
''''    ' RESIZE sspBottom AND ALL CONTAINED CONTROLS.
''''    '
''''    XX = sspAll.Height - (sspSep.Top + sspSep.Height + USE_MARGIN_sspSep)
''''    XX = IIf(XX > 10, XX, 10)
''''    sspBottom.Move _
''''        0, _
''''        sspSep.Top + sspSep.Height + USE_MARGIN_sspSep, _
''''        sspAll.Width, _
''''        XX
''''      '
''''      ' SOME CONTAINED CONTROLS ...
''''      '
''''Dim Height_of_Each_Property_Frame As Double
''''      XX = sspBottom.Height - 60 - 60
''''      Height_of_Each_Property_Frame = IIf(XX > 10, XX, 10)
''''      '
''''      ' RESIZE ssfPropSheets AND ALL CONTAINED CONTROLS.
''''      '
''''      ssfPropSheets.Move _
''''          60, _
''''          60, _
''''          ssfPropSheets.Width, _
''''          Height_of_Each_Property_Frame
''''        '
''''        ' RESIZE lstPropSheets AND ALL CONTAINED CONTROLS.
''''        '
''''        XX = ssfPropSheets.Width - 120 - 120
''''        Width_of_This = IIf(XX > 100#, XX, 100#)
''''        XX = ssfPropSheets.Height - 300 - 120
''''        Height_of_This = IIf(XX > 100#, XX, 100#)
''''        lstPropSheets.Move _
''''            120, _
''''            300, _
''''            Width_of_This, _
''''            Height_of_This
''''      '
''''      ' RESIZE ssfMain AND ALL CONTAINED CONTROLS.
''''      '
''''      XX = sspBottom.Width - 60 - (ssfPropSheets.Left + ssfPropSheets.Width)
''''      XX = IIf(XX > 10, XX, 10)
''''      ssfMain.Move _
''''          ssfPropSheets.Left + ssfPropSheets.Width, _
''''          60, _
''''          XX, _
''''          Height_of_Each_Property_Frame
''''        '
''''        ' RESIZE lvMain AND ALL CONTAINED CONTROLS.
''''        '
''''        XX = ssfMain.Width - 120 - 120
''''        Width_of_This = IIf(XX > 100#, XX, 100#)
''''        XX = ssfMain.Height - 300 - 120
''''        Height_of_This = IIf(XX > 100#, XX, 100#)
''''        lvMain.Move _
''''            120, _
''''            300, _
''''            Width_of_This, _
''''            Height_of_This
''''        '
''''        ' RESIZE txtChemNote AND ALL CONTAINED CONTROLS.
''''        '
''''        txtChemNote.Move _
''''            lvMain.Left, _
''''            lvMain.Top, _
''''            lvMain.Width, _
''''            lvMain.Height
''''        '
''''        ' RESIZE sspBasic AND ALL CONTAINED CONTROLS.
''''        '
''''        sspBasic.Move _
''''            lvMain.Left, _
''''            lvMain.Top, _
''''            lvMain.Width, _
''''            lvMain.Height
  
  
  
  
  '
  '////////// END OF MAIN RESIZING CODE. ///////////////////////////////////////
  '
exit_normally_ThisFunc:
  frmCustomProperties_Resize = True
  Exit Function
exit_err_ThisFunc:
  frmCustomProperties_Resize = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("frmCustomProperties_Resize")
  GoTo exit_err_ThisFunc
End Function


Private Sub cmdBottomRight1Cmds_Click(Index As Integer)
On Error GoTo err_ThisFunc
Dim lstBottom_Props_Index As Long
Dim List_TechCodes() As Long
Dim out_List_PropCodes() As Long
Dim out_List_idxPropSheet_First_Occurrences() As Integer
Dim out_List_idxPropOrd_First_Occurrences() As Integer
Dim UB_Selected As Integer
Dim UB_Available As Integer
Dim i As Integer
Dim j As Integer
Dim This_TechCode As Long
Dim Current_Technique_Code() As Long
Dim out_idx_Elem As Integer
Dim out_idx_Elem_ThisProp As Integer
Dim Temp_Storage As Long
  If (lstBottom_Props.ListIndex < 0) Then
    Beep
    Exit Sub
  End If
  '
  ' POPULATE LIST OF ALL POSSIBLE TECHNIQUES.
  '
  lstBottom_Props_Index = lstBottom_Props.ItemData(lstBottom_Props.ListIndex)
  Call Get_Complete_List_of_TechCodes( _
      lstBottom_Props_Index, _
      List_TechCodes)
  UB_Available = UBound(List_TechCodes)
  Call PropertyOrder_Get_List_of_Unique_Property_Codes( _
      out_List_PropCodes, _
      out_List_idxPropSheet_First_Occurrences, _
      out_List_idxPropOrd_First_Occurrences)
  If (False = sc_ElemFind( _
      out_List_PropCodes, _
      lstBottom_Props_Index, _
      out_idx_Elem)) Then
    GoTo exit_err_ThisFunc
  End If
  out_idx_Elem_ThisProp = out_idx_Elem
  Current_Technique_Code = _
      NowProj.UserHierarchy. _
      PropertySheetOrder( _
          out_List_idxPropSheet_First_Occurrences(out_idx_Elem_ThisProp)). _
      PropertyOrder( _
          out_List_idxPropOrd_First_Occurrences(out_idx_Elem_ThisProp)). _
      Technique_Code
  UB_Selected = UBound(Current_Technique_Code)
  '
  ' DETERMINE SELECTION TYPE (IF ANY) IN lstTechniques.
  '
Dim Is_MultiSelection As Boolean
Dim Is_SingleSelection As Boolean
Dim Is_FirstSelected As Boolean
Dim Is_LastSelected As Boolean
Dim idx_FirstSelection As Integer
  Is_MultiSelection = False
  Is_SingleSelection = False
  Is_FirstSelected = False
  Is_LastSelected = False
  For i = 0 To lstTechniques.ListCount - 1
    If (lstTechniques.Selected(i) = True) Then
      If (i = 0) Then Is_FirstSelected = True
      If (i = lstTechniques.ListCount - 1) Then Is_LastSelected = True
      If (Is_SingleSelection = False) And (Is_MultiSelection = False) Then
        Is_SingleSelection = True
        idx_FirstSelection = i + 1
      Else
        Is_SingleSelection = False
        Is_MultiSelection = True
      End If
    End If
  Next i
  Select Case Index
    '
    '////////////////////////////////////////////////////////////////////
    Case 0:       'SELECT TECHNIQUE(S).
      '
      ' ADD TECHNIQUE(S) FROM TEMPORARY VARIABLE.
      '
      For i = 0 To lstTechniquesAll.ListCount - 1
        If (lstTechniquesAll.Selected(i) = True) Then
          This_TechCode = lstTechniquesAll.ItemData(i)
          UB_Selected = UBound(Current_Technique_Code)
          UB_Selected = UB_Selected + 1
          If (UB_Selected = 1) Then
            ReDim Current_Technique_Code(1 To 1)
          Else
            ReDim Preserve Current_Technique_Code(1 To UB_Selected)
          End If
          Current_Technique_Code(UB_Selected) = This_TechCode
        End If
      Next i
      '
      ' COPY TEMPORARY VARIABLE TO ALL OCCURRENCES OF THE PROPERTY.
      '
      Call PropertyOrder_Update_Technique_Code( _
          Current_Technique_Code(), _
          lstBottom_Props_Index)
      '
      ' REFRESH WINDOW.
      '
      Call frmCustomProperties_Refresh
    '
    '////////////////////////////////////////////////////////////////////
    Case 1:       'SELECT ALL TECHNIQUES.
      If (lstTechniquesAll.ListCount = 0) Then
        Beep
        GoTo exit_err_ThisFunc
      End If
      For i = 0 To lstTechniquesAll.ListCount - 1
        lstTechniquesAll.Selected(i) = True
      Next i
      Call cmdBottomRight1Cmds_Click(0)
    '
    '////////////////////////////////////////////////////////////////////
    Case 2:       'DESELECT TECHNIQUE(S).
      '
      ' REMOVE TECHNIQUE(S) FROM TEMPORARY VARIABLE.
      '
      For i = 0 To lstTechniques.ListCount - 1
        If (lstTechniques.Selected(i) = True) Then
          This_TechCode = lstTechniques.ItemData(i)
          If (True = sc_ElemFind( _
              Current_Technique_Code(), _
              This_TechCode, _
              out_idx_Elem)) Then
            UB_Selected = UBound(Current_Technique_Code)
            For j = out_idx_Elem To UB_Selected - 1
              Current_Technique_Code(j) = Current_Technique_Code(j + 1)
            Next j
            UB_Selected = UB_Selected - 1
            If (UB_Selected = 0) Then
              ReDim Current_Technique_Code(0 To 0)
            Else
              ReDim Preserve Current_Technique_Code(1 To UB_Selected)
            End If
          End If
        End If
      Next i
      '
      ' COPY TEMPORARY VARIABLE TO ALL OCCURRENCES OF THE PROPERTY.
      '
      Call PropertyOrder_Update_Technique_Code( _
          Current_Technique_Code(), _
          lstBottom_Props_Index)
      '
      ' REFRESH WINDOW.
      '
      Call frmCustomProperties_Refresh
    '
    '////////////////////////////////////////////////////////////////////
    Case 3:       'DESELECT ALL TECHNIQUES.
      If (lstTechniques.ListCount = 0) Then
        Beep
        GoTo exit_err_ThisFunc
      End If
      For i = 0 To lstTechniques.ListCount - 1
        lstTechniques.Selected(i) = True
      Next i
      Call cmdBottomRight1Cmds_Click(2)
    '
    '////////////////////////////////////////////////////////////////////
    Case 4:       'MOVE UP.
      If (Is_MultiSelection = True) Or (Is_SingleSelection = False) Then
        Call Show_Error("To use this command, you must highlight " & _
            "only a single technique.")
        GoTo exit_err_ThisFunc
      End If
      If (Is_FirstSelected) Then
        Beep
        GoTo exit_err_ThisFunc
      End If
      '
      ' SWAP THE TWO VALUES.
      '
      Temp_Storage = Current_Technique_Code(idx_FirstSelection)
      Current_Technique_Code(idx_FirstSelection) = Current_Technique_Code(idx_FirstSelection - 1)
      Current_Technique_Code(idx_FirstSelection - 1) = Temp_Storage
      '
      ' COPY TEMPORARY VARIABLE TO ALL OCCURRENCES OF THE PROPERTY.
      '
      Call PropertyOrder_Update_Technique_Code( _
          Current_Technique_Code(), _
          lstBottom_Props_Index)
      '
      ' REFRESH WINDOW.
      '
      Call frmCustomProperties_Refresh
    '
    '////////////////////////////////////////////////////////////////////
    Case 5:       'MOVE DOWN.
      If (Is_MultiSelection = True) Or (Is_SingleSelection = False) Then
        Call Show_Error("To use this command, you must highlight " & _
            "only a single technique.")
        GoTo exit_err_ThisFunc
      End If
      If (Is_LastSelected) Then
        Beep
        GoTo exit_err_ThisFunc
      End If
      '
      ' SWAP THE TWO VALUES.
      '
      Temp_Storage = Current_Technique_Code(idx_FirstSelection)
      Current_Technique_Code(idx_FirstSelection) = Current_Technique_Code(idx_FirstSelection + 1)
      Current_Technique_Code(idx_FirstSelection + 1) = Temp_Storage
      '
      ' COPY TEMPORARY VARIABLE TO ALL OCCURRENCES OF THE PROPERTY.
      '
      Call PropertyOrder_Update_Technique_Code( _
          Current_Technique_Code(), _
          lstBottom_Props_Index)
      '
      ' REFRESH WINDOW.
      '
      Call frmCustomProperties_Refresh
  End Select
exit_normally_ThisFunc:
  ';xxxxx = True
  Exit Sub
exit_err_ThisFunc:
  'xxxxx = False
  Exit Sub
err_ThisFunc:
  Call Show_Trapped_Error("cmdBottomRight1Cmds_Click")
  Resume exit_err_ThisFunc
End Sub


Private Sub cmdCancelOK_Click(Index As Integer)
  Select Case Index
    Case 0:       'CANCEL.
      USER_HIT_CANCEL = True
      USER_HIT_OK = False
      Unload Me
      Exit Sub
    Case 1:       'OK.
      If (vbNo = MsgBox("Are you sure you want to accept the " & _
          "new property hierarchy?  This will result in the " & _
          "deletion of all user input data (if any).  This step cannot be undone.", _
          vbQuestion + vbYesNo, _
          "Accept new property hierarchy: Are you sure?")) Then
        Exit Sub
      End If
      USER_HIT_CANCEL = False
      USER_HIT_OK = True
      Unload Me
      Exit Sub
  End Select
End Sub


Private Sub cmdRestoreDefaults_Click()
  If (vbNo = MsgBox("Are you sure you want to restore the " & _
      "default property hierarchy?  This step cannot be undone.", _
      vbQuestion + vbYesNo, _
      "Restore default property hierarchy: Are you sure?")) Then
    Exit Sub
  End If
  '
  ' SET DEFAULT HIERARCHY.
  '
  Call Project_UserHierarchy_SetDefaults(NowProj)
  '
  ' REFRESH WINDOW.
  '
  Call frmCustomProperties_Refresh
End Sub


Private Sub cmdTopLeftCmds_Click(Index As Integer)
On Error GoTo err_ThisFunc
Dim idx_PropSheet As Integer
Dim Name_PropSheet As String
Dim UB As Integer
Dim i As Integer
Dim New_Name As String
Dim is_aborted As Boolean
Dim Temp_PropSheetOrder As PropertySheetOrder_Type
  idx_PropSheet = lstTop_PropSheets.ListIndex + 1
  If (idx_PropSheet = 0) Then
    Name_PropSheet = ""
  Else
    Name_PropSheet = lstTop_PropSheets.List(lstTop_PropSheets.ListIndex)
  End If
  Select Case Index
    '
    '////////////////////////////////////////////////////////////////////
    Case 0:           'ADD PROPERTY SHEET.
      If (UBound(NowProj.UserHierarchy.PropertySheetOrder) >= MAX_PROPERTYSHEETS) Then
        Call Show_Error("The maximum number of property sheets " & _
            "has been reached (" & Trim$(Str$(MAX_PROPERTYSHEETS)) & _
            ".  You cannot add any more property sheets.")
        GoTo exit_err_ThisFunc
      End If
      Name_PropSheet = PropertySheetOrder_GetNewNameDefault()
      Do While (1 = 1)
        New_Name = frmNewName.frmNewName_GetName( _
            "Rename property sheet", _
            "Enter new name for the selected property sheet.  Each " & _
            "property sheet must have a unique name.", _
            Name_PropSheet, _
            is_aborted)
        If (is_aborted) Then GoTo exit_normally_ThisFunc
        If ((Trim$(UCase$(New_Name)) = Trim$(UCase$(PROPERTYSHEETNAME_BASIC_CHEMICAL_INFO))) Or _
            (Trim$(UCase$(New_Name)) = Trim$(UCase$(PROPERTYSHEETNAME_CHEMICAL_NOTE)))) Then
          Call Show_Error("You cannot add a property sheet " & _
              "named `" & PROPERTYSHEETNAME_BASIC_CHEMICAL_INFO & _
              "` or `" & PROPERTYSHEETNAME_CHEMICAL_NOTE & "`.  " & _
              "Please enter a different name or hit Cancel.")
        Else
          If (PropertySheetOrder_IsKeyExist(New_Name) = False) Then
            Exit Do
          Else
            Call Show_Error("The name `" & New_Name & _
                "` already exists.  Please enter a different name or hit Cancel.")
          End If
        End If
      Loop
      '
      ' ADD IT.
      '
      UB = UBound(NowProj.UserHierarchy.PropertySheetOrder)
      UB = UB + 1
      If (UB = 1) Then
        ' THEORETICALLY IMPOSSIBLE TO GET HERE!
        ReDim NowProj.UserHierarchy.PropertySheetOrder(1 To UB)
      Else
        ReDim Preserve NowProj.UserHierarchy.PropertySheetOrder(1 To UB)
      End If
      ReDim NowProj.UserHierarchy.PropertySheetOrder(UB).PropertyOrder(0 To 0)
      NowProj.UserHierarchy.PropertySheetOrder(UB).Name = New_Name
      '
      ' REFRESH WINDOW.
      '
      Call frmCustomProperties_Refresh
    '
    '////////////////////////////////////////////////////////////////////
    Case 1:           'DELETE PROPERTY SHEET.
      If (idx_PropSheet = 0) Then
        Beep
        GoTo exit_err_ThisFunc
      End If
      If ((Trim$(UCase$(Name_PropSheet)) = Trim$(UCase$(PROPERTYSHEETNAME_BASIC_CHEMICAL_INFO))) Or _
          (Trim$(UCase$(Name_PropSheet)) = Trim$(UCase$(PROPERTYSHEETNAME_CHEMICAL_NOTE)))) Then
        Call Show_Error("You cannot delete or rename the property sheets " & _
            "named `" & PROPERTYSHEETNAME_BASIC_CHEMICAL_INFO & _
            "` or `" & PROPERTYSHEETNAME_CHEMICAL_NOTE & "`.")
        GoTo exit_err_ThisFunc
      End If
      If (vbNo = MsgBox("Are you sure you want to delete the property " & _
          "sheet named `" & Name_PropSheet & "`?  This step cannot be undone.", _
          vbQuestion + vbYesNo, _
          "Delete property sheet: Are you sure?")) Then
        GoTo exit_normally_ThisFunc
      End If
      '
      ' DELETE IT.
      '
      UB = UBound(NowProj.UserHierarchy.PropertySheetOrder)
      For i = idx_PropSheet To UB - 1
        NowProj.UserHierarchy.PropertySheetOrder(i) = _
            NowProj.UserHierarchy.PropertySheetOrder(i + 1)
      Next i
      UB = UB - 1
      If (UB = 0) Then
        ' THEORETICALLY IMPOSSIBLE TO GET HERE!
        ReDim NowProj.UserHierarchy.PropertySheetOrder(0 To 0)
      Else
        ReDim Preserve NowProj.UserHierarchy.PropertySheetOrder(1 To UB)
      End If
      '
      ' REFRESH WINDOW.
      '
      Call frmCustomProperties_Refresh
    '
    '////////////////////////////////////////////////////////////////////
    Case 2:           'RENAME PROPERTY SHEET.
      If (idx_PropSheet = 0) Then
        Beep
        GoTo exit_err_ThisFunc
      End If
      If ((Trim$(UCase$(Name_PropSheet)) = Trim$(UCase$(PROPERTYSHEETNAME_BASIC_CHEMICAL_INFO))) Or _
          (Trim$(UCase$(Name_PropSheet)) = Trim$(UCase$(PROPERTYSHEETNAME_CHEMICAL_NOTE)))) Then
        Call Show_Error("You cannot delete or rename the property sheets " & _
            "named `" & PROPERTYSHEETNAME_BASIC_CHEMICAL_INFO & _
            "` or `" & PROPERTYSHEETNAME_CHEMICAL_NOTE & "`.")
        GoTo exit_err_ThisFunc
      End If
      Do While (1 = 1)
        New_Name = frmNewName.frmNewName_GetName( _
            "Rename property sheet", _
            "Enter new name for the selected property sheet.  Each " & _
            "property sheet must have a unique name.", _
            Name_PropSheet, _
            is_aborted)
        If (is_aborted) Then GoTo exit_normally_ThisFunc
        If ((Trim$(UCase$(New_Name)) = Trim$(UCase$(PROPERTYSHEETNAME_BASIC_CHEMICAL_INFO))) Or _
            (Trim$(UCase$(New_Name)) = Trim$(UCase$(PROPERTYSHEETNAME_CHEMICAL_NOTE)))) Then
          Call Show_Error("You cannot rename a property sheet " & _
              "to `" & PROPERTYSHEETNAME_BASIC_CHEMICAL_INFO & _
              "` or `" & PROPERTYSHEETNAME_CHEMICAL_NOTE & "`.  " & _
              "Please enter a different name or hit Cancel.")
        Else
          If (PropertySheetOrder_IsKeyExist(New_Name) = False) Then
            Exit Do
          Else
            Call Show_Error("The name `" & New_Name & _
                "` already exists.  Please enter a different name or hit Cancel.")
          End If
        End If
      Loop
      '
      ' RENAME IT.
      '
      NowProj.UserHierarchy.PropertySheetOrder(idx_PropSheet).Name = New_Name
      '
      ' REFRESH WINDOW.
      '
      Call frmCustomProperties_Refresh
    '
    '////////////////////////////////////////////////////////////////////
    Case 3:           'MOVE PROPERTY SHEET UP.
      If (idx_PropSheet = 0) Then
        Beep
        GoTo exit_err_ThisFunc
      End If
      If (idx_PropSheet = 1) Then
        Beep
        GoTo exit_err_ThisFunc
      End If
      '
      ' SWAP THE TWO VALUES.
      '
      Temp_PropSheetOrder = _
          NowProj.UserHierarchy.PropertySheetOrder(idx_PropSheet)
      NowProj.UserHierarchy.PropertySheetOrder(idx_PropSheet) = _
          NowProj.UserHierarchy.PropertySheetOrder(idx_PropSheet - 1)
      NowProj.UserHierarchy.PropertySheetOrder(idx_PropSheet - 1) = _
          Temp_PropSheetOrder
      '
      ' REFRESH WINDOW.
      '
      Call frmCustomProperties_Refresh
    '
    '////////////////////////////////////////////////////////////////////
    Case 4:           'MOVE PROPERTY SHEET DOWN.
      If (idx_PropSheet = 0) Then
        Beep
        GoTo exit_err_ThisFunc
      End If
      If (idx_PropSheet = UBound(NowProj.UserHierarchy.PropertySheetOrder)) Then
        Beep
        GoTo exit_err_ThisFunc
      End If
      '
      ' SWAP THE TWO VALUES.
      '
      Temp_PropSheetOrder = _
          NowProj.UserHierarchy.PropertySheetOrder(idx_PropSheet)
      NowProj.UserHierarchy.PropertySheetOrder(idx_PropSheet) = _
          NowProj.UserHierarchy.PropertySheetOrder(idx_PropSheet + 1)
      NowProj.UserHierarchy.PropertySheetOrder(idx_PropSheet + 1) = _
          Temp_PropSheetOrder
      '
      ' REFRESH WINDOW.
      '
      Call frmCustomProperties_Refresh
  End Select
exit_normally_ThisFunc:
  ';xxxxx = True
  Exit Sub
exit_err_ThisFunc:
  'xxxxx = False
  Exit Sub
err_ThisFunc:
  Call Show_Trapped_Error("cmdBottomRight1Cmds_Click")
  Resume exit_err_ThisFunc
End Sub


Private Sub cmdTopRight1Cmds_Click(Index As Integer)
On Error GoTo err_ThisFunc
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim Temp_Storage As PropertyOrder_Type
Dim idx_PropSheet As Integer
Dim This_PropCode As Long
Dim UB_Selected As Integer
Dim out_idx_Elem As Integer
  idx_PropSheet = lstTop_PropSheets.ListIndex + 1
  If (idx_PropSheet = 0) Then
    ' THEORETICALLY IMPOSSIBLE TO GET HERE.
    Call Show_Error("You must first select a property sheet.")
    GoTo exit_err_ThisFunc
  End If
  ''''If (idx_PropSheet = 0) Then
  ''''  Name_PropSheet = ""
  ''''Else
  ''''  Name_PropSheet = lstTop_PropSheets.List(lstTop_PropSheets.ListIndex)
  ''''End If
  '
  ' DETERMINE SELECTION TYPE (IF ANY) IN lstTechniques.
  '
Dim Is_MultiSelection As Boolean
Dim Is_SingleSelection As Boolean
Dim Is_FirstSelected As Boolean
Dim Is_LastSelected As Boolean
Dim idx_FirstSelection As Integer
  Is_MultiSelection = False
  Is_SingleSelection = False
  Is_FirstSelected = False
  Is_LastSelected = False
  For i = 0 To lstTop_Props.ListCount - 1
    If (lstTop_Props.Selected(i) = True) Then
      If (i = 0) Then Is_FirstSelected = True
      If (i = lstTop_Props.ListCount - 1) Then Is_LastSelected = True
      If (Is_SingleSelection = False) And (Is_MultiSelection = False) Then
        Is_SingleSelection = True
        idx_FirstSelection = i + 1
      Else
        Is_SingleSelection = False
        Is_MultiSelection = True
      End If
    End If
  Next i
  Select Case Index
    '
    '////////////////////////////////////////////////////////////////////
    Case 0:       'SELECT PROPERTY/PROPERTIES.
      '
      ' ADD PROPERTY/PROPERTIES.
      '
      For i = 0 To lstTop_PropsAll.ListCount - 1
        If (lstTop_PropsAll.Selected(i) = True) Then
          This_PropCode = lstTop_PropsAll.ItemData(i)
          UB_Selected = UBound(NowProj.UserHierarchy.PropertySheetOrder(idx_PropSheet).PropertyOrder)
          UB_Selected = UB_Selected + 1
          If (UB_Selected = 1) Then
            ReDim NowProj.UserHierarchy.PropertySheetOrder(idx_PropSheet).PropertyOrder(1 To 1)
          Else
            ReDim Preserve NowProj.UserHierarchy.PropertySheetOrder(idx_PropSheet).PropertyOrder(1 To UB_Selected)
          End If
          With NowProj.UserHierarchy.PropertySheetOrder(idx_PropSheet).PropertyOrder(UB_Selected)
            .Property_Code = This_PropCode
            ReDim .Technique_Code(0 To 0)
            Call Get_Complete_List_of_TechCodes( _
                .Property_Code, _
                .Technique_Code)
          End With
        End If
      Next i
      '
      ' REFRESH WINDOW.
      '
      Call frmCustomProperties_Refresh
    '
    '////////////////////////////////////////////////////////////////////
    Case 1:       'SELECT ALL PROPERTIES.
      If (lstTop_PropsAll.ListCount = 0) Then
        Beep
        GoTo exit_err_ThisFunc
      End If
      For i = 0 To lstTop_PropsAll.ListCount - 1
        lstTop_PropsAll.Selected(i) = True
      Next i
      Call cmdTopRight1Cmds_Click(0)
    '
    '////////////////////////////////////////////////////////////////////
    Case 2:       'DESELECT PROPERTY/PROPERTIES.
'      '
'      ' DELETE PROPERTY/PROPERTIES.
'      '
'      For i = 0 To lstTop_Props.ListCount - 1
'        If (lstTop_Props.Selected(i) = True) Then
'          This_PropCode = lstTop_Props.ItemData(i)
'          UB_Selected = UBound(NowProj.UserHierarchy.PropertySheetOrder(idx_PropSheet).PropertyOrder)
'          UB_Selected = UB_Selected + 1
'          If (UB_Selected = 1) Then
'            ReDim NowProj.UserHierarchy.PropertySheetOrder(idx_PropSheet).PropertyOrder(1 To 1)
'          Else
'            ReDim Preserve NowProj.UserHierarchy.PropertySheetOrder(idx_PropSheet).PropertyOrder(1 To UB_Selected)
'          End If
'          With NowProj.UserHierarchy.PropertySheetOrder(idx_PropSheet).PropertyOrder(UB_Selected)
'            .Property_Code = This_PropCode
'            ReDim .Technique_Code(0 To 0)
'            Call Get_Complete_List_of_TechCodes( _
'                .Property_Code, _
'                .Technique_Code)
'          End With
'        End If
'      Next i
      '
      ' REMOVE PROPERTY/PROPERTIES.
      '
      For i = 0 To lstTop_Props.ListCount - 1
        If (lstTop_Props.Selected(i) = True) Then
          This_PropCode = lstTop_Props.ItemData(i)
          UB_Selected = UBound(NowProj.UserHierarchy.PropertySheetOrder(idx_PropSheet).PropertyOrder)
          For j = 1 To UB_Selected
            If (NowProj.UserHierarchy.PropertySheetOrder(idx_PropSheet).PropertyOrder(j).Property_Code = _
                This_PropCode) Then
              out_idx_Elem = j
              For k = out_idx_Elem To UB_Selected - 1
                NowProj.UserHierarchy.PropertySheetOrder(idx_PropSheet).PropertyOrder(k) = _
                    NowProj.UserHierarchy.PropertySheetOrder(idx_PropSheet).PropertyOrder(k + 1)
              Next k
              UB_Selected = UB_Selected - 1
              If (UB_Selected = 0) Then
                ReDim NowProj.UserHierarchy.PropertySheetOrder(idx_PropSheet).PropertyOrder(0 To 0)
              Else
                ReDim Preserve NowProj.UserHierarchy.PropertySheetOrder(idx_PropSheet).PropertyOrder(1 To UB_Selected)
              End If
            End If
            Exit For
          Next j
        End If
      Next i
      '
      ' REFRESH WINDOW.
      '
      Call frmCustomProperties_Refresh
    '
    '////////////////////////////////////////////////////////////////////
    Case 3:       'DESELECT ALL PROPERTIES.
      If (lstTop_Props.ListCount = 0) Then
        Beep
        GoTo exit_err_ThisFunc
      End If
      For i = 0 To lstTop_Props.ListCount - 1
        lstTop_Props.Selected(i) = True
      Next i
      Call cmdTopRight1Cmds_Click(2)
    '
    '////////////////////////////////////////////////////////////////////
    Case 4:       'MOVE UP.
      If (Is_MultiSelection = True) Or (Is_SingleSelection = False) Then
        Call Show_Error("To use this command, you must highlight " & _
            "only a single property.")
        GoTo exit_err_ThisFunc
      End If
      If (Is_FirstSelected) Then
        Beep
        GoTo exit_err_ThisFunc
      End If
      '
      ' SWAP THE TWO VALUES.
      '
      With NowProj.UserHierarchy.PropertySheetOrder(idx_PropSheet)
        Temp_Storage = .PropertyOrder(idx_FirstSelection)
        .PropertyOrder(idx_FirstSelection) = .PropertyOrder(idx_FirstSelection - 1)
        .PropertyOrder(idx_FirstSelection - 1) = Temp_Storage
      End With
      '
      ' REFRESH WINDOW.
      '
      Call frmCustomProperties_Refresh
    '
    '////////////////////////////////////////////////////////////////////
    Case 5:       'MOVE DOWN.
      If (Is_MultiSelection = True) Or (Is_SingleSelection = False) Then
        Call Show_Error("To use this command, you must highlight " & _
            "only a single property.")
        GoTo exit_err_ThisFunc
      End If
      If (Is_LastSelected) Then
        Beep
        GoTo exit_err_ThisFunc
      End If
      '
      ' SWAP THE TWO VALUES.
      '
      With NowProj.UserHierarchy.PropertySheetOrder(idx_PropSheet)
        Temp_Storage = .PropertyOrder(idx_FirstSelection)
        .PropertyOrder(idx_FirstSelection) = .PropertyOrder(idx_FirstSelection + 1)
        .PropertyOrder(idx_FirstSelection + 1) = Temp_Storage
      End With
      '
      ' REFRESH WINDOW.
      '
      Call frmCustomProperties_Refresh
  End Select
exit_normally_ThisFunc:
  ';xxxxx = True
  Exit Sub
exit_err_ThisFunc:
  'xxxxx = False
  Exit Sub
err_ThisFunc:
  Call Show_Trapped_Error("cmdBottomRight1Cmds_Click")
  Resume exit_err_ThisFunc
End Sub


Private Sub Form_Load()
  '
  ' MISC INITS.
  '
  USER_HIT_CANCEL = False
  USER_HIT_OK = False
  HALT_ALL_CONTROLS = False
  Me.Width = 9600
  Me.Height = 7200
  Call CenterOnForm(Me, frmMain)
  ''''Call frmPrefEnvironment_PopulateFirstTime_SeveralControls
  '
  ' REMOVE BEVELS FROM PANEL CONTROLS.
  '
  sspAll.BevelWidth = 0
  sspTop.BevelWidth = 0
  sspBottom.BevelWidth = 0
  sspMainButtons.BevelWidth = 0
  '
  ' FIRST REFRESH AND RESIZING.
  '
  Call frmCustomProperties_Resize
  Call frmCustomProperties_Refresh
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If (USER_HIT_CANCEL = False) And (USER_HIT_OK = False) Then
    Call Show_Error("To exit this window, please press either the " & _
        "`Cancel` button or the `Accept` button.")
    Cancel = True
    Exit Sub
  End If
End Sub
Private Sub Form_Resize()
  If (Me.WindowState <> vbMinimized) Then
    '
    ' WARNING: RESIZING WHILE MINIMIZED CAN CAUSE ERRORS!
    '
    Call frmCustomProperties_Resize
  End If
End Sub


Private Sub lstBottom_Props_Click()
  If (HALT_ALL_CONTROLS = True) Then Exit Sub
  Call frmCustomProperties_Refresh
  On Error Resume Next
  lstBottom_Props.SetFocus
End Sub
Private Sub lstTop_PropSheets_Click()
  If (HALT_ALL_CONTROLS = True) Then Exit Sub
  Call frmCustomProperties_Refresh
  On Error Resume Next
  lstTop_PropSheets.SetFocus
End Sub




