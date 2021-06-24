VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmContaminantPropertyEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contaminant Properties--{ModelName}"
   ClientHeight    =   5640
   ClientLeft      =   960
   ClientTop       =   1905
   ClientWidth     =   7680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cbochemname 
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
      Left            =   3150
      Style           =   2  'Dropdown List
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   180
      Width           =   4455
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   210
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   4740
      Width           =   2175
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   210
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4260
      Width           =   2175
   End
   Begin VB.TextBox txtTestDDE 
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   2
      Left            =   330
      TabIndex        =   13
      Text            =   "For DDE, not visible"
      Top             =   3720
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.TextBox txtTestDDE 
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   1
      Left            =   330
      TabIndex        =   12
      Text            =   "For DDE, not visible"
      Top             =   3420
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.TextBox txtTestDDE 
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   0
      Left            =   330
      TabIndex        =   11
      Text            =   "For DDE, not visible"
      Top             =   3120
      Visible         =   0   'False
      Width           =   1812
   End
   Begin Threed.SSFrame Frame1 
      Height          =   2055
      Left            =   30
      TabIndex        =   9
      Top             =   120
      Width           =   2535
      _Version        =   65536
      _ExtentX        =   4471
      _ExtentY        =   3625
      _StockProps     =   14
      Caption         =   "StEPP"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmdImport 
         Appearance      =   0  'Flat
         Caption         =   "Import StEPP Export File"
         Height          =   465
         Left            =   150
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   870
         Width           =   2250
      End
      Begin VB.CommandButton cmdImportClipboard 
         Appearance      =   0  'Flat
         Caption         =   "Import From Clipboard"
         Height          =   465
         Left            =   150
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1320
         Width           =   2250
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Obtain component properties from StEPP:"
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   150
         TabIndex        =   18
         Top             =   360
         Width           =   2295
      End
   End
   Begin Threed.SSFrame fraContaminantProperties 
      Height          =   4335
      Left            =   2730
      TabIndex        =   10
      Top             =   900
      Width           =   4875
      _Version        =   65536
      _ExtentX        =   8599
      _ExtentY        =   7646
      _StockProps     =   14
      Caption         =   "Pick Flow and Loading Parameters to Specify:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtConcentration 
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
         Index           =   1
         Left            =   2190
         TabIndex        =   8
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox txtConcentration 
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
         Left            =   2190
         TabIndex        =   7
         Top             =   3480
         Width           =   1215
      End
      Begin VB.TextBox txtContaminantProperties 
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
         Index           =   5
         Left            =   2190
         TabIndex        =   6
         Top             =   2940
         Width           =   1215
      End
      Begin VB.TextBox txtContaminantProperties 
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
         Index           =   4
         Left            =   2190
         TabIndex        =   5
         Top             =   2580
         Width           =   1215
      End
      Begin VB.TextBox txtContaminantProperties 
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
         Index           =   3
         Left            =   2190
         TabIndex        =   4
         Top             =   2220
         Width           =   1215
      End
      Begin VB.TextBox txtContaminantProperties 
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
         Index           =   2
         Left            =   2190
         TabIndex        =   3
         Top             =   1860
         Width           =   1215
      End
      Begin VB.TextBox txtContaminantProperties 
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
         Index           =   1
         Left            =   2190
         TabIndex        =   2
         Top             =   1500
         Width           =   1215
      End
      Begin VB.TextBox txtContaminantProperties 
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
         Left            =   2190
         TabIndex        =   1
         Top             =   1140
         Width           =   1215
      End
      Begin VB.TextBox txtContaminantName 
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
         Left            =   150
         TabIndex        =   0
         Top             =   600
         Width           =   4455
      End
      Begin VB.ComboBox UnitsProp 
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
         Index           =   0
         Left            =   3450
         Style           =   2  'Dropdown List
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1140
         Width           =   1155
      End
      Begin VB.ComboBox UnitsProp 
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
         Left            =   3450
         Style           =   2  'Dropdown List
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1860
         Width           =   1155
      End
      Begin VB.ComboBox UnitsProp 
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
         Left            =   3450
         Style           =   2  'Dropdown List
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   2220
         Width           =   1155
      End
      Begin VB.ComboBox UnitsProp 
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
         Left            =   3450
         Style           =   2  'Dropdown List
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   2580
         Width           =   1155
      End
      Begin VB.ComboBox UnitsProp 
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
         Left            =   3450
         Style           =   2  'Dropdown List
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   2940
         Width           =   1155
      End
      Begin VB.ComboBox UnitsConc 
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
         Index           =   0
         Left            =   3450
         Style           =   2  'Dropdown List
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   3480
         Width           =   1155
      End
      Begin VB.ComboBox UnitsConc 
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
         Left            =   3450
         Style           =   2  'Dropdown List
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   3840
         Width           =   1155
      End
      Begin VB.Label lblContaminantProperties 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Treatment Obj."
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
         Left            =   -270
         TabIndex        =   36
         Top             =   3840
         Width           =   2295
      End
      Begin VB.Label lblContaminantProperties 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Influent Conc."
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
         Left            =   -270
         TabIndex        =   35
         Top             =   3480
         Width           =   2295
      End
      Begin VB.Label lblContaminantProperties 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Gas Diffusivity"
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
         Left            =   -270
         TabIndex        =   34
         Top             =   2940
         Width           =   2295
      End
      Begin VB.Label lblContaminantProperties 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Liquid Diffusivity"
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
         Left            =   -270
         TabIndex        =   33
         Top             =   2580
         Width           =   2295
      End
      Begin VB.Label lblContaminantProperties 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Normal Boiling Point"
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
         Left            =   -270
         TabIndex        =   32
         Top             =   2220
         Width           =   2295
      End
      Begin VB.Label lblContaminantProperties 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Molar Volume"
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
         Left            =   -270
         TabIndex        =   31
         Top             =   1860
         Width           =   2295
      End
      Begin VB.Label lblContaminantProperties 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Henry's Constant"
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
         Left            =   -270
         TabIndex        =   30
         Top             =   1500
         Width           =   2295
      End
      Begin VB.Label lblContaminantProperties 
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
         Left            =   -270
         TabIndex        =   29
         Top             =   1140
         Width           =   2295
      End
      Begin VB.Label lblContaminantProperties 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
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
         Left            =   150
         TabIndex        =   28
         Top             =   300
         Width           =   675
      End
      Begin VB.Line Line1 
         X1              =   -60
         X2              =   5040
         Y1              =   1020
         Y2              =   1020
      End
      Begin VB.Line Line2 
         X1              =   -90
         X2              =   4950
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Label lblMisc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "( - )"
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
         Left            =   3450
         TabIndex        =   27
         Top             =   1500
         Width           =   1155
      End
   End
End
Attribute VB_Name = "frmContaminantPropertyEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Temp_Text As String
Dim ThisContaminant As ContaminantPropertyType

Dim loading As Integer

Sub LOCAL___Reset_DemoVersionDisablings()
  If (IsThisADemo() = True) Then
    cmdImport.Enabled = False
    cmdImportClipboard.Enabled = False
    cmdOK.Enabled = False
  End If
End Sub


Private Sub cbochemname_Click()
Dim i As Integer
Dim x As rec_frmContaminantPropertyEdit
Dim xu As rec_Units_frmContaminantPropertyEdit
  
 If loading Then
      Screen.MousePointer = 11

      'change name on list
      x = Data_frmContaminantPropertyEdit
     cbochemname.List(x.DoEditNumber - 1) = txtContaminantName.Text

      'save changes
      x.Contaminants(x.DoEditNumber) = ThisContaminant

    'setup new chemical
    ThisContaminant = x.Contaminants(cbochemname.ListIndex + 1)
    x.DoEditNumber = cbochemname.ListIndex + 1
    
    Data_frmContaminantPropertyEdit = x
    
    'display new chemical
    txtContaminantName.Text = ThisContaminant.Name
    txtContaminantName.Enabled = True
    txtContaminantProperties(0).Text = Format$(ThisContaminant.MolecularWeight.value, "0.00")
    txtContaminantProperties(1).Text = Format$(ThisContaminant.HenrysConstant.value, GetTheFormat(ThisContaminant.HenrysConstant.value))
    txtContaminantProperties(2).Text = Format$(ThisContaminant.MolarVolume.value, GetTheFormat(ThisContaminant.MolarVolume.value))
    txtContaminantProperties(3).Text = Format$(ThisContaminant.NormalBoilingPoint.value - 273.15, "0.0")
    txtContaminantProperties(4).Text = Format$(ThisContaminant.LiquidDiffusivity.value, GetTheFormat(ThisContaminant.LiquidDiffusivity.value))
    txtContaminantProperties(5).Text = Format$(ThisContaminant.GasDiffusivity.value, GetTheFormat(ThisContaminant.GasDiffusivity.value))
    txtConcentration(0).Text = Format$(ThisContaminant.Influent.value, GetTheFormat(ThisContaminant.Influent.value))
    txtConcentration(1).Text = Format$(ThisContaminant.TreatmentObjective.value, GetTheFormat(ThisContaminant.TreatmentObjective.value))
     
    
    xu = Units_frmContaminantPropertyEdit
    Call LabelsPropContaminant(UNITSTYPE_SI)
    For i = 0 To 5
      If (i <> 1) Then Call SetUnits(UnitsProp(i), xu.UnitsProp(i))
    Next i
    For i = 0 To 1
      Call SetUnits(UnitsConc(i), xu.UnitsConc(i))
    Next i

    Data_frmContaminantPropertyEdit = x
    Screen.MousePointer = 0
 End If

End Sub

Private Sub cmdCancel_Click()
Dim x As rec_frmContaminantPropertyEdit

  x = Data_frmContaminantPropertyEdit
  If (x.DoAdd) Then
    x.CancelledEdit = False
    x.CancelledAdd = True
  Else
    x.CancelledEdit = True
    x.CancelledAdd = False
  End If
  Data_frmContaminantPropertyEdit = x
    
  '  Scr1.Contaminant(Scr1.Chemical).Pressure = OriginalProperties.Pressure
  '  Scr1.Contaminant(Scr1.Chemical).Temperature = OriginalProperties.Temperature
  '  Scr1.Contaminant(Scr1.Chemical).Name = OriginalProperties.Name
  '  Scr1.Contaminant(Scr1.Chemical).MolecularWeight.Value = OriginalProperties.MolecularWeight.Value
  '  Scr1.Contaminant(Scr1.Chemical).HenrysConstant.Value = OriginalProperties.HenrysConstant.Value
  '  Scr1.Contaminant(Scr1.Chemical).MolarVolume.Value = OriginalProperties.MolarVolume.Value
  '  Scr1.Contaminant(Scr1.Chemical).NormalBoilingPoint.Value = OriginalProperties.NormalBoilingPoint.Value
  '  Scr1.Contaminant(Scr1.Chemical).LiquidDiffusivity.Value = OriginalProperties.LiquidDiffusivity.Value
  '  Scr1.Contaminant(Scr1.Chemical).GasDiffusivity.Value = OriginalProperties.GasDiffusivity.Value
  '  Scr1.Contaminant(Scr1.Chemical).Influent.Value = OriginalProperties.Influent.Value
  '  Scr1.Contaminant(Scr1.Chemical).TreatmentObjective.Value = OriginalProperties.TreatmentObjective.Value
  '
  '
  '  If (AddFlag) Then
  '    Scr1.NumChemical = Scr1.NumChemical - 1
  '    Scr1.Chemical = Scr1.NumChemical
  '  End If

  Unload Me

End Sub

Private Sub cmdImport_Click()
Dim f As Integer
Dim LineCount As Integer
Dim ThisLine As String
Dim AllLines As String
Dim InvalidFile As Integer
Const MAX_LINE_COUNT = 1000     'SOMEWHAT ARBITRARY.
Dim Ctl As Control
Set Ctl = frmBubble.CommonDialog1
Dim Use_Filename As String
  
  On Error GoTo err_cmdImportFromFile_Click
  Ctl.CancelError = True
  Ctl.DialogTitle = "Load StEPP Export File"
  Ctl.Filter = "All Files (*.*)|*.*|StEPP Export Files (*.exp)|*.exp"
  Ctl.FilterIndex = 2
  Ctl.Action = 1
  Use_Filename = Ctl.Filename
  If (Use_Filename = "") Then
    Exit Sub
  End If
  f = FreeFile
  LineCount = 0
  Open Use_Filename For Input As #f
  ''''Use_Filename = ""
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
    MsgBox "This is not a valid StEPP export file."
    Exit Sub
  End If
  'DO THE IMPORT.
  Clipboard.SetText AllLines
  Call cmdImportClipboard_Click
  Exit Sub
exit_err_cmdImportFromFile_Click:
  Exit Sub
err_cmdImportFromFile_Click:
  If (Err = 32755) Then
    'DO NOTHING.
  Else
    MsgBox "An error #" & Trim$(Str$(Err)) & " occurred."
  End If
  Resume exit_err_cmdImportFromFile_Click




'Dim f As Integer
'Dim s1 As String
'Dim s2 As String
'Dim s3 As String
'Dim msg As String
'Dim response As Integer
'Dim n As Integer
'Dim import_count As Integer
'Dim now_chemical As String
'Dim now_cas As String
'Dim ImportFile As String
'Dim this_chemical As String
'Dim thiscomp As ContaminantPropertyType
'Dim xxx As StEPPLink_Property
'Dim i As Integer
'
  'On Error GoTo ImportError
  ''tell user find the file
  'On Error Resume Next
  'frmBubble!CMDialog1.Filename = ""
  'frmBubble!CMDialog1.DefaultExt = "exp"
  'frmBubble!CMDialog1.Filter = "StEPP Export Files (*.exp)|*.exp"
  'frmBubble!CMDialog1.DialogTitle = "Load StEPP Export File"
  'frmBubble!CMDialog1.Flags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST
  'frmBubble!CMDialog1.Action = 1
  'ImportFile = frmBubble!CMDialog1.Filename
  'frmBubble!CMDialog1.Filename = ""
  'If Not fileexists(ImportFile) Then
  '  Exit Sub
  'End If
  '
  ''read it in
  'n = 0
  'f = FreeFile
  'Open ImportFile For Input As #f

'  Input #f, s1
'
'  ' make sure valid export file from stepp
'  If s1 <> "1234567890" Then
'      MsgBox "File not valid Stepp export file", 48, "Asap-File import"
'      Exit Sub
'  End If
'
'  Input #f, s1
'    frmStEPPLink_Pressure = CDbl(Trim(s1))
'  Input #f, s1
'    frmStEPPLink_Temperature = CDbl(Trim(s1))
'
'  'handle temp pressure check
'  If frmBubble.Visible = True Then
'
'    If ((frmStEPPLink_Pressure <> frmBubble!txtOperatingPressure) Or (frmStEPPLink_Temperature <> frmBubble!txtOperatingTemperature)) Then
'
'   msg = "The physical property file from StEPP contains properties at a pressure of "
'   msg = msg & frmStEPPLink_Pressure & " and a temperature " & frmStEPPLink_Temperature & "."
'   msg = msg & " Do you wish to continue importing the data?"
'
'   response = MsgBox(msg, MB_ICONSTOP + MB_YESNO, Application_Name)
'   If response = 7 Then Exit Sub
'    End If
'
'  ElseIf frmSurface.Visible = True Then
'    If ((frmStEPPLink_Pressure <> frmSurface!txtOperatingPressure) Or (frmStEPPLink_Temperature <> frmSurface!txtOperatingTemperature)) Then
'
'   msg = "The physical property file from StEPP contains properties at a pressure of "
'   msg = msg & frmStEPPLink_Pressure & " and a temperature " & frmStEPPLink_Temperature & "."
'   msg = msg & " Do you wish to continue importing the data?"
'
'   response = MsgBox(msg, MB_ICONSTOP + MB_YESNO, Application_Name)
'   If response = 7 Then Exit Sub
'    End If
'
'  ElseIf frmPTADScreen1.Visible = True Then
'    If ((frmStEPPLink_Pressure <> frmPTADScreen1!txtOperatingPressure) Or (frmStEPPLink_Temperature <> frmPTADScreen1!txtOperatingTemperature)) Then
'
'     msg = "The physical property file from StEPP contains properties at a pressure of "
'     msg = msg & frmStEPPLink_Pressure & " and a temperature " & frmStEPPLink_Temperature & "."
'     msg = msg & " Do you wish to continue importing the data?"
'
'     response = MsgBox(msg, MB_ICONSTOP + MB_YESNO, Application_Name)
'     If response = 7 Then Exit Sub
'    End If
'
'  Else
'    If ((frmStEPPLink_Pressure <> frmPTADScreen2!txtOperatingPressure) Or (frmStEPPLink_Temperature <> frmPTADScreen2!txtOperatingTemperature)) Then
'
'     msg = "The physical property file from StEPP contains properties at a pressure of "
'     msg = msg & frmStEPPLink_Pressure & " and a temperature " & frmStEPPLink_Temperature & "."
'     msg = msg & " Do you wish to continue importing the data?"
'
'     response = MsgBox(msg, MB_ICONSTOP + MB_YESNO, Application_Name)
'     If response = 7 Then Exit Sub
'    End If
'
'  End If
'
'
'  Do While (1 = 1)
'    If (EOF(f)) Then Exit Do
'    Input #f, s1, s2, s3
'    If (s1 = "END_OF_FILE") Then Exit Do
'    If (UCase$(s1) = "CHEMICAL") Then
'
'    If frmSurface.Visible = True Then
'       response = 0
'
'       For i = 1 To frmSurface!cboDesignContaminant.ListCount
'         If Trim(s2) = frmSurface!cboDesignContaminant.List(i) Then
'            response = 1
'         End If
'       Next i
'
'    ElseIf frmBubble.Visible = True Then
'       response = 0
'
'       For i = 1 To frmBubble!cboDesignContaminant.ListCount
'         If Trim(s2) = frmBubble!cboDesignContaminant.List(i) Then
'            response = 1
'         End If
'       Next i
'
'    ElseIf frmPTADScreen1.Visible = True Then
'       response = 0
'
'       For i = 1 To frmPTADScreen1!cboSelectCompo.ListCount
'         If Trim(s2) = frmPTADScreen1!cboSelectCompo.List(i) Then
'            response = 1
'         End If
'       Next i
'
'    Else
'       response = 0
'
'       For i = 1 To frmPTADScreen2!cboSelectCompo.ListCount
'         If Trim(s2) = frmPTADScreen2!cboSelectCompo.List(i) Then
'            response = 1
'         End If
'       Next i
'    End If
'
'      If response = 1 Then
'         s2 = ""
'      Else
'         now_chemical = s2
'         now_cas = s3
'         import_count = import_count + 1
'      End If
'    Else
'      N = N + 1
'      ReDim Preserve StEPPLink_AllProps(1 To N)
'      StEPPLink_AllProps(N).chemical = now_chemical
'      StEPPLink_AllProps(N).CAS = now_cas
'      StEPPLink_AllProps(N).propname = s1
'      StEPPLink_AllProps(N).units = s2
'      If (UCase$(s3) <> "UNAVAILABLE") Then
'        StEPPLink_AllProps(N).Val = CDbl(s3)
'        StEPPLink_AllProps(N).avail = True
'      Else
'        StEPPLink_AllProps(N).Val = 0#
'        StEPPLink_AllProps(N).avail = False
'      End If
'    End If
'  Loop
'
'  Close #f
'
'
''********** take stuff read in and store it corecctly
'    ReDim StEPPLink_RequiredProps(1 To 6)
'    StEPPLink_RequiredProps(1) = "HenrysConstant"
'    StEPPLink_RequiredProps(2) = "MolecularWeight"
'    StEPPLink_RequiredProps(3) = "NormalBoilingPoint"
'    StEPPLink_RequiredProps(4) = "MolarVolumeAtNBP"
'    StEPPLink_RequiredProps(5) = "LiquidDiffusivity"
'    StEPPLink_RequiredProps(6) = "GasDiffusivity"
'
'    ReDim StEPPLink_CurrentChemicalNames(0 To import_count)
'
'    '-- Call routine to filter out components that lack any of
'    '.. the properties listed above
'    Call StEPPLink_FilterUnimportable
'
'    '-- CLIENT-SPECIFIC: Import the acceptable components to memory structures
'    For i = 1 To UBound(StEPPLink_ImportSucceeded_Name)
'      this_chemical = StEPPLink_ImportSucceeded_Name(i)
'
'      'Set all property defaults
'      'Call SetComponentDefaults(thiscomp, -1)
'      ThisComp.Influent.Value = 10#
'      ThisComp.TreatmentObjective.Value = 2#
'      XXX = StEPPLink_FilteredProps(StEPPLink_FindProp(this_chemical, "MolecularWeight"))
'      ThisComp.Name = XXX.chemical
'      'thiscomp.Cas = CLng(xxx.Cas)
'
'      'Import each property
'      XXX = StEPPLink_FilteredProps(StEPPLink_FindProp(this_chemical, "MolecularWeight"))
'      ThisComp.MolecularWeight.Value = XXX.Val
'      ThisComp.MolecularWeight.ValChanged = True
'      ThisComp.MolecularWeight.UserInput = False
'
'      XXX = StEPPLink_FilteredProps(StEPPLink_FindProp(this_chemical, "HenrysConstant"))
'      ThisComp.HenrysConstant.Value = XXX.Val
'      ThisComp.HenrysConstant.ValChanged = True
'      ThisComp.HenrysConstant.UserInput = False
'
'      XXX = StEPPLink_FilteredProps(StEPPLink_FindProp(this_chemical, "NormalBoilingPoint"))
'      ThisComp.NormalBoilingPoint.Value = (XXX.Val) + 273.15
'      ThisComp.NormalBoilingPoint.ValChanged = True
'      ThisComp.NormalBoilingPoint.UserInput = False
'
'      XXX = StEPPLink_FilteredProps(StEPPLink_FindProp(this_chemical, "MolarVolumeAtNBP"))
'      ThisComp.MolarVolume.Value = XXX.Val
'      ThisComp.MolarVolume.ValChanged = True
'      ThisComp.MolarVolume.UserInput = False
'
'      XXX = StEPPLink_FilteredProps(StEPPLink_FindProp(this_chemical, "LiquidDiffusivity"))
'      ThisComp.LiquidDiffusivity.Value = XXX.Val
'      ThisComp.LiquidDiffusivity.ValChanged = True
'      ThisComp.LiquidDiffusivity.UserInput = False
'      'thiscomp.Refractive_Index = xxx.val
'
'      XXX = StEPPLink_FilteredProps(StEPPLink_FindProp(this_chemical, "GasDiffusivity"))
'      ThisComp.GasDiffusivity.Value = XXX.Val
'      ThisComp.GasDiffusivity.ValChanged = True
'      ThisComp.GasDiffusivity.UserInput = False
'
'      'Place imported component into memory
'      Data_frmContaminantPropertyEdit.NewNumCompo = Data_frmContaminantPropertyEdit.NewNumCompo + 1
'      Data_frmContaminantPropertyEdit.Contaminants(Data_frmContaminantPropertyEdit.NewNumCompo) = ThisComp
'    Next i
'
'    '-- Display standard "Import Succeeded" dialog box
'
'    StEPPLink_DontForget = "Remember to set the correct values of influent and treatment objective concentrations for these components."
'    Call StEPPLink_DisplayImportSucceeded
'
'    Unload Me
'
'ChDrive App.Path
'ChDir App.Path
'
'Exit Sub
'
'exit_ImportError:
'  Exit Sub
'
'ImportError:
'  ChDrive App.Path
'  ChDir App.Path
'  If Err = 32755 Then   'Cancel selected by user
'    Resume exit_ImportError
'  End If
'  MsgBox "Error importing file!", 48, "ASAP StEPP Import"
'  Resume exit_ImportError

End Sub

Private Sub cmdImportClipboard_Click()
Dim num_lines As Integer
Dim cliptext As String
Dim line_in As String
Dim r As Integer
Dim link_pressure As Double
Dim link_temperature As Double
Dim link_ChemCount As Integer
Const CHEMPROP_MIN = 0
Const CHEMPROP_MAX = 12
ReDim link_ChemProp(CHEMPROP_MIN To CHEMPROP_MAX, 1 To 1) As Double
ReDim link_ChemName(1 To 1) As String
ReDim link_ChemCAS(1 To 1) As String
ReDim link_ChemPropAvailable(CHEMPROP_MIN To CHEMPROP_MAX, 1 To 1) As Integer
ReDim link_IsImportable(1 To 1) As Integer
Dim i As Integer
Dim j As Integer
Const PROP_VAPORPRESSURE = 0
Const PROP_ACTIVITYCOEFFICIENT = 1
Const PROP_HENRYSCONSTANT = 2
Const PROP_MOLECULARWEIGHT = 3
Const PROP_NORMALBOILINGPOINT = 4
Const PROP_LIQUIDDENSITY = 5
Const PROP_MOLARVOLUMEATOPT = 6
Const PROP_MOLARVOLUMEATNBP = 7
Const PROP_REFRACTIVEINDEX = 8
Const PROP_AQUEOUSSOLUBILITY = 9
Const PROP_LOGKOW = 10
Const PROP_LIQUIDDIFFUSIVITY = 11
Const PROP_GASDIFFUSIVITY = 12
Dim num_imported As Integer
'Dim thiscomp As ComponentPropertyType
Dim thiscomp As ContaminantPropertyType
Dim msg As String
Dim vb3CrLf As String
Dim num_failed As Integer

  On Error GoTo err_cmdImportClipboard_Click
  cliptext = Clipboard.GetText()
  cliptext = Parser_RemoveCharacters(Chr$(10), cliptext)
  num_lines = Parser_GetNumArgs(Chr$(13), cliptext)
  r = 1
  Call Parser_GetArg(Chr$(13), cliptext, r, line_in)
  If (Trim$(UCase$(line_in)) <> Trim$(UCase$("1234567890:START_OF_STEPP_CLIPBOARD_EXPORT"))) Then
    GoTo err_nonfatal_cmdImportClipboard_Click
  End If
  r = r + 2
  Call Parser_GetArg(Chr$(13), cliptext, r, line_in)
  link_pressure = CDbl(Val(line_in))
  If (link_pressure <= 0#) Then GoTo err_nonfatal_cmdImportClipboard_Click
  r = r + 2
  Call Parser_GetArg(Chr$(13), cliptext, r, line_in)
  link_temperature = CDbl(Val(line_in))
  If (link_temperature <= 0#) Then GoTo err_nonfatal_cmdImportClipboard_Click
  r = r + 2
  Call Parser_GetArg(Chr$(13), cliptext, r, line_in)
  link_ChemCount = CInt(Val(line_in))
  If (link_ChemCount <= 0) Then GoTo err_nonfatal_cmdImportClipboard_Click
  ReDim link_ChemProp(CHEMPROP_MIN To CHEMPROP_MAX, 1 To link_ChemCount)
  ReDim link_ChemName(1 To link_ChemCount)
  ReDim link_ChemCAS(1 To link_ChemCount)
  ReDim link_ChemPropAvailable(CHEMPROP_MIN To CHEMPROP_MAX, 1 To link_ChemCount)
  ReDim link_IsImportable(1 To link_ChemCount)
  For i = 1 To link_ChemCount
    For j = CHEMPROP_MIN To CHEMPROP_MAX
      link_ChemPropAvailable(j, i) = True
    Next j
    r = r + 2
    Call Parser_GetArg(Chr$(13), cliptext, r, line_in)
    link_ChemName(i) = Trim$(UCase$(line_in))
    If (link_ChemName(i) = "") Then GoTo err_nonfatal_cmdImportClipboard_Click
    r = r + 2
    Call Parser_GetArg(Chr$(13), cliptext, r, line_in)
    link_ChemCAS(i) = Trim$(UCase$(line_in))
    'If (link_ChemCAS(i) = "") Then GoTo err_nonfatal_cmdImportClipboard_Click
    For j = CHEMPROP_MIN To CHEMPROP_MAX
      r = r + 2
      Call Parser_GetArg(Chr$(13), cliptext, r, line_in)
      line_in = Trim$(UCase$(line_in))
      If (Trim$(UCase$("UNAVAILABLE")) = line_in) Then
        link_ChemPropAvailable(j, i) = False
      Else
        link_ChemProp(j, i) = CDbl(Val(line_in))
      End If
    Next j
  Next i
  r = r + 1
  Call Parser_GetArg(Chr$(13), cliptext, r, line_in)
  If (Trim$(UCase$(line_in)) <> Trim$(UCase$("1234567890:END_OF_STEPP_CLIPBOARD_EXPORT"))) Then
    GoTo err_nonfatal_cmdImportClipboard_Click
  End If
  
  'ARE THERE ENOUGH EMPTY COMPONENT SLOTS REMAINING?
  If (frmBubble.Visible) Then
    If (bub.NumChemical + 1 > MAXCHEMICAL) Then GoTo err_TooManyChemicals
  End If
  If (frmSurface.Visible) Then
    If (sur.NumChemical + 1 > MAXCHEMICAL) Then GoTo err_TooManyChemicals
  End If
  If (frmPTADScreen1.Visible) Then
    If (scr1.NumChemical + 1 > MAXCHEMICAL) Then GoTo err_TooManyChemicals
  End If
  If (frmPTADScreen2.Visible) Then
    If (Scr2.NumChemical + 1 > MAXCHEMICAL) Then GoTo err_TooManyChemicals
  End If
  GoTo Bypass_err_TooManyChemicals

err_TooManyChemicals:
  MsgBox "Unable to import all of chemicals in file because the maximum number of chemicals has been reached.", 48, Application_Name
  Unload Me
  Exit Sub
Bypass_err_TooManyChemicals:

  'DOES THE USER REALLY WANT TO IMPORT AT THIS TEMPERATURE AND PRESSURE?
      '---- I'VE DECIDED TO SKIP THIS STEP.  THE USER BEWARE.
  
  'DETERMINE WHICH COMPONENTS ARE IMPORTABLE.
  For i = 1 To link_ChemCount
    link_IsImportable(i) = True
    If (Not link_ChemPropAvailable(PROP_HENRYSCONSTANT, i)) Then link_IsImportable(i) = False
    If (Not link_ChemPropAvailable(PROP_MOLECULARWEIGHT, i)) Then link_IsImportable(i) = False
    If (Not link_ChemPropAvailable(PROP_NORMALBOILINGPOINT, i)) Then link_IsImportable(i) = False
    If (Not link_ChemPropAvailable(PROP_MOLARVOLUMEATNBP, i)) Then link_IsImportable(i) = False
    If (Not link_ChemPropAvailable(PROP_LIQUIDDIFFUSIVITY, i)) Then link_IsImportable(i) = False
    If (Not link_ChemPropAvailable(PROP_GASDIFFUSIVITY, i)) Then link_IsImportable(i) = False
  Next i

  'IMPORT ALL IMPORTABLE COMPONENTS.
  num_imported = 0
  For i = 1 To link_ChemCount
    If (link_IsImportable(i)) Then
      num_imported = num_imported + 1
      
      'Call SetComponentDefaults(thiscomp, -1)
      'thiscomp.name = link_ChemName(i)
      'thiscomp.Cas = CLng(Val(link_ChemCAS(i)))
      'thiscomp.Vapor_Pressure = link_ChemProp(PROP_VAPORPRESSURE, i)
      'thiscomp.MW = link_ChemProp(PROP_MOLECULARWEIGHT, i)
      'thiscomp.BP = link_ChemProp(PROP_NORMALBOILINGPOINT, i)
      'thiscomp.Liquid_Density = link_ChemProp(PROP_LIQUIDDENSITY, i) / 1000#
      'thiscomp.MolarVolume = link_ChemProp(PROP_MOLARVOLUMEATNBP, i) * 1000#
      'thiscomp.Refractive_Index = link_ChemProp(PROP_REFRACTIVEINDEX, i)
      'thiscomp.Aqueous_Solubility = link_ChemProp(PROP_AQUEOUSSOLUBILITY, i)
      'Number_Component = Number_Component + 1
      'Component(Number_Component) = thiscomp

      'Set all property defaults
      thiscomp.Name = link_ChemName(i)
      thiscomp.Influent.value = 10#
      thiscomp.TreatmentObjective.value = 2#

      'IMPORT EACH PROPERTY.
      thiscomp.MolecularWeight.value = link_ChemProp(PROP_MOLECULARWEIGHT, i)
      thiscomp.MolecularWeight.ValChanged = True
      thiscomp.MolecularWeight.UserInput = False

      thiscomp.HenrysConstant.value = link_ChemProp(PROP_HENRYSCONSTANT, i)
      thiscomp.HenrysConstant.ValChanged = True
      thiscomp.HenrysConstant.UserInput = False

      thiscomp.NormalBoilingPoint.value = link_ChemProp(PROP_NORMALBOILINGPOINT, i) + 273.15
      thiscomp.NormalBoilingPoint.ValChanged = True
      thiscomp.NormalBoilingPoint.UserInput = False

      thiscomp.MolarVolume.value = link_ChemProp(PROP_MOLARVOLUMEATNBP, i)
      thiscomp.MolarVolume.ValChanged = True
      thiscomp.MolarVolume.UserInput = False

      thiscomp.LiquidDiffusivity.value = link_ChemProp(PROP_LIQUIDDIFFUSIVITY, i)
      thiscomp.LiquidDiffusivity.ValChanged = True
      thiscomp.LiquidDiffusivity.UserInput = False

      thiscomp.GasDiffusivity.value = link_ChemProp(PROP_GASDIFFUSIVITY, i)
      thiscomp.GasDiffusivity.ValChanged = True
      thiscomp.GasDiffusivity.UserInput = False
      
      'STORE COMPONENT INTO MEMORY.
      Data_frmContaminantPropertyEdit.NewNumCompo = Data_frmContaminantPropertyEdit.NewNumCompo + 1
      Data_frmContaminantPropertyEdit.Contaminants(Data_frmContaminantPropertyEdit.NewNumCompo) = thiscomp
    End If
  Next i

  'DISPLAY WARNING/SUCCESS MESSAGE.
  vb3CrLf = Chr$(13) & Chr$(10)
  If (num_imported <> 0) Then
    msg = "Successfully imported " & Trim$(Str$(num_imported)) & " component"
    If (num_imported <> 1) Then msg = msg & "s"
    msg = msg & " from StEPP:" & vb3CrLf
    For i = 1 To link_ChemCount
      If (link_IsImportable(i)) Then
        msg = msg & "    " & Trim$(link_ChemName(i)) & vb3CrLf
      End If
    Next i
    msg = msg & "The properties are for a "
    msg = msg & "pressure of " & Trim$(Str$(link_pressure)) & " Pa "
    msg = msg & "and a "
    msg = msg & "temperature of " & Trim$(Str$(link_temperature)) & " degrees Celcius." & vb3CrLf
    msg = msg & vb3CrLf
    msg = msg & "Don't forget to set the correct values of influent concentration "
    msg = msg & "and treatment objective concentration for each "
    msg = msg & "of these components." & vb3CrLf
  Else
    msg = "Unable to import any components from StEPP." & vb3CrLf
  End If
  num_failed = link_ChemCount - num_imported
  If (num_failed <> 0) Then
    msg = msg & vb3CrLf
    msg = msg & "Failed to import the following component"
    If (num_failed <> 1) Then msg = msg & "s"
    msg = msg & ":" & vb3CrLf
    For i = 1 To link_ChemCount
      If (Not link_IsImportable(i)) Then
        msg = msg & "    " & Trim$(link_ChemName(i)) & vb3CrLf
      End If
    Next i
    msg = msg & vb3CrLf
    msg = msg & "Important note: In order to successfully import a component "
    msg = msg & "from StEPP, the following properties must be available: "
    msg = msg & "Henry's constant, "
    msg = msg & "molecular weight, "
    msg = msg & "normal boiling point, "
    msg = msg & "molar volume at the normal boiling point, "
    msg = msg & "liquid diffusivity, "
    msg = msg & "and gas diffusivity.  "
    msg = msg & "To force an import to occur, you may modify the user input "
    msg = msg & "value of the unavailable properties from within StEPP."
    msg = msg & vb3CrLf
  End If
  MsgBox msg, MB_ICONINFORMATION, Application_Name
  Unload Me

exit_err_cmdImportClipboard_Click:
  Exit Sub
err_nonfatal_cmdImportClipboard_Click:
  MsgBox "An error occurred during the import process.", 48, Application_Name
  GoTo exit_err_cmdImportClipboard_Click
err_cmdImportClipboard_Click:
  MsgBox "An error occurred during the import process.", 48, Application_Name
  Resume exit_err_cmdImportClipboard_Click
End Sub

Private Sub cmdOK_Click()
Dim i As Integer, Response As Integer, msg As String
Dim NumPrevious As Integer 'Location of the previous occurrence of the current chemical (if any)
Dim x As rec_frmContaminantPropertyEdit
Dim xu As rec_Units_frmContaminantPropertyEdit

  x = Data_frmContaminantPropertyEdit
  x.CancelledEdit = False
  x.CancelledAdd = False
  
  If (x.DoAdd) Then
    x.NewNumCompo = x.OldNumCompo + 1
    NumPrevious = 0
    For i = 1 To scr1.Chemical - 1
      If scr1.Contaminant(i).Name = txtContaminantName.Text Then
        msg = txtContaminantName.Text & " already appears in your list" & Chr$(13)
        msg = msg + "of contaminants." & Chr$(13) & Chr$(13)
        msg = msg + "Do you wish to replace the previous property values?" & Chr$(13) & Chr$(13)
        Response = MsgBox(msg, 36, "Warning")
        NumPrevious = i
        Exit For
      End If
    Next i
    If (NumPrevious = 0) Then
      'Do nothing--calling subroutine will update its contaminant list.
    ElseIf (Response = IDYES) Then
      'Do nothing--calling subroutine will update its contaminant list.
    Else
      'Cancel the addition of this contaminant.
      x.CancelledEdit = False
      x.CancelledAdd = True
      x.NewNumCompo = x.OldNumCompo
    End If
  Else
    'Do nothing--calling subroutine will update its contaminant list.
    x.NewNumCompo = x.OldNumCompo
  End If
  x.Contaminants(x.DoEditNumber) = ThisContaminant
  Data_frmContaminantPropertyEdit = x

  'Save units from this screen.
  xu = Units_frmContaminantPropertyEdit
  For i = 0 To 5
    If (i <> 1) Then xu.UnitsProp(i) = GetUnits(UnitsProp(i))
  Next i
  For i = 0 To 1
    xu.UnitsConc(i) = GetUnits(UnitsConc(i))
  Next i
  Units_frmContaminantPropertyEdit = xu
    
  Select Case x.ModelType
    Case MODELTYPE_PACKEDTOWER
      If (frmPTADScreen1.Visible) Then
        frmPTADScreen1.cboSelectCompo.ListIndex = cbochemname.ListIndex
      ElseIf (frmPTADScreen2.Visible) Then
        frmPTADScreen2.cboSelectCompo.ListIndex = cbochemname.ListIndex
      End If
    Case MODELTYPE_SURFACE
      frmSurface.cboDesignContaminant.ListIndex = cbochemname.ListIndex
    Case MODELTYPE_BUBBLE
      frmBubble.cboDesignContaminant.ListIndex = cbochemname.ListIndex
  End Select


'    If (AddFlag) Then
'          NumPrevious = 0
'          For i = 1 To Scr1.Chemical - 1
'              If Scr1.Contaminant(i).Name = txtContaminantName.Text Then
'                 Msg = txtContaminantName.Text & " already appears in your list" & Chr$(13)
'                 Msg = Msg + "of contaminants." & Chr$(13) & Chr$(13)
'                 Msg = Msg + "Do you wish to replace the previous property values?" & Chr$(13) & Chr$(13)
'                 Response = MsgBox(Msg, 36, "Warning")
'                 NumPrevious = i
'                 Exit For
'              End If
'          Next i
'
'          If NumPrevious = 0 Then
'             Scr1.NumChemical = Scr1.NumChemical + 1
'             Scr1.Contaminant(Scr1.NumChemical) = Scr1.Contaminant(0)
'
'             frmListContaminant.ListContaminants.AddItem txtContaminantName.Text
'             frmPTADScreen1.cboDesignContaminant.AddItem txtContaminantName.Text
'             frmPTADScreen1.cboSelectCompo.AddItem txtContaminantName.Text
'
'
'             If Not frmListContaminant.ListContaminants.Visible Then frmListContaminant.ListContaminants.Visible = True
'             'frmListContaminant.ListContaminants.Selected(Scr1.Chemical - 1) = True
'          ElseIf Response = IDYES Then  'Replace previous values of properties for the duplicate chemical with the new values.  Do not move the chemical from its current location in the list.
'             Scr1.Contaminant(NumPrevious).Pressure = Scr1.Contaminant(Scr1.Chemical).Pressure
'             Scr1.Contaminant(NumPrevious).Temperature = Scr1.Contaminant(Scr1.Chemical).Temperature
'
'             Scr1.Contaminant(NumPrevious).MolecularWeight.Value = Scr1.Contaminant(Scr1.Chemical).MolecularWeight.Value
'             Scr1.Contaminant(NumPrevious).MolecularWeight.ValChanged = Scr1.Contaminant(Scr1.Chemical).MolecularWeight.ValChanged
'             Scr1.Contaminant(NumPrevious).MolecularWeight.UserInput = Scr1.Contaminant(Scr1.Chemical).MolecularWeight.UserInput
'
'             Scr1.Contaminant(NumPrevious).HenrysConstant.Value = Scr1.Contaminant(Scr1.Chemical).HenrysConstant.Value
'             Scr1.Contaminant(NumPrevious).HenrysConstant.ValChanged = Scr1.Contaminant(Scr1.Chemical).HenrysConstant.ValChanged
'             Scr1.Contaminant(NumPrevious).HenrysConstant.UserInput = Scr1.Contaminant(Scr1.Chemical).HenrysConstant.UserInput
'
'             Scr1.Contaminant(NumPrevious).MolarVolume.Value = Scr1.Contaminant(Scr1.Chemical).MolarVolume.Value
'             Scr1.Contaminant(NumPrevious).MolarVolume.ValChanged = Scr1.Contaminant(Scr1.Chemical).MolarVolume.ValChanged
'             Scr1.Contaminant(NumPrevious).MolarVolume.UserInput = Scr1.Contaminant(Scr1.Chemical).MolarVolume.UserInput
'
'             Scr1.Contaminant(NumPrevious).NormalBoilingPoint.Value = Scr1.Contaminant(Scr1.Chemical).NormalBoilingPoint.Value
'             Scr1.Contaminant(NumPrevious).NormalBoilingPoint.ValChanged = Scr1.Contaminant(Scr1.Chemical).NormalBoilingPoint.ValChanged
'             Scr1.Contaminant(NumPrevious).NormalBoilingPoint.UserInput = Scr1.Contaminant(Scr1.Chemical).NormalBoilingPoint.UserInput
'
'             Scr1.Contaminant(NumPrevious).LiquidDiffusivity.Value = Scr1.Contaminant(Scr1.Chemical).LiquidDiffusivity.Value
'             Scr1.Contaminant(NumPrevious).LiquidDiffusivity.ValChanged = Scr1.Contaminant(Scr1.Chemical).LiquidDiffusivity.ValChanged
'             Scr1.Contaminant(NumPrevious).LiquidDiffusivity.UserInput = Scr1.Contaminant(Scr1.Chemical).LiquidDiffusivity.UserInput
'
'             Scr1.Contaminant(NumPrevious).GasDiffusivity.Value = Scr1.Contaminant(Scr1.Chemical).GasDiffusivity.Value
'             Scr1.Contaminant(NumPrevious).GasDiffusivity.ValChanged = Scr1.Contaminant(Scr1.Chemical).GasDiffusivity.ValChanged
'             Scr1.Contaminant(NumPrevious).GasDiffusivity.UserInput = Scr1.Contaminant(Scr1.Chemical).GasDiffusivity.UserInput
'
'             Scr1.NumChemical = Scr1.NumChemical - 1
'             Scr1.Chemical = NumPrevious
'             frmListContaminant.ListContaminants.Selected(Scr1.Chemical - 1) = True
'          Else
'             Scr1.NumChemical = Scr1.NumChemical - 1
'             Scr1.Chemical = NumPrevious
'             frmListContaminant.ListContaminants.Selected(Scr1.Chemical - 1) = True
'          End If
'
'    End If

    'If frmListContaminant.mnuOptionsManipulateContaminant(1).Enabled = False Then
    '   frmListContaminant.mnuOptionsManipulateContaminant(1).Enabled = True
    '   frmListContaminant.mnuOptionsManipulateContaminant(3).Enabled = True
    '   frmListContaminant.mnuOptionsManipulateContaminant(4).Enabled = True
    '   frmListContaminant.mnuOptionsSave.Enabled = True
    '   frmListContaminant.mnuOptionsView.Enabled = True
    'End If

  Unload Me
    
End Sub

Private Sub cmdStEPP_Click_OLD()
'
'  'StEPPImportSuccess = False
'  'frmSteppLink.Show 1
'  'If (StEPPImportSuccess) Then
'  '  Unload Me
'  'End If
'
'Dim i As Integer
'Dim j As Integer
'Dim this_chemical As String
'Dim xxx As StEPPLink_Property
'Dim thiscomp As ContaminantPropertyType
'
'  '---- CLIENT-SPECIFIC: Minimize all forms but StEPP Link ...
'  frmContaminantPropertyEdit.Hide
'  If (frmPTADScreen1.Visible) Then
'    frmPTADScreen1.Hide
'    designtype = 1
'  End If
'  If (frmPTADScreen2.Visible) Then
'    frmPTADScreen2.Hide
'    designtype = 2
'  End If
'  If (frmBubble.Visible) Then frmBubble.Hide
'  If (frmSurface.Visible) Then frmSurface.Hide
'
'  '---- CLIENT-SPECIFIC: Miscellaneous variables
'  Select Case Data_frmContaminantPropertyEdit.ModelType
'    Case MODELTYPE_BUBBLE:
'      frmStEPPLink_Pressure = CDbl(bub.OperatingPressure.Value) * 101325#
'      frmStEPPLink_Temperature = CDbl(bub.operatingtemperature.Value) - 273.15
'    Case MODELTYPE_SURFACE:
'      frmStEPPLink_Pressure = CDbl(sur.OperatingPressure.Value) * 101325#
'      frmStEPPLink_Temperature = CDbl(sur.operatingtemperature.Value) - 273.15
'    Case MODELTYPE_PACKEDTOWER:
'      Select Case designtype
'        Case 1:
'          frmStEPPLink_Pressure = CDbl(scr1.OperatingPressure.Value) * 101325#
'          frmStEPPLink_Temperature = CDbl(scr1.operatingtemperature.Value) - 273.15
'        Case 2:
'          frmStEPPLink_Pressure = CDbl(Scr2.OperatingPressure.Value) * 101325#
'          frmStEPPLink_Temperature = CDbl(Scr2.operatingtemperature.Value) - 273.15
'      End Select
'  End Select
'  frmStEPPLink_ClientName = "ASAP"
'  ReDim StEPPLink_CurrentChemicalNames(0 To 0)
'  If (Data_frmContaminantPropertyEdit.OldNumCompo <> 0) Then
'    ReDim StEPPLink_CurrentChemicalNames(1 To Data_frmContaminantPropertyEdit.OldNumCompo)
'    For i = 1 To Data_frmContaminantPropertyEdit.OldNumCompo
'      StEPPLink_CurrentChemicalNames(i) = Trim$(UCase$(Data_frmContaminantPropertyEdit.Contaminants(i).Name))
'    Next i
'  End If
'
'  '---- Call StEPP Link into existence
'  frmSteppLink.Show 1
'
'  '---- Process imported properties
'  If (frmStEPPLink_Success) Then
'    '-- CLIENT-SPECIFIC: Store names of strictly required properties
'    ReDim StEPPLink_RequiredProps(1 To 6)
'    StEPPLink_RequiredProps(1) = "HenrysConstant"
'    StEPPLink_RequiredProps(2) = "MolecularWeight"
'    StEPPLink_RequiredProps(3) = "NormalBoilingPoint"
'    StEPPLink_RequiredProps(4) = "MolarVolumeAtNBP"
'    StEPPLink_RequiredProps(5) = "LiquidDiffusivity"
'    StEPPLink_RequiredProps(6) = "GasDiffusivity"
'
'    '-- Call routine to filter out components that lack any of
'    '.. the properties listed above
'    Call StEPPLink_FilterUnimportable
'
'    '-- CLIENT-SPECIFIC: Import the acceptable components to memory structures
'    For i = 1 To UBound(StEPPLink_ImportSucceeded_Name)
'      this_chemical = StEPPLink_ImportSucceeded_Name(i)
'
'      'Set all property defaults
'      'Call SetComponentDefaults(thiscomp, -1)
'      thiscomp.Influent.Value = 10#
'      thiscomp.TreatmentObjective.Value = 2#
'      xxx = StEPPLink_FilteredProps(StEPPLink_FindProp(this_chemical, "MolecularWeight"))
'      thiscomp.Name = xxx.Chemical
'      'thiscomp.Cas = CLng(xxx.Cas)
'
'      'Import each property
'      xxx = StEPPLink_FilteredProps(StEPPLink_FindProp(this_chemical, "MolecularWeight"))
'      thiscomp.MolecularWeight.Value = xxx.val
'      thiscomp.MolecularWeight.ValChanged = True
'      thiscomp.MolecularWeight.UserInput = False
'
'      xxx = StEPPLink_FilteredProps(StEPPLink_FindProp(this_chemical, "HenrysConstant"))
'      thiscomp.HenrysConstant.Value = xxx.val
'      thiscomp.HenrysConstant.ValChanged = True
'      thiscomp.HenrysConstant.UserInput = False
'
'      xxx = StEPPLink_FilteredProps(StEPPLink_FindProp(this_chemical, "NormalBoilingPoint"))
'      thiscomp.NormalBoilingPoint.Value = (xxx.val) + 273.15
'      thiscomp.NormalBoilingPoint.ValChanged = True
'      thiscomp.NormalBoilingPoint.UserInput = False
'
'      xxx = StEPPLink_FilteredProps(StEPPLink_FindProp(this_chemical, "MolarVolumeAtNBP"))
'      thiscomp.MolarVolume.Value = xxx.val
'      thiscomp.MolarVolume.ValChanged = True
'      thiscomp.MolarVolume.UserInput = False
'
'      xxx = StEPPLink_FilteredProps(StEPPLink_FindProp(this_chemical, "LiquidDiffusivity"))
'      thiscomp.LiquidDiffusivity.Value = xxx.val
'      thiscomp.LiquidDiffusivity.ValChanged = True
'      thiscomp.LiquidDiffusivity.UserInput = False
'      'thiscomp.Refractive_Index = xxx.val
'
'      xxx = StEPPLink_FilteredProps(StEPPLink_FindProp(this_chemical, "GasDiffusivity"))
'      thiscomp.GasDiffusivity.Value = xxx.val
'      thiscomp.GasDiffusivity.ValChanged = True
'      thiscomp.GasDiffusivity.UserInput = False
'
'      'Place imported component into memory
'      Data_frmContaminantPropertyEdit.NewNumCompo = Data_frmContaminantPropertyEdit.NewNumCompo + 1
'      Data_frmContaminantPropertyEdit.Contaminants(Data_frmContaminantPropertyEdit.NewNumCompo) = thiscomp
'    Next i
'
'    '-- Display standard "Import Succeeded" dialog box
'
'    StEPPLink_DontForget = "Remember to set the correct values of influent and treatment objective concentrations for these components."
'    Call StEPPLink_DisplayImportSucceeded
'
'    Unload Me
'  End If
'
End Sub

Private Sub Form_Activate()
  Call CenterThisForm(Me)
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim PositionLeft As Integer
Dim x As rec_frmContaminantPropertyEdit
Dim xu As rec_Units_frmContaminantPropertyEdit

  Call CenterThisForm(Me)

'tells it to skip certain change events
loading = 0

' DEMO MODE STOP USER FROM ADDING CONTAMINANTS :: TACK
If DemoMode% Then
    txtContaminantName.Enabled = False
    For i% = 0 To 5
        txtContaminantProperties(i%).Enabled = False
    Next i%
'    txtconcentration(0).Enabled = False
'    txtconcentration(1).Enabled = False
    '''lblpropoperatingconditions(0).Enabled = False
    '''lblpropoperatingconditions(1).Enabled = False
End If
' END DEMO STUFF

    'frmPropContaminant.WindowState = 0
    'frmPropContaminant.Height = 7140
    ''Position the form on the screen (Centered in Right Half of It)
    'If WindowState = 0 Then
    '   'don't attempt if screen Minimized or Maximized
    '   PositionLeft = frmPTADScreen1.Left + frmPTADScreen1.Width - Screen.Width / 2
    '   PositionLeft = (PositionLeft / 2 - frmPropContaminant.Width / 2)
    '   Move (Screen.Width / 2 + PositionLeft), (Screen.Height - frmPropContaminant.Height) / 2
    'End If
   ' Call CenterForm(Me)
    cmdOK.Enabled = False
    
    x = Data_frmContaminantPropertyEdit
    Caption = "Contaminant Properties - " & x.ModelName
    x.NewNumCompo = x.OldNumCompo
    If (x.DoAdd) Then

      'New contaminant properties:
      cbochemname.Visible = False

      For i = 0 To 1
        txtConcentration(i).Text = "0.0"
      Next i
      For i = 0 To 5
        txtContaminantProperties(i).Text = "0.0"
      Next i
      txtContaminantName.Enabled = True
      x.DoEditNumber = x.OldNumCompo + 1
      'cmdStEPP.Enabled = True
      'cmdStEPP.Visible = True
      cmdImportClipboard.Enabled = True
      cmdImportClipboard.Visible = True
      cmdImport.Enabled = True
      cmdImport.Visible = True
      Frame1.Visible = True

      'Set default values for ThisContaminant.
      ThisContaminant.Name = ""
      ThisContaminant.Pressure = 0#
      ThisContaminant.Temperature = 0#
      ThisContaminant.AirWaterInterfaceConcentration = 0#
      ThisContaminant.MolecularWeight.value = 131.4
      ThisContaminant.HenrysConstant.value = 0.231
      ThisContaminant.MolarVolume.value = 0.0981
      ThisContaminant.NormalBoilingPoint.value = 360#
      ThisContaminant.LiquidDiffusivity.value = 0.000000000656956
      ThisContaminant.GasDiffusivity.value = 0.00000799877
      ThisContaminant.Influent.value = 200#
      ThisContaminant.TreatmentObjective.value = 5#
      ThisContaminant.Effluent.value = 0#
    Else
        'cmdStEPP.Visible = False
        cmdImportClipboard.Visible = False
        cmdImport.Visible = False
        Frame1.Visible = False

        For i = 1 To x.NewNumCompo
            cbochemname.AddItem x.Contaminants(i).Name
        Next
        cbochemname.ListIndex = x.DoEditNumber - 1
      
      'Display the current contaminant being edited:
      ThisContaminant = x.Contaminants(x.DoEditNumber)
    End If
    txtContaminantName.Text = ThisContaminant.Name
    txtContaminantName.Enabled = True
    txtContaminantProperties(0).Text = Format$(ThisContaminant.MolecularWeight.value, "0.00")
    txtContaminantProperties(1).Text = Format$(ThisContaminant.HenrysConstant.value, GetTheFormat(ThisContaminant.HenrysConstant.value))
    txtContaminantProperties(2).Text = Format$(ThisContaminant.MolarVolume.value, GetTheFormat(ThisContaminant.MolarVolume.value))
    txtContaminantProperties(3).Text = Format$(ThisContaminant.NormalBoilingPoint.value - 273.15, "0.0")
    txtContaminantProperties(4).Text = Format$(ThisContaminant.LiquidDiffusivity.value, GetTheFormat(ThisContaminant.LiquidDiffusivity.value))
    txtContaminantProperties(5).Text = Format$(ThisContaminant.GasDiffusivity.value, GetTheFormat(ThisContaminant.GasDiffusivity.value))
    txtConcentration(0).Text = Format$(ThisContaminant.Influent.value, GetTheFormat(ThisContaminant.Influent.value))
    txtConcentration(1).Text = Format$(ThisContaminant.TreatmentObjective.value, GetTheFormat(ThisContaminant.TreatmentObjective.value))
    
    xu = Units_frmContaminantPropertyEdit
    Call LabelsPropContaminant(UNITSTYPE_SI)
    For i = 0 To 5
      If (i <> 1) Then Call SetUnits(UnitsProp(i), xu.UnitsProp(i))
    Next i
    For i = 0 To 1
      Call SetUnits(UnitsConc(i), xu.UnitsConc(i))
    Next i

    Select Case x.ModelType
      Case MODELTYPE_PACKEDTOWER
        'Do nothing.

      Case MODELTYPE_BUBBLE
        'Hide Normal Boiling Point property.
        txtContaminantProperties(3).Visible = False
        lblContaminantProperties(4).Visible = False
        UnitsProp(3).Visible = False
        
        'Hide Gas Diffusivity property.
        txtContaminantProperties(5).Visible = False
        lblContaminantProperties(6).Visible = False
        UnitsProp(5).Visible = False
      
      Case MODELTYPE_SURFACE
        'Hide Normal Boiling Point property.
        txtContaminantProperties(3).Visible = False
        lblContaminantProperties(4).Visible = False
        UnitsProp(3).Visible = False

        'Hide Gas Diffusivity property.
        txtContaminantProperties(5).Visible = False
        lblContaminantProperties(6).Visible = False
        UnitsProp(5).Visible = False

    End Select

    Data_frmContaminantPropertyEdit = x
    loading = 1
  '
  ' DEMO SETTINGS.
  '
  Call LOCAL___Reset_DemoVersionDisablings
End Sub

Private Sub Old_StEPP_Link()
'
'    Dim response As Integer
'
'    On Error GoTo Error_DDE
'       txttestdde(0).LinkTopic = "StEPP|MainForm"
'       If DemoMode% Then txttestdde(0).LinkTopic = "DStEPP|MainForm"
'       txttestdde(0).LinkItem = "lblSelectedContaminant"
'       txttestdde(0).LinkMode = 3
'       txttestdde(0).LinkRequest
'       txttestdde(0).LinkMode = 0
'       txttestdde(1).LinkTopic = "StEPP|MainForm"
'       If DemoMode% Then txttestdde(1).LinkTopic = "DStEPP|MainForm"
'       txttestdde(1).LinkItem = "txtOperatingPressure"
'       txttestdde(1).LinkMode = 3
'       txttestdde(1).LinkRequest
'       txttestdde(1).LinkMode = 0
'       txttestdde(2).LinkTopic = "StEPP|MainForm"
'       If DemoMode% Then txttestdde(2).LinkTopic = "DStEPP|MainForm"
'       txttestdde(2).LinkItem = "txtOperatingTemperature"
'       txttestdde(2).LinkMode = 3
'       txttestdde(2).LinkRequest
'       txttestdde(2).LinkMode = 0
'
'       'Check if contaminant name, operating pressure, operating temperature in StEPP and ASAP agree.  If they don't, print appropriate warning.
'       If txtcontaminantname.Text <> "" Then
'          If UCase$(Trim$(txtcontaminantname.Text)) <> UCase$(Trim$(txttestdde(0).Text)) Then
'             response = MsgBox("Contaminant Name in StEPP = " & Trim$(txttestdde(0).Text) & ".  Contaminant Name in ASAP = " & Trim$(txtcontaminantname.Text) & "." & Chr$(13) & Chr$(13) & "Proceed with DDE Link?", MB_ICONQUESTION + MB_YESNO, "Warning")
'             If response = IDNO Then Exit Sub
'          End If
'       End If
'
'       If Abs((CDbl(frmPTADScreen1!txtOperatingPressure.Text) - CDbl(txttestdde(1).Text))) > TOLERANCE Then
'          response = MsgBox("Operating Pressure in StEPP = " & Trim$(txttestdde(1).Text) & ".  Operating Pressure in ASAP = " & Trim$(frmPTADScreen1!txtOperatingPressure.Text) & "." & Chr$(13) & Chr$(13) & "Proceed with DDE link?", MB_ICONQUESTION + MB_YESNO, "Warning")
'          If response = IDNO Then Exit Sub
'       End If
'
'       If Abs((CDbl(frmPTADScreen1!txtOperatingTemperature.Text) - CDbl(txttestdde(2).Text))) > TOLERANCE Then
'          response = MsgBox("Operating Temperature in StEPP = " & Trim$(txttestdde(2).Text) & ".  Operating Temperature in ASAP = " & Trim$(frmPTADScreen1!txtOperatingTemperature.Text) & "." & Chr$(13) & Chr$(13) & "Proceed with DDE link?", MB_ICONQUESTION + MB_YESNO, "Warning")
'          If response = IDNO Then Exit Sub
'       End If
'
'       If ListContaminantMenuOptionsIndex = 2 Then  '2 --> Add Contaminant
'          Call txtContaminantName_GotFocus
'          txtcontaminantname = Trim$(LCase$(txttestdde(0).Text))
'          Call txtContaminantName_LostFocus
'       End If
'
'       txtcontaminantproperties(0).LinkTopic = "StEPP|MainForm"
'       If DemoMode% Then txtcontaminantproperties(0).LinkTopic = "DStEPP|MainForm"
'       txtcontaminantproperties(0).LinkItem = "lblContaminantProperties(3)"
'       txtcontaminantproperties(0).LinkMode = 3
'       Call txtContaminantProperties_GotFocus(0)
'       txtcontaminantproperties(0).LinkRequest
'       txtcontaminantproperties(0).LinkMode = 0
'       Call txtContaminantProperties_LostFocus(0)
'
'       txtcontaminantproperties(1).LinkTopic = "StEPP|MainForm"
'       If DemoMode% Then txtcontaminantproperties(1).LinkTopic = "DStEPP|MainForm"
'       txtcontaminantproperties(1).LinkItem = "lblContaminantProperties(2)"
'       txtcontaminantproperties(1).LinkMode = 3
'       Call txtContaminantProperties_GotFocus(1)
'       txtcontaminantproperties(1).LinkRequest
'       txtcontaminantproperties(1).LinkMode = 0
'       Call txtContaminantProperties_LostFocus(1)
''
'       txtcontaminantproperties(2).LinkTopic = "StEPP|MainForm"
'       If DemoMode% Then txtcontaminantproperties(2).LinkTopic = "DStEPP|MainForm"
'       txtcontaminantproperties(2).LinkItem = "lblContaminantProperties(7)"
'       txtcontaminantproperties(2).LinkMode = 3
'       Call txtContaminantProperties_GotFocus(2)
'       txtcontaminantproperties(2).LinkRequest
'       txtcontaminantproperties(2).LinkMode = 0
'       Call txtContaminantProperties_LostFocus(2)
'
'       txtcontaminantproperties(3).LinkTopic = "StEPP|MainForm"
'       If DemoMode% Then txtcontaminantproperties(3).LinkTopic = "DStEPP|MainForm"
'       txtcontaminantproperties(3).LinkItem = "lblContaminantProperties(4)"
'       txtcontaminantproperties(3).LinkMode = 3
'       Call txtContaminantProperties_GotFocus(3)
'       txtcontaminantproperties(3).LinkRequest
'       txtcontaminantproperties(3).LinkMode = 0
'       Call txtContaminantProperties_LostFocus(3)
'
'       txtcontaminantproperties(4).LinkTopic = "StEPP|MainForm"
'       If DemoMode% Then txtcontaminantproperties(4).LinkTopic = "DStEPP|MainForm"
'       txtcontaminantproperties(4).LinkItem = "lblContaminantProperties(11)"
'       txtcontaminantproperties(4).LinkMode = 3
'       Call txtContaminantProperties_GotFocus(4)
'       txtcontaminantproperties(4).LinkRequest
'       txtcontaminantproperties(4).LinkMode = 0
'       Call txtContaminantProperties_LostFocus(4)
'
'       txtcontaminantproperties(5).LinkTopic = "StEPP|MainForm"
'       If DemoMode% Then txtcontaminantproperties(5).LinkTopic = "DStEPP|MainForm"
'       txtcontaminantproperties(5).LinkItem = "lblContaminantProperties(12)"
'       txtcontaminantproperties(5).LinkMode = 3
'       Call txtContaminantProperties_GotFocus(5)
'       txtcontaminantproperties(5).LinkRequest
'       txtcontaminantproperties(5).LinkMode = 0
'       Call txtContaminantProperties_LostFocus(5)
'
'        Exit Sub
'
'Error_DDE:
'   If Err = 282 Then
'      MsgBox "Error while performing DDE." & Chr$(13) & Chr$(13) & "StEPP must be running in order to perform DDE", 48, "Aeration System Analysis Program - Bubble Aeration"
'   Else
'      MsgBox "Unknown error while performing DDE with StEPP.", 48, "Aeration System Analysis Program - Bubble Aeration"
'   End If
'   Exit Sub
'
'
End Sub

Private Sub txtConcentration_Change(Index As Integer)
Dim i As Integer

  cmdOK.Enabled = True
  Call LOCAL___Reset_DemoVersionDisablings
  If (txtContaminantName.Text = "") Then cmdOK.Enabled = False
  If (cmdOK.Enabled) Then
    For i = 0 To 5
      If (i <> 3) Then
        If (Val(txtContaminantProperties(i).Text) <= 0#) Then
          If (txtContaminantProperties(i).Visible) Then
            cmdOK.Enabled = False
          End If
        End If
      End If
      If (i = 3) Then
        If (Val(txtContaminantProperties(i).Text) <= -300#) Then
          If (txtContaminantProperties(i).Visible) Then
            cmdOK.Enabled = False
          End If
        End If
      End If
    Next i
  End If
  For i = 0 To 1
    If (Val(txtConcentration(i).Text) <= 0#) Then
      cmdOK.Enabled = False
    End If
  Next i

End Sub

Private Sub txtConcentration_GotFocus(Index As Integer)
  Call GotFocus_Handle(Me, txtConcentration(Index), Temp_Text)

End Sub

Private Sub txtConcentration_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
       KeyAscii = 0
       SendKeys "{Tab}"
       Exit Sub
    End If

    Call NumberCheck(KeyAscii)

End Sub

Private Sub txtConcentration_LostFocus(Index As Integer)
Dim x As rec_frmContaminantPropertyEdit
Dim NewVal As Double
Dim IsNew As Integer
Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtConcentration(Index))) Then
     Exit Sub
   End If
   
   flag_ok = True

  x = Data_frmContaminantPropertyEdit
  IsNew = False
  
  Select Case Index
    Case 0        'Influent Conc.
      If (Unitted_LostFocus(UNITS_CONCENTRATION, txtConcentration(0), UnitsConc(0), NewVal, Temp_Text)) Then
        IsNew = True
        ThisContaminant.Influent.value = NewVal
        ThisContaminant.Influent.ValChanged = True
        ThisContaminant.Influent.UserInput = True
      End If
    Case 1        'Treatment Obj.
      If (Unitted_LostFocus(UNITS_CONCENTRATION, txtConcentration(1), UnitsConc(1), NewVal, Temp_Text)) Then
        IsNew = True
        ThisContaminant.TreatmentObjective.value = NewVal
        ThisContaminant.TreatmentObjective.ValChanged = True
        ThisContaminant.TreatmentObjective.UserInput = True
      End If
  End Select

  If (IsNew) Then
    x.Contaminants(x.DoEditNumber) = ThisContaminant
    Data_frmContaminantPropertyEdit = x
  End If
  Call LostFocus_Handle(Me, txtConcentration(Index), flag_ok)


End Sub

Private Sub txtContaminantName_Change()
Dim i As Integer
Dim x As rec_frmContaminantPropertyEdit

  txtContaminantName.Text = LTrim$(txtContaminantName.Text)
  cmdOK.Enabled = True
  Call LOCAL___Reset_DemoVersionDisablings
  If txtContaminantName.Text = "" Then cmdOK.Enabled = False
  If cmdOK.Enabled Then
    For i = 0 To 5
      If (i <> 3) Then
        If (Val(txtContaminantProperties(i).Text) <= 0#) Then
          If (txtContaminantProperties(i).Visible) Then
            cmdOK.Enabled = False
          End If
        End If
      End If
      If i = 3 Then
        If (Val(txtContaminantProperties(i).Text) <= -300#) Then
          If (txtContaminantProperties(i).Visible) Then
            cmdOK.Enabled = False
          End If
        End If
      End If
    Next i
  End If
  For i = 0 To 1
    If (Val(txtConcentration(i).Text) <= 0#) Then
      cmdOK.Enabled = False
    End If
  Next i

  If loading Then
    x = Data_frmContaminantPropertyEdit
    If (Not x.DoAdd) Then
      cbochemname.List(x.DoEditNumber - 1) = txtContaminantName.Text
    End If
'cbochemname.AddItem txtcontaminantname.Text
'cbochemname.List(x.doeditnumber - 1) = txtcontaminantname.Text
  End If

End Sub

Private Sub txtContaminantName_GotFocus()
  Call GotFocus_Handle(Me, txtContaminantName, Temp_Text)

End Sub

Private Sub txtContaminantName_KeyPress(KeyAscii As Integer)

    If (KeyAscii = 13) Then
       KeyAscii = 0
       SendKeys "{Tab}"
       Exit Sub
    End If

    If (ListContaminantMenuOptionsIndex = 1) Then   'Edit Contaminant
       KeyAscii = 0
       Exit Sub
    End If

End Sub

Private Sub txtContaminantName_LostFocus()
Dim x As rec_frmContaminantPropertyEdit
Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtContaminantName)) Then
     Exit Sub
   End If
   
   flag_ok = True

    txtContaminantName.Text = RTrim$(txtContaminantName.Text)
    If (txtContaminantName.Text <> "") Then
      'x = Data_frmContaminantPropertyEdit
      ThisContaminant.Name = txtContaminantName.Text
      'Data_frmContaminantPropertyEdit = x
    End If
  Call LostFocus_Handle(Me, txtContaminantName, flag_ok)


End Sub

Private Sub txtContaminantProperties_Change(Index As Integer)
Dim i As Integer

  cmdOK.Enabled = True
  Call LOCAL___Reset_DemoVersionDisablings
  If txtContaminantName.Text = "" Then cmdOK.Enabled = False
  If cmdOK.Enabled Then
    For i = 0 To 5
      If i <> 3 Then
        If Val(txtContaminantProperties(i).Text) <= 0# Then
          If (txtContaminantProperties(i).Visible) Then
            cmdOK.Enabled = False
          End If
        End If
      End If
      If i = 3 Then
        If Val(txtContaminantProperties(i).Text) <= -300# Then
          If (txtContaminantProperties(i).Visible) Then
            cmdOK.Enabled = False
          End If
        End If
      End If
    Next i
  End If
    
  For i = 0 To 1
    If Val(txtConcentration(i).Text) <= 0# Then
      cmdOK.Enabled = False
    End If
  Next i

End Sub

Private Sub txtContaminantProperties_GotFocus(Index As Integer)
  Call GotFocus_Handle(Me, txtContaminantProperties(Index), Temp_Text)

End Sub

Private Sub txtContaminantProperties_KeyPress(Index As Integer, KeyAscii As Integer)
  
    If KeyAscii = 13 Then
       KeyAscii = 0
       SendKeys "{Tab}"
       Exit Sub
    End If

    Call NumberCheck(KeyAscii)

End Sub

Private Sub txtContaminantProperties_LostFocus(Index As Integer)
Dim HaveNeededValues As Integer
Dim ValueChanged As Integer
Dim Dummy As Double
Dim TempKelvin As Double, NBPKelvin As Double
Dim x As rec_frmContaminantPropertyEdit
Dim NewVal As Double
Dim IsNew As Integer
Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtContaminantProperties(Index))) Then
     Exit Sub
   End If
   
   flag_ok = True

  x = Data_frmContaminantPropertyEdit
  IsNew = False
  
  Select Case Index
    Case 0        'Molecular Weight
      If (Unitted_LostFocus(UNITS_MW, txtContaminantProperties(0), UnitsProp(0), NewVal, Temp_Text)) Then
        IsNew = True
        ThisContaminant.MolecularWeight.value = NewVal
        ThisContaminant.MolecularWeight.ValChanged = True
        ThisContaminant.MolecularWeight.UserInput = True
      End If
    Case 1        'Henry's Constant
      If (NoUnits_LostFocus(txtContaminantProperties(1), NewVal, Temp_Text)) Then
        IsNew = True
        ThisContaminant.HenrysConstant.value = NewVal
        ThisContaminant.HenrysConstant.ValChanged = True
        ThisContaminant.HenrysConstant.UserInput = True
      End If
    Case 2        'Molar Volume
      If (Unitted_LostFocus(UNITS_MOLARVOLUME, txtContaminantProperties(2), UnitsProp(2), NewVal, Temp_Text)) Then
        IsNew = True
        ThisContaminant.MolarVolume.value = NewVal
        ThisContaminant.MolarVolume.ValChanged = True
        ThisContaminant.MolarVolume.UserInput = True
      End If
    Case 3        'Normal Boiling Point
      If (Unitted_LostFocus(UNITS_TEMPERATURE, txtContaminantProperties(3), UnitsProp(3), NewVal, Temp_Text)) Then
        IsNew = True
        ThisContaminant.NormalBoilingPoint.value = NewVal       '+ 273.15
        ThisContaminant.NormalBoilingPoint.ValChanged = True
        ThisContaminant.NormalBoilingPoint.UserInput = True
      End If
    Case 4        'Liquid Diffusivity
      If (Unitted_LostFocus(UNITS_DIFFUSIVITY, txtContaminantProperties(4), UnitsProp(4), NewVal, Temp_Text)) Then
        IsNew = True
        ThisContaminant.LiquidDiffusivity.value = NewVal
        ThisContaminant.LiquidDiffusivity.ValChanged = True
        ThisContaminant.LiquidDiffusivity.UserInput = True
      End If
    Case 5        'Gas Diffusivity
      If (Unitted_LostFocus(UNITS_DIFFUSIVITY, txtContaminantProperties(5), UnitsProp(5), NewVal, Temp_Text)) Then
        IsNew = True
        ThisContaminant.GasDiffusivity.value = NewVal
        ThisContaminant.GasDiffusivity.ValChanged = True
        ThisContaminant.GasDiffusivity.UserInput = True
      End If

  End Select

  If (IsNew) Then
    x.Contaminants(x.DoEditNumber) = ThisContaminant
    Data_frmContaminantPropertyEdit = x
  End If
  Call LostFocus_Handle(Me, txtContaminantProperties(Index), flag_ok)


End Sub

Private Sub txtTestDDE_GotFocus(Index As Integer)
  Call GotFocus_Handle(Me, txtTestDDE(Index), Temp_Text)

End Sub

Private Sub txtTestDDE_LostFocus(Index As Integer)
Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtTestDDE(Index))) Then
     Exit Sub
   End If
   
   flag_ok = True
  
  Call LostFocus_Handle(Me, txtTestDDE(Index), flag_ok)


End Sub

Private Sub UnitsConc_Click(Index As Integer)
Dim Dummy As Double

  Select Case Index
    Case 0            'Influent Conc.
      Dummy = ThisContaminant.Influent.value
      Call Unitted_UnitChange(UNITS_CONCENTRATION, Dummy, UnitsConc(0), txtConcentration(0))

    Case 1            'Treatment Obj.
      Dummy = ThisContaminant.TreatmentObjective.value
      Call Unitted_UnitChange(UNITS_CONCENTRATION, Dummy, UnitsConc(1), txtConcentration(1))

  End Select

End Sub

Private Sub UnitsProp_Click(Index As Integer)
Dim Dummy As Double

  Select Case Index
    Case 0            'Molecular Weight
      Dummy = ThisContaminant.MolecularWeight.value
      Call Unitted_UnitChange(UNITS_MW, Dummy, UnitsProp(0), txtContaminantProperties(0))

    Case 2            'Molar Volume
      Dummy = ThisContaminant.MolarVolume.value
      Call Unitted_UnitChange(UNITS_MOLARVOLUME, Dummy, UnitsProp(2), txtContaminantProperties(2))

    Case 3            'Normal Boiling Point
      Dummy = ThisContaminant.NormalBoilingPoint.value
      Call Unitted_UnitChange(UNITS_TEMPERATURE, Dummy, UnitsProp(3), txtContaminantProperties(3))

    Case 4            'Liquid Diffusivity
      Dummy = ThisContaminant.LiquidDiffusivity.value
      Call Unitted_UnitChange(UNITS_DIFFUSIVITY, Dummy, UnitsProp(4), txtContaminantProperties(4))

    Case 5            'Gas Diffusivity
      Dummy = ThisContaminant.GasDiffusivity.value
      Call Unitted_UnitChange(UNITS_DIFFUSIVITY, Dummy, UnitsProp(5), txtContaminantProperties(5))

  End Select

End Sub


