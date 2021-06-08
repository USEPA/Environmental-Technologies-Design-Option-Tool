VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print"
   ClientHeight    =   6210
   ClientLeft      =   1620
   ClientTop       =   1545
   ClientWidth     =   9405
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   9405
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraDestination 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Destination"
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
      Height          =   1095
      Left            =   420
      TabIndex        =   31
      Top             =   60
      Width           =   1455
      Begin VB.OptionButton optDestination 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Printer"
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
         Height          =   252
         Index           =   0
         Left            =   240
         TabIndex        =   33
         Top             =   360
         Width           =   1092
      End
      Begin VB.OptionButton optDestination 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Text File"
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
         Height          =   252
         Index           =   1
         Left            =   240
         TabIndex        =   32
         Top             =   720
         Width           =   1092
      End
   End
   Begin VB.CommandButton cmdPrint 
      Appearance      =   0  'Flat
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   0
      Left            =   7620
      TabIndex        =   30
      Top             =   180
      Width           =   1692
   End
   Begin VB.CommandButton cmdPrint 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   1
      Left            =   7620
      TabIndex        =   29
      Top             =   780
      Width           =   1692
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Properties to Print"
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
      Height          =   4572
      Left            =   420
      TabIndex        =   4
      Top             =   1380
      Width           =   8892
      Begin VB.ComboBox cboUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         TabIndex        =   34
         Top             =   720
         Width           =   3972
      End
      Begin VB.OptionButton optPrintProperties 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "All Properties"
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
         Height          =   252
         Index           =   0
         Left            =   240
         TabIndex        =   26
         Top             =   360
         Width           =   2532
      End
      Begin VB.OptionButton optPrintProperties 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Selected Properties"
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
         Height          =   252
         Index           =   1
         Left            =   240
         TabIndex        =   25
         Top             =   720
         Width           =   2532
      End
      Begin VB.CheckBox chkProperties 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Vapor Pressure"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   0
         Left            =   720
         TabIndex        =   24
         Top             =   1800
         Width           =   2652
      End
      Begin VB.CheckBox chkProperties 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Activity Coefficient"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   1
         Left            =   720
         TabIndex        =   23
         Top             =   2160
         Width           =   2652
      End
      Begin VB.CheckBox chkProperties 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Henry's Constant"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   2
         Left            =   720
         TabIndex        =   22
         Top             =   2520
         Width           =   2652
      End
      Begin VB.CheckBox chkProperties 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Molecular Weight"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   3
         Left            =   720
         TabIndex        =   21
         Top             =   2880
         Width           =   2652
      End
      Begin VB.CheckBox chkProperties 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Normal Boiling Point (NBP)"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   4
         Left            =   720
         TabIndex        =   20
         Top             =   3240
         Width           =   2652
      End
      Begin VB.CheckBox chkProperties 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Liquid Density"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   5
         Left            =   720
         TabIndex        =   19
         Top             =   3600
         Width           =   2652
      End
      Begin VB.CheckBox chkProperties 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Molar Volume @ Operating T"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   720
         TabIndex        =   18
         Top             =   3960
         Width           =   2775
      End
      Begin VB.CheckBox chkProperties 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Molar Volume @ NBP"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   7
         Left            =   3600
         TabIndex        =   17
         Top             =   1800
         Width           =   2412
      End
      Begin VB.CheckBox chkProperties 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Refractive Index"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   8
         Left            =   3600
         TabIndex        =   16
         Top             =   2160
         Width           =   2412
      End
      Begin VB.CheckBox chkProperties 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Aqueous Solubility"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   9
         Left            =   3600
         TabIndex        =   15
         Top             =   2520
         Width           =   2412
      End
      Begin VB.CheckBox chkProperties 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "log Oct. Water Part. Coeff."
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   10
         Left            =   3600
         TabIndex        =   14
         Top             =   2880
         Width           =   2412
      End
      Begin VB.CheckBox chkProperties 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Liquid Diffusivity"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   11
         Left            =   3600
         TabIndex        =   13
         Top             =   3240
         Width           =   2412
      End
      Begin VB.CheckBox chkProperties 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Gas Diffusivity"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   12
         Left            =   3600
         TabIndex        =   12
         Top             =   3600
         Width           =   2412
      End
      Begin VB.CheckBox chkProperties 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Water Density"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   13
         Left            =   6240
         TabIndex        =   11
         Top             =   1800
         Width           =   2412
      End
      Begin VB.CheckBox chkProperties 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Water Viscosity"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   14
         Left            =   6240
         TabIndex        =   10
         Top             =   2160
         Width           =   2412
      End
      Begin VB.CheckBox chkProperties 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Water Surface Tension"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   15
         Left            =   6240
         TabIndex        =   9
         Top             =   2520
         Width           =   2412
      End
      Begin VB.CheckBox chkProperties 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Air Density"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   16
         Left            =   6240
         TabIndex        =   8
         Top             =   2880
         Width           =   2412
      End
      Begin VB.CheckBox chkProperties 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Air Viscosity"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   17
         Left            =   6240
         TabIndex        =   7
         Top             =   3240
         Width           =   2412
      End
      Begin VB.ComboBox cboPropertyDescription 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         TabIndex        =   6
         Top             =   270
         Width           =   3972
      End
      Begin VB.CheckBox chkProperties 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Errors/Warnings"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   18
         Left            =   3600
         TabIndex        =   5
         Top             =   3960
         Width           =   2412
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Contaminant Properties"
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
         Height          =   252
         Left            =   720
         TabIndex        =   28
         Top             =   1320
         Width           =   5292
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Properties of Air and Water"
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
         Height          =   252
         Left            =   6240
         TabIndex        =   27
         Top             =   1320
         Width           =   2412
      End
      Begin VB.Shape Shape1 
         Height          =   3135
         Left            =   600
         Top             =   1200
         Width           =   5535
      End
      Begin VB.Shape Shape2 
         Height          =   3135
         Left            =   6120
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Line Line1 
         X1              =   600
         X2              =   8760
         Y1              =   1680
         Y2              =   1680
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Selected Contaminants"
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
      Height          =   1095
      Left            =   1980
      TabIndex        =   0
      Top             =   60
      Width           =   5535
      Begin VB.OptionButton optPrintContaminants 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "All of them"
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
         Height          =   252
         Index           =   0
         Left            =   80
         TabIndex        =   2
         Top             =   360
         Width           =   1550
      End
      Begin VB.OptionButton optPrintContaminants 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Current One"
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
         Height          =   252
         Index           =   1
         Left            =   80
         TabIndex        =   1
         Top             =   720
         Width           =   1550
      End
      Begin VB.Label lblCurrentContaminant 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   1680
         TabIndex        =   3
         Top             =   720
         Width           =   3735
      End
   End
   Begin MSComDlg.CommonDialog CMDialog1 
      Left            =   60
      Top             =   3420
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontSize        =   0
      MaxFileSize     =   256
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'TAB SETTINGS FOR USE IN PRINTING
Const TABVALUE = 35
Const VALUELENGTH = 13
Const TABUNITS = TABVALUE + VALUELENGTH + 2
Const TABSOURCE = TABUNITS + 12
Const TABFULLSOURCE = 5
Const TABFULLVALUE = TABFULLSOURCE + 30
Const TABFULLUNITS = TABFULLVALUE + VALUELENGTH + 2
Const TABFULLTEMPERATURE = TABFULLUNITS + 12
Const TEMPLENGTH = 4
Const TABFULLTEMPUNITS = TABFULLTEMPERATURE + TEMPLENGTH + 1
Const TABFULLCODE = TABFULLTEMPUNITS + 4
Const TABCODE = TABFULLSOURCE
Const CODELENGTH = 6
Const TABCODEDESCRIPTION = TABCODE + CODELENGTH + 2

'Heights of Print Text for Use in Printing Full Description of Properties to Printer
Dim HeightTitle As Integer
Dim HeightOperatingConditions As Integer
Dim HeightVaporPressure As Integer
Dim HeightActivityCoefficient As Integer
Dim HeightHenrysConstant As Integer
Dim HeightMolecularWeight As Integer
Dim HeightBoilingPoint As Integer
Dim HeightLiquidDensity As Integer
Dim HeightMolarVolumeOpT As Integer
Dim HeightMolarVolumeNBP As Integer
Dim HeightRefractiveIndex As Integer
Dim HeightAqueousSolubility As Integer
Dim HeightOctWaterPartCoeff As Integer
Dim HeightLiquidDiffusivity As Integer
Dim HeightGasDiffusivity As Integer
Dim HeightWaterDensity As Integer
Dim HeightWaterViscosity As Integer
Dim HeightWaterSurfaceTension As Integer
Dim HeightAirDensity As Integer
Dim HeightAirViscosity As Integer

Dim TotalHeightThisPage As Integer   'Total Print Height on a Page So Far

Dim PrintMsg As String   'Variable to store a message used to determine height

'Number of Lines Needed to Fully Print Each Property (used to determine print height)

Const NUMLINES_VAPOR_PRESSURE = 6
Const NUMLINES_ACTIVITY_COEFFICIENT = 5
Const NUMLINES_HENRYS_CONSTANT = 10   'Note:  does not include more than 1 value from database or UNIFAC at Database T as these numbers can vary and must be accounted for separately
Const NUMLINES_MOLECULAR_WEIGHT = 7
Const NUMLINES_BOILING_POINT = 6
Const NUMLINES_LIQUID_DENSITY = 7
Const NUMLINES_MOLAR_VOLUME_OPT = 7
Const NUMLINES_MOLAR_VOLUME_NBP = 6
Const NUMLINES_REFRACTIVE_INDEX = 6
Const NUMLINES_AQUEOUS_SOLUBILITY = 9
Const NUMLINES_OCT_WATER_PART_COEFF = 8
Const NUMLINES_LIQUID_DIFFUSIVITY = 8
Const NUMLINES_GAS_DIFFUSIVITY = 6
Const NUMLINES_WATER_DENSITY = 6
Const NUMLINES_WATER_VISCOSITY = 6
Const NUMLINES_WATER_SURFACE_TENSION = 6
Const NUMLINES_AIR_DENSITY = 6
Const NUMLINES_AIR_VISCOSITY = 6

Const NUMLINES_PROPERTY_NAME = 2

Const BOTTOM_MARGIN_SAFETY_FACTOR = 1440    'Bottom margin will be at least this big when printing (1440 twips = 1 inch)

Dim HeightOneLinePropertyName As Integer 'Height of one line when fully printing name of property to printer
Dim HeightOneLinePropertyValues          'Height of one line when fully printing the values for a property to printer

Dim PrintFileName As String

Private Sub cboPropertyDescription_Click()
    If optPrintProperties(1).Value = True Then
       If frmPrint!cboPropertyDescription.ListIndex = 1 Then
          chkProperties(18).Enabled = True
       Else
          chkProperties(18).Enabled = False
       End If
    End If

End Sub

Private Sub cboUnits_Click()
'    Dim msg As String

'    If cboUnits.ListIndex = 1 Then
'       msg = "The ability to print values in English Units has not been implemented yet.  For now, it is only possible to print results in SI units."
'       MsgBox msg, MB_ICONSTOP, "Routine Not Available"
'       cboUnits.ListIndex = 0
'    End If

End Sub

Private Sub chkProperties_Click(Index As Integer)
    If chkProperties(Index).Value Then
       chkProperties(Index).BackColor = &H800000
       chkProperties(Index).ForeColor = &HFFFFFF
    Else
       chkProperties(Index).BackColor = &HC0C0C0
       chkProperties(Index).ForeColor = &H80000008
    End If
End Sub

Private Sub cmdPrint_Click(Index As Integer)
    Dim ChosenAtLeastOneValue As Integer
    Dim i As Integer, msg As String

    Select Case Index
       Case 0   'Print

          If frmPrint!optPrintProperties(1).Value Then  'If printing chosen properties, find out if at least one property is chosen
             ChosenAtLeastOneValue = False
             For i = 0 To 17
                 If (frmPrint!chkProperties(i).Value = 1) Then
                    ChosenAtLeastOneValue = True
                    Exit For
                 End If
             Next i

             If Not ChosenAtLeastOneValue Then
                msg = "In order to print chosen properties, you must "
                msg = msg + "choose at least one property" & Chr$(13)
                MsgBox msg, MB_ICONSTOP, "Error"
                Exit Sub
             End If
          End If

          Screen.MousePointer = 11   'Hourglass
          
          If cboUnits.ListIndex = 0 Then
             Call CreateUnitsArraySI
          Else
             Call CreateUnitsArrayEnglish
          End If

          If frmPrint!optDestination(0).Value Then   'Print to printer
             Call PrintToPrinter
          ElseIf frmPrint!optDestination(1).Value Then   'Print to file
             Call GetPrintFileName
             If PrintFileName$ = "" Then Exit Sub

             Open PrintFileName$ For Output As #1
             Call PrintToFile
             Close #1
          End If
          frmPrint.Hide
          
          If CurrentUnits = EnglishUnits Then Call CreateUnitsArrayEnglish   'Return to English units array if we changed this to accommodate printing.  This statement can be removed when we create the ability to print the English units.
           '******************
          ChDir App.Path

          Screen.MousePointer = 0    'Arrow
       Case 1   'Cancel
          frmPrint.Hide
     End Select
End Sub

Private Sub Form_Load()

    'Center the form on the screen
    If WindowState = 0 Then
       'don't attempt if screen Minimized or Maximized
      Move contam_prop_form.Left + (contam_prop_form.Width / 2) - (frmPrint.Width / 2), contam_prop_form.Top + (contam_prop_form.Height / 2) - (frmPrint.Height / 2)
    End If

End Sub

Private Sub FullyPrintActivityCoefficientToFile()
    Dim ValueString As String
    Dim TempString As String
    Dim Header As Integer

On Error GoTo error_ActivityCoefficient
    
    'Set header flag
    Header = 1
    'Print Activity Coefficient from UNIFAC
    If PROPAVAILABLE(ACTIVITY_COEFFICIENT_UNIFAC) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.ActivityCoefficient.UNIFAC.Value, GetTheFormat(phprop.ActivityCoefficient.UNIFAC.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.ActivityCoefficient.UNIFAC.temperature, GetTheFormat(phprop.ActivityCoefficient.UNIFAC.temperature))
       Print #1, Tab(TABFULLSOURCE); "UNIFAC AT Operating T"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(ACTIVITY_COEFFICIENT); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE); Tab(TABFULLCODE); phprop.ActivityCoefficient.BinaryInteractionParameterDatabase;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.ActivityCoefficient.UNIFAC.error > 0 Then
             Print #1, ","; phprop.ActivityCoefficient.UNIFAC.error
          Else
             Print #1, ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Print #1, Tab(TABFULLSOURCE); "UNIFAC AT Operating T"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLCODE);
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.ActivityCoefficient.UNIFAC.error < 0 Then
             Print #1, phprop.ActivityCoefficient.UNIFAC.error
          Else
             Print #1, ""
          End If
       End If
    End If
    'Print relevant codes for this property
    If phprop.ActivityCoefficient.BinaryInteractionParameterDatabase > 0 Then
       Call PrintTheCodesToFile(phprop.ActivityCoefficient.BinaryInteractionParameterDatabase)
       Header = 0
    End If
    If optPrintProperties(1).Value = True Then
       If frmPrint!chkProperties(18).Value = 1 Then
          Call PrintTheErrorsToFile(Header, phprop.ActivityCoefficient.UNIFAC.error)
       End If
    Else
       Call PrintTheErrorsToFile(Header, phprop.ActivityCoefficient.UNIFAC.error)
    End If

resume_exit:
Exit Sub

error_ActivityCoefficient:
Resume resume_exit
End Sub

Private Sub FullyPrintActivityCoefficientToPrinter()
    Dim ValueString As String
    Dim TempString As String
    Dim Header As Integer

On Error GoTo error_printactivitycoefficient
    'Set header flag
    Header = 1
    'Print Activity Coefficient from UNIFAC
    If PROPAVAILABLE(ACTIVITY_COEFFICIENT_UNIFAC) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.ActivityCoefficient.UNIFAC.Value, GetTheFormat(phprop.ActivityCoefficient.UNIFAC.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.ActivityCoefficient.UNIFAC.temperature, GetTheFormat(phprop.ActivityCoefficient.UNIFAC.temperature))
       Printer.Print Tab(TABFULLSOURCE); "UNIFAC AT Operating T"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(ACTIVITY_COEFFICIENT); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE);
       Printer.Print Tab(TABFULLCODE); phprop.ActivityCoefficient.BinaryInteractionParameterDatabase;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.ActivityCoefficient.UNIFAC.error > 0 Then
             Printer.Print ","; phprop.ActivityCoefficient.UNIFAC.error
          Else
             Printer.Print
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Printer.Print Tab(TABFULLSOURCE); "UNIFAC AT Operating T"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLCODE);
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.ActivityCoefficient.UNIFAC.error < 0 Then
             Printer.Print phprop.ActivityCoefficient.UNIFAC.error
          Else
             Printer.Print
          End If
       End If
    End If
    'Print relevant codes for this property
    If phprop.ActivityCoefficient.BinaryInteractionParameterDatabase > 0 Then
       Call PrintTheCodes(phprop.ActivityCoefficient.BinaryInteractionParameterDatabase)
       Header = 0
    End If
    If optPrintProperties(1).Value = True Then
       If frmPrint!chkProperties(18).Value = 1 Then
          Call PrintTheErrors(Header, phprop.ActivityCoefficient.UNIFAC.error)
       End If
    Else
       Call PrintTheErrors(Header, phprop.ActivityCoefficient.UNIFAC.error)
    End If

resume_exit1:
Exit Sub

error_printactivitycoefficient:
Resume resume_exit1

End Sub

Private Sub FullyPrintAirDensityToFile()
    Dim ValueString As String
    Dim TempString As String

On Error GoTo error_airdensity

    'Print Air Density from Correlation
    If PROPAVAILABLE(AIR_DENSITY_CORRELATION) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.AirDensity.correlation.Value, GetTheFormat(phprop.AirDensity.correlation.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.AirDensity.correlation.temperature, GetTheFormat(phprop.AirDensity.correlation.temperature))
       Print #1, Tab(TABFULLSOURCE); GetSource(phprop.AirDensity.correlation.Source.short); Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(AIR_DENSITY); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE);
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.AirDensity.correlation.error > 0 Then
             Print #1, Tab(TABFULLCODE); phprop.AirDensity.correlation.error
          Else
             Print #1, Tab(TABFULLCODE); ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Print #1, Tab(TABFULLSOURCE); GetSource(phprop.AirDensity.correlation.Source.short); Tab(TABFULLVALUE); ValueString;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.AirDensity.correlation.error < 0 Then
             Print #1, Tab(TABFULLCODE); phprop.AirDensity.correlation.error
          Else
             Print #1, Tab(TABFULLCODE); ""
          End If
       End If
    End If
    'Print Air Density from User Input
    If PROPAVAILABLE(AIR_DENSITY_INPUT) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.AirDensity.input.Value, GetTheFormat(phprop.AirDensity.input.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.AirDensity.input.temperature, GetTheFormat(phprop.AirDensity.input.temperature))
       Print #1, Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(AIR_DENSITY); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE)
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Print #1, Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString
    End If
    'Print the Errors/Warnings for this property
    If optPrintProperties(1).Value = True Then
       If frmPrint!chkProperties(18).Value = 1 Then
          Call PrintTheErrorsToFile(1, phprop.AirDensity.correlation.error)
       End If
    Else
       Call PrintTheErrorsToFile(1, phprop.AirDensity.correlation.error)
    End If

resume_exit2:
Exit Sub

error_airdensity:
Resume resume_exit2

End Sub

Private Sub FullyPrintAirDensityToPrinter()
    Dim ValueString As String
    Dim TempString As String

On Error GoTo error_printairdensity

    'Print Air Density from Correlation
    If PROPAVAILABLE(AIR_DENSITY_CORRELATION) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.AirDensity.correlation.Value, GetTheFormat(phprop.AirDensity.correlation.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.AirDensity.correlation.temperature, GetTheFormat(phprop.AirDensity.correlation.temperature))
       Printer.Print Tab(TABFULLSOURCE); GetSource(phprop.AirDensity.correlation.Source.short); Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(AIR_DENSITY); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE);
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.AirDensity.correlation.error > 0 Then
             Printer.Print Tab(TABFULLCODE); phprop.AirDensity.correlation.error
          Else
             Printer.Print Tab(TABFULLCODE); ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Printer.Print Tab(TABFULLSOURCE); GetSource(phprop.AirDensity.correlation.Source.short); Tab(TABFULLVALUE); ValueString;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.AirDensity.correlation.error < 0 Then
             Printer.Print Tab(TABFULLCODE); phprop.AirDensity.correlation.error
          Else
             Printer.Print Tab(TABFULLCODE); ""
          End If
       End If
    End If
    'Print Air Density from User Input
    If PROPAVAILABLE(AIR_DENSITY_INPUT) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.AirDensity.input.Value, GetTheFormat(phprop.AirDensity.input.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.AirDensity.input.temperature, GetTheFormat(phprop.AirDensity.input.temperature))
       Printer.Print Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(AIR_DENSITY); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE)
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Printer.Print Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString
    End If
    'Print the Errors/Warnings for this property
    If optPrintProperties(1).Value = True Then
       If frmPrint!chkProperties(18).Value = 1 Then
          Call PrintTheErrors(1, phprop.AirDensity.correlation.error)
       End If
    Else
       Call PrintTheErrors(1, phprop.AirDensity.correlation.error)
    End If

resume_exit3:
Exit Sub

error_printairdensity:
Resume resume_exit3


End Sub

Private Sub FullyPrintAirViscosityToFile()
    Dim ValueString As String
    Dim TempString As String

On Error GoTo error_airviscosity

    'Print Air Viscosity from Correlation
    If PROPAVAILABLE(AIR_VISCOSITY_CORRELATION) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.AirViscosity.correlation.Value, GetTheFormat(phprop.AirViscosity.correlation.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.AirViscosity.correlation.temperature, GetTheFormat(phprop.AirViscosity.correlation.temperature))
       Print #1, Tab(TABFULLSOURCE); GetSource(phprop.AirViscosity.correlation.Source.short); Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(AIR_VISCOSITY); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE);
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.AirViscosity.correlation.error > 0 Then
             Print #1, Tab(TABFULLCODE); phprop.AirViscosity.correlation.error
          Else
             Print #1, Tab(TABFULLCODE); ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Print #1, Tab(TABFULLSOURCE); GetSource(phprop.AirViscosity.correlation.Source.short); Tab(TABFULLVALUE); ValueString;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.AirViscosity.correlation.error < 0 Then
             Print #1, Tab(TABFULLCODE); phprop.AirViscosity.correlation.error
          Else
             Print #1, Tab(TABFULLCODE); ""
          End If
       End If
    End If
    'Print Air Viscosity from User Input
    If PROPAVAILABLE(AIR_VISCOSITY_INPUT) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.AirViscosity.input.Value, GetTheFormat(phprop.AirViscosity.input.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.AirViscosity.input.temperature, GetTheFormat(phprop.AirViscosity.input.temperature))
       Print #1, Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(AIR_VISCOSITY); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE)
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Print #1, Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString
    End If
    'Print the Errors/Warnings for this property
    If optPrintProperties(1).Value = True Then
       If frmPrint!chkProperties(18).Value = 1 Then
          Call PrintTheErrorsToFile(1, phprop.AirViscosity.correlation.error)
       End If
    Else
       Call PrintTheErrorsToFile(1, phprop.AirViscosity.correlation.error)
    End If

resume_exit4:
Exit Sub

error_airviscosity:
Resume resume_exit4

End Sub

Private Sub FullyPrintAirViscosityToPrinter()
    Dim ValueString As String
    Dim TempString As String

On Error GoTo error_printairviscosity

    'Print Air Viscosity from Correlation
    If PROPAVAILABLE(AIR_VISCOSITY_CORRELATION) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.AirViscosity.correlation.Value, GetTheFormat(phprop.AirViscosity.correlation.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.AirViscosity.correlation.temperature, GetTheFormat(phprop.AirViscosity.correlation.temperature))
       Printer.Print Tab(TABFULLSOURCE); GetSource(phprop.AirViscosity.correlation.Source.short); Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(AIR_VISCOSITY); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE);
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.AirViscosity.correlation.error > 0 Then
             Printer.Print Tab(TABFULLCODE); phprop.AirViscosity.correlation.error
          Else
             Printer.Print Tab(TABFULLCODE); ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Printer.Print Tab(TABFULLSOURCE); GetSource(phprop.AirViscosity.correlation.Source.short); Tab(TABFULLVALUE); ValueString;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.AirViscosity.correlation.error < 0 Then
             Printer.Print Tab(TABFULLCODE); phprop.AirViscosity.correlation.error
          Else
             Printer.Print Tab(TABFULLCODE); ""
          End If
       End If
    End If
    'Print Air Viscosity from User Input
    If PROPAVAILABLE(AIR_VISCOSITY_INPUT) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.AirViscosity.input.Value, GetTheFormat(phprop.AirViscosity.input.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.AirViscosity.input.temperature, GetTheFormat(phprop.AirViscosity.input.temperature))
       Printer.Print Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(AIR_VISCOSITY); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE)
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Printer.Print Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString
    End If
    'Print the Errors/Warnings for this property
    If optPrintProperties(1).Value = True Then
       If frmPrint!chkProperties(18).Value = 1 Then
          Call PrintTheErrors(1, phprop.AirViscosity.correlation.error)
       End If
    Else
       Call PrintTheErrors(1, phprop.AirViscosity.correlation.error)
    End If

resume_exit5:
Exit Sub

error_printairviscosity:
Resume resume_exit5

End Sub

Private Sub FullyPrintAqueousSolubilityToFile()
    Dim ValueString As String
    Dim TempString As String
    Dim Header As Integer

On Error GoTo error_aqueousSolubility

    'Set header flag
    Header = 1
    'Print Aqueous Solubility from UNIFAC Fit
    If PROPAVAILABLE(AQUEOUS_SOLUBILITY_FIT) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.AqueousSolubility.fit.UNIFAC.Value, GetTheFormat(phprop.AqueousSolubility.fit.UNIFAC.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.AqueousSolubility.fit.UNIFAC.temperature, GetTheFormat(phprop.AqueousSolubility.fit.UNIFAC.temperature))
       Print #1, Tab(TABFULLSOURCE); "UNIFAC Fit with Data Point"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(AQUEOUS_SOLUBILITY); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE); Tab(TABFULLCODE); phprop.AqueousSolubility.BinaryInteractionParameterDatabase;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.AqueousSolubility.fit.UNIFAC.error > 0 Then
             Print #1, ","; phprop.AqueousSolubility.fit.UNIFAC.error
          Else
             Print #1, ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Print #1, Tab(TABFULLSOURCE); "UNIFAC Fit with Data Point"; Tab(TABFULLVALUE); ValueString;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.AqueousSolubility.fit.UNIFAC.error < 0 Then
             Print #1, Tab(TABFULLCODE); phprop.AqueousSolubility.fit.UNIFAC.error
          Else
             Print #1, Tab(TABFULLCODE); ""
          End If
       End If
    End If

    'Print Aqueous Solubility from UNIFAC at operating temperature
    If PROPAVAILABLE(AQUEOUS_SOLUBILITY_OPT_UNIFAC) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.AqueousSolubility.operatingT.UNIFAC.Value, GetTheFormat(phprop.AqueousSolubility.operatingT.UNIFAC.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.AqueousSolubility.operatingT.UNIFAC.temperature, GetTheFormat(phprop.AqueousSolubility.operatingT.UNIFAC.temperature))
       Print #1, Tab(TABFULLSOURCE); "UNIFAC at Operating T"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(AQUEOUS_SOLUBILITY); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE); Tab(TABFULLCODE); phprop.AqueousSolubility.BinaryInteractionParameterDatabase;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.AqueousSolubility.operatingT.UNIFAC.error > 0 Then
             Print #1, ","; phprop.AqueousSolubility.operatingT.UNIFAC.error
          Else
             Print #1, ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Print #1, Tab(TABFULLSOURCE); "UNIFAC at Operating T"; Tab(TABFULLVALUE); ValueString;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.AqueousSolubility.operatingT.UNIFAC.error < 0 Then
             Print #1, Tab(TABFULLCODE); phprop.AqueousSolubility.operatingT.UNIFAC.error
          Else
             Print #1, Tab(TABFULLCODE); ""
          End If
       End If
    End If

    'Print Aqueous Solubility from Database
    If PROPAVAILABLE(AQUEOUS_SOLUBILITY_DATABASE) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.AqueousSolubility.database.Value, GetTheFormat(phprop.AqueousSolubility.database.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.AqueousSolubility.database.temperature, GetTheFormat(phprop.AqueousSolubility.database.temperature))
       Print #1, Tab(TABFULLSOURCE); "Database (" & GetSource(phprop.AqueousSolubility.database.Source.short) & ")"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(AQUEOUS_SOLUBILITY); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE);
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.AqueousSolubility.database.error > 0 Then
             Print #1, Tab(TABFULLCODE); phprop.AqueousSolubility.database.error
          Else
             Print #1, Tab(TABFULLCODE); ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Print #1, Tab(TABFULLSOURCE); "Database"; Tab(TABFULLVALUE); ValueString
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.AqueousSolubility.database.error < 0 Then
             Print #1, Tab(TABFULLCODE); phprop.AqueousSolubility.database.error
          Else
             Print #1, Tab(TABFULLCODE); ""
          End If
       End If
    End If

    'Print Aqueous Solubility from UNIFAC at database temperature
    If PROPAVAILABLE(AQUEOUS_SOLUBILITY_DBT_UNIFAC) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.AqueousSolubility.UNIFAC.Value, GetTheFormat(phprop.AqueousSolubility.UNIFAC.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.AqueousSolubility.UNIFAC.temperature, GetTheFormat(phprop.AqueousSolubility.UNIFAC.temperature))
       Print #1, Tab(TABFULLSOURCE); "UNIFAC at Database T"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(AQUEOUS_SOLUBILITY); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE); Tab(TABFULLCODE); phprop.AqueousSolubility.BinaryInteractionParameterDatabase;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.AqueousSolubility.UNIFAC.error > 0 Then
             Print #1, ","; phprop.AqueousSolubility.UNIFAC.error
          Else
             Print #1, ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Print #1, Tab(TABFULLSOURCE); "UNIFAC at Database T"; Tab(TABFULLVALUE); ValueString;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.AqueousSolubility.UNIFAC.error < 0 Then
             Print #1, Tab(TABFULLCODE); phprop.AqueousSolubility.UNIFAC.error
          Else
             Print #1, Tab(TABFULLCODE); ""
          End If
       End If
    End If

    'Print Aqueous Solubility from User Input
    If PROPAVAILABLE(AQUEOUS_SOLUBILITY_INPUT) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.AqueousSolubility.input.Value, GetTheFormat(phprop.AqueousSolubility.input.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.AqueousSolubility.input.temperature, GetTheFormat(phprop.AqueousSolubility.input.temperature))
       Print #1, Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(AQUEOUS_SOLUBILITY); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE)
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Print #1, Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString
    End If

    'Print relevant codes for this property
    If phprop.AqueousSolubility.BinaryInteractionParameterDatabase > 0 Then
       Call PrintTheCodesToFile(phprop.AqueousSolubility.BinaryInteractionParameterDatabase)
       Header = 0
    End If
    If optPrintProperties(1).Value = True Then
       If frmPrint!chkProperties(18).Value = 1 Then
          Call PrintTheErrorsToFile(Header, phprop.AqueousSolubility.fit.UNIFAC.error)
          Call PrintTheErrorsToFile(Header, phprop.AqueousSolubility.operatingT.UNIFAC.error)
          Call PrintTheErrorsToFile(Header, phprop.AqueousSolubility.database.error)
          Call PrintTheErrorsToFile(Header, phprop.AqueousSolubility.UNIFAC.error)
       End If
    Else
       Call PrintTheErrorsToFile(Header, phprop.AqueousSolubility.fit.UNIFAC.error)
       Call PrintTheErrorsToFile(Header, phprop.AqueousSolubility.operatingT.UNIFAC.error)
       Call PrintTheErrorsToFile(Header, phprop.AqueousSolubility.database.error)
       Call PrintTheErrorsToFile(Header, phprop.AqueousSolubility.UNIFAC.error)
    End If


resume_exit6:
Exit Sub

error_aqueousSolubility:
Resume resume_exit6

End Sub

Private Sub FullyPrintAqueousSolubilityToPrinter()
    Dim ValueString As String
    Dim TempString As String
    Dim Header As Integer

On Error GoTo error_printAqueousSolubility

    'Set header flag
    Header = 1
    'Print Aqueous Solubility from UNIFAC Fit
    If PROPAVAILABLE(AQUEOUS_SOLUBILITY_FIT) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.AqueousSolubility.fit.UNIFAC.Value, GetTheFormat(phprop.AqueousSolubility.fit.UNIFAC.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.AqueousSolubility.fit.UNIFAC.temperature, GetTheFormat(phprop.AqueousSolubility.fit.UNIFAC.temperature))
       Printer.Print Tab(TABFULLSOURCE); "UNIFAC Fit with Data Point"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(AQUEOUS_SOLUBILITY); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE); Tab(TABFULLCODE); phprop.AqueousSolubility.BinaryInteractionParameterDatabase;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.AqueousSolubility.fit.UNIFAC.error > 0 Then
             Printer.Print ","; phprop.AqueousSolubility.fit.UNIFAC.error
          Else
             Printer.Print ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Printer.Print Tab(TABFULLSOURCE); "UNIFAC Fit with Data Point"; Tab(TABFULLVALUE); ValueString;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.AqueousSolubility.fit.UNIFAC.error < 0 Then
             Printer.Print Tab(TABFULLCODE); phprop.AqueousSolubility.fit.UNIFAC.error
          Else
             Printer.Print Tab(TABFULLCODE); ""
          End If
       End If
    End If

    'Print Aqueous Solubility from UNIFAC at operating temperature
    If PROPAVAILABLE(AQUEOUS_SOLUBILITY_OPT_UNIFAC) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.AqueousSolubility.operatingT.UNIFAC.Value, GetTheFormat(phprop.AqueousSolubility.operatingT.UNIFAC.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.AqueousSolubility.operatingT.UNIFAC.temperature, GetTheFormat(phprop.AqueousSolubility.operatingT.UNIFAC.temperature))
       Printer.Print Tab(TABFULLSOURCE); "UNIFAC at Operating T"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(AQUEOUS_SOLUBILITY); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE); Tab(TABFULLCODE); phprop.AqueousSolubility.BinaryInteractionParameterDatabase;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.AqueousSolubility.operatingT.UNIFAC.error > 0 Then
             Printer.Print ","; phprop.AqueousSolubility.operatingT.UNIFAC.error
          Else
             Printer.Print ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Printer.Print Tab(TABFULLSOURCE); "UNIFAC at Operating T"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLCODE);
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.AqueousSolubility.operatingT.UNIFAC.error < 0 Then
             Printer.Print Tab(TABFULLCODE); phprop.AqueousSolubility.operatingT.UNIFAC.error
          Else
             Printer.Print Tab(TABFULLCODE); ""
          End If
       End If
    End If

    'Print Aqueous Solubility from Database
    If PROPAVAILABLE(AQUEOUS_SOLUBILITY_DATABASE) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.AqueousSolubility.database.Value, GetTheFormat(phprop.AqueousSolubility.database.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.AqueousSolubility.database.temperature, GetTheFormat(phprop.AqueousSolubility.database.temperature))
       Printer.Print Tab(TABFULLSOURCE); "Database (" & GetSource(phprop.AqueousSolubility.database.Source.short) & ")"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(AQUEOUS_SOLUBILITY); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE);
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.AqueousSolubility.database.error > 0 Then
             Printer.Print Tab(TABFULLCODE); phprop.AqueousSolubility.database.error
          Else
             Printer.Print Tab(TABFULLCODE); ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Printer.Print Tab(TABFULLSOURCE); "Database"; Tab(TABFULLVALUE); ValueString;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.AqueousSolubility.database.error < 0 Then
             Printer.Print Tab(TABFULLCODE); phprop.AqueousSolubility.database.error
          Else
             Printer.Print Tab(TABFULLCODE); ""
          End If
       End If
    End If

    'Print Aqueous Solubility from UNIFAC at database temperature
    If PROPAVAILABLE(AQUEOUS_SOLUBILITY_DBT_UNIFAC) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.AqueousSolubility.UNIFAC.Value, GetTheFormat(phprop.AqueousSolubility.UNIFAC.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.AqueousSolubility.UNIFAC.temperature, GetTheFormat(phprop.AqueousSolubility.UNIFAC.temperature))
       Printer.Print Tab(TABFULLSOURCE); "UNIFAC at Database T"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(AQUEOUS_SOLUBILITY); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE); Tab(TABFULLCODE); phprop.AqueousSolubility.BinaryInteractionParameterDatabase;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.AqueousSolubility.UNIFAC.error > 0 Then
             Printer.Print ","; phprop.AqueousSolubility.UNIFAC.error
          Else
             Printer.Print ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Printer.Print Tab(TABFULLSOURCE); "UNIFAC at Database T"; Tab(TABFULLVALUE); ValueString;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.AqueousSolubility.UNIFAC.error < 0 Then
             Printer.Print Tab(TABFULLCODE); phprop.AqueousSolubility.UNIFAC.error
          Else
             Printer.Print Tab(TABFULLCODE); ""
          End If
       End If
    End If

    'Print Aqueous Solubility from User Input
    If PROPAVAILABLE(AQUEOUS_SOLUBILITY_INPUT) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.AqueousSolubility.input.Value, GetTheFormat(phprop.AqueousSolubility.input.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.AqueousSolubility.input.temperature, GetTheFormat(phprop.AqueousSolubility.input.temperature))
       Printer.Print Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(AQUEOUS_SOLUBILITY); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE)
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Printer.Print Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString
    End If

    'Print relevant codes for this property
    If phprop.AqueousSolubility.BinaryInteractionParameterDatabase > 0 Then
       Call PrintTheCodes(phprop.AqueousSolubility.BinaryInteractionParameterDatabase)
       Header = 0
    End If
    If optPrintProperties(1).Value = True Then
       If frmPrint!chkProperties(18).Value = 1 Then
          Call PrintTheErrors(Header, phprop.AqueousSolubility.fit.UNIFAC.error)
          Call PrintTheErrors(0, phprop.AqueousSolubility.operatingT.UNIFAC.error)
          Call PrintTheErrors(0, phprop.AqueousSolubility.database.error)
          Call PrintTheErrors(0, phprop.AqueousSolubility.UNIFAC.error)
       End If
    Else
       Call PrintTheErrors(Header, phprop.AqueousSolubility.fit.UNIFAC.error)
       Call PrintTheErrors(0, phprop.AqueousSolubility.operatingT.UNIFAC.error)
       Call PrintTheErrors(0, phprop.AqueousSolubility.database.error)
       Call PrintTheErrors(0, phprop.AqueousSolubility.UNIFAC.error)
    End If

resume_exit7:
Exit Sub

error_printAqueousSolubility:
Resume resume_exit7


End Sub

Private Sub FullyPrintBoilingPointToFile()
    Dim ValueString As String

On Error GoTo error_boilingpoint

    'Print Normal Boiling Point From Database
    If PROPAVAILABLE(BOILING_POINT_DATABASE) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.BoilingPoint.database.Value, GetTheFormat(phprop.BoilingPoint.database.Value))
       Print #1, Tab(TABFULLSOURCE); "Database (" & GetSource(phprop.BoilingPoint.database.Source.short) & ")"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(BOILING_POINT);
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.BoilingPoint.database.error > 0 Then
             Print #1, Tab(TABFULLCODE); phprop.BoilingPoint.database.error
          Else
             Print #1, Tab(TABFULLCODE); ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Print #1, Tab(TABFULLSOURCE); "Database"; Tab(TABFULLVALUE); ValueString;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.BoilingPoint.database.error < 0 Then
             Print #1, Tab(TABFULLCODE); phprop.BoilingPoint.database.error
          Else
             Print #1, Tab(TABFULLCODE); ""
          End If
       End If
    End If

    'Print Normal Boiling Point from User Input
    If PROPAVAILABLE(BOILING_POINT_INPUT) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.BoilingPoint.input.Value, GetTheFormat(phprop.BoilingPoint.input.Value))
       Print #1, Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(BOILING_POINT)
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Print #1, Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString
    End If
    'Print relevant codes for this property
    If optPrintProperties(1).Value = True Then
       If frmPrint!chkProperties(18).Value = 1 Then
          Call PrintTheErrorsToFile(1, phprop.BoilingPoint.database.error)
       End If
    Else
       Call PrintTheErrorsToFile(1, phprop.BoilingPoint.database.error)
    End If

resume_exit8:
Exit Sub

error_boilingpoint:
Resume resume_exit8


End Sub

Private Sub FullyPrintBoilingPointToPrinter()
    Dim ValueString As String

On Error GoTo error_printBoilingPoint

    'Print Normal Boiling Point From Database
    If PROPAVAILABLE(BOILING_POINT_DATABASE) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.BoilingPoint.database.Value, GetTheFormat(phprop.BoilingPoint.database.Value))
       Printer.Print Tab(TABFULLSOURCE); "Database (" & GetSource(phprop.BoilingPoint.database.Source.short) & ")"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(BOILING_POINT);
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.BoilingPoint.database.error > 0 Then
             Printer.Print Tab(TABFULLCODE); phprop.BoilingPoint.database.error
          Else
             Printer.Print Tab(TABFULLCODE); ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Printer.Print Tab(TABFULLSOURCE); "Database"; Tab(TABFULLVALUE); ValueString;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.BoilingPoint.database.error < 0 Then
             Printer.Print Tab(TABFULLCODE); phprop.BoilingPoint.database.error
          Else
             Printer.Print Tab(TABFULLCODE); ""
          End If
       End If
    End If

    'Print Normal Boiling Point from User Input
    If PROPAVAILABLE(BOILING_POINT_INPUT) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.BoilingPoint.input.Value, GetTheFormat(phprop.BoilingPoint.input.Value))
       Printer.Print Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(BOILING_POINT)
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Printer.Print Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString
    End If
    'Print relevant codes for this property
    If optPrintProperties(1).Value = True Then
       If frmPrint!chkProperties(18).Value = 1 Then
          Call PrintTheErrors(1, phprop.BoilingPoint.database.error)
       End If
    Else
       Call PrintTheErrors(1, phprop.BoilingPoint.database.error)
    End If

resume_exit9:
Exit Sub

error_printBoilingPoint:
Resume resume_exit9


End Sub

Private Sub FullyPrintGasDiffusivityToFile()
    Dim ValueString As String
    Dim TempString As String

On Error GoTo error_gasdiffusivity

    'Print Gas Diffusivity from Wilke-Lee Modification of Hirschfelder-Bird-Spotz Method
    If PROPAVAILABLE(GAS_DIFFUSIVITY_WILKELEE) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.GasDiffusivity.wilkeLee.Value, GetTheFormat(phprop.GasDiffusivity.wilkeLee.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.GasDiffusivity.wilkeLee.temperature, GetTheFormat(phprop.GasDiffusivity.wilkeLee.temperature))
       Print #1, Tab(TABFULLSOURCE); GetSource(phprop.GasDiffusivity.wilkeLee.Source.short); Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(GAS_DIFFUSIVITY); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE);
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.GasDiffusivity.wilkeLee.error > 0 Then
             Print #1, Tab(TABFULLCODE); phprop.GasDiffusivity.wilkeLee.error
          Else
             Print #1, Tab(TABFULLCODE); ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Print #1, Tab(TABFULLSOURCE); GetSource(phprop.GasDiffusivity.wilkeLee.Source.short); Tab(TABFULLVALUE); ValueString;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.GasDiffusivity.wilkeLee.error < 0 Then
             Print #1, Tab(TABFULLCODE); phprop.GasDiffusivity.wilkeLee.error
          Else
             Print #1, Tab(TABFULLCODE); ""
          End If
       End If
    End If

    'Print Gas Diffusivity from User Input
    If PROPAVAILABLE(GAS_DIFFUSIVITY_INPUT) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.GasDiffusivity.input.Value, GetTheFormat(phprop.GasDiffusivity.input.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.GasDiffusivity.input.temperature, GetTheFormat(phprop.GasDiffusivity.input.temperature))
       Print #1, Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(GAS_DIFFUSIVITY); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE)
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Print #1, Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString
    End If
    'Print relevant codes for this property
    If optPrintProperties(1).Value = True Then
       If frmPrint!chkProperties(18).Value = 1 Then
          Call PrintTheErrorsToFile(1, phprop.GasDiffusivity.wilkeLee.error)
       End If
    Else
       Call PrintTheErrorsToFile(1, phprop.GasDiffusivity.wilkeLee.error)
    End If

resume_exit10:
Exit Sub

error_gasdiffusivity:
Resume resume_exit10

End Sub

Private Sub FullyPrintGasDiffusivityToPrinter()
    Dim ValueString As String
    Dim TempString As String

On Error GoTo error_printgasdiffusivity

    'Print Gas Diffusivity from Wilke-Lee Modification of Hirschfelder-Bird-Spotz Method
    If PROPAVAILABLE(GAS_DIFFUSIVITY_WILKELEE) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.GasDiffusivity.wilkeLee.Value, GetTheFormat(phprop.GasDiffusivity.wilkeLee.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.GasDiffusivity.wilkeLee.temperature, GetTheFormat(phprop.GasDiffusivity.wilkeLee.temperature))
       Printer.Print Tab(TABFULLSOURCE); GetSource(phprop.GasDiffusivity.wilkeLee.Source.short); Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(GAS_DIFFUSIVITY); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE);
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.GasDiffusivity.wilkeLee.error > 0 Then
             Printer.Print Tab(TABFULLCODE); phprop.GasDiffusivity.wilkeLee.error
          Else
             Printer.Print Tab(TABFULLCODE); ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Printer.Print Tab(TABFULLSOURCE); GetSource(phprop.GasDiffusivity.wilkeLee.Source.short); Tab(TABFULLVALUE); ValueString;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.GasDiffusivity.wilkeLee.error < 0 Then
             Printer.Print Tab(TABFULLCODE); phprop.GasDiffusivity.wilkeLee.error
          Else
             Printer.Print Tab(TABFULLCODE); ""
          End If
       End If
    End If

    'Print Gas Diffusivity from User Input
    If PROPAVAILABLE(GAS_DIFFUSIVITY_INPUT) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.GasDiffusivity.input.Value, GetTheFormat(phprop.GasDiffusivity.input.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.GasDiffusivity.input.temperature, GetTheFormat(phprop.GasDiffusivity.input.temperature))
       Printer.Print Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(GAS_DIFFUSIVITY); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE)
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Printer.Print Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString
    End If
    'Print relevant codes for this property
    If optPrintProperties(1).Value = True Then
       If frmPrint!chkProperties(18).Value = 1 Then
          Call PrintTheErrors(1, phprop.GasDiffusivity.wilkeLee.error)
       End If
    Else
       Call PrintTheErrors(1, phprop.GasDiffusivity.wilkeLee.error)
    End If

resume_exit11:
Exit Sub

error_printgasdiffusivity:
Resume resume_exit11

End Sub

Private Sub FullyPrintHenrysConstantToFile()
    Dim ValueString As String
    Dim TempString As String
    Dim i As Integer
    Dim Header As Integer

On Error GoTo error_henryconstant

    'Set header flag
    Header = 1
    'Print Henry's Constant From Regression of Data Points
    If PROPAVAILABLE(HENRYS_CONSTANT_REGRESS) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.HenrysConstant.regress.Value, GetTheFormat(phprop.HenrysConstant.regress.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.HenrysConstant.regress.temperature, GetTheFormat(phprop.HenrysConstant.regress.temperature))
       Print #1, Tab(TABFULLSOURCE); "Regression of Data Points"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(HENRYS_CONSTANT); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE);
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.HenrysConstant.regress.error > 0 Then
             Print #1, Tab(TABFULLCODE); phprop.HenrysConstant.regress.error
          Else
             Print #1, Tab(TABFULLCODE); ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Print #1, Tab(TABFULLSOURCE); "Regression of Data Points"; Tab(TABFULLVALUE); ValueString;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.HenrysConstant.regress.error < 0 Then
             Print #1, Tab(TABFULLCODE); phprop.HenrysConstant.regress.error
          Else
             Print #1, Tab(TABFULLCODE); ""
          End If
       End If
    End If

    'Print Henry's Constant from UNIFAC Fit
    If PROPAVAILABLE(HENRYS_CONSTANT_FIT) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.HenrysConstant.fit.UNIFAC.Value, GetTheFormat(phprop.HenrysConstant.fit.UNIFAC.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.HenrysConstant.fit.UNIFAC.temperature, GetTheFormat(phprop.HenrysConstant.fit.UNIFAC.temperature))
       Print #1, Tab(TABFULLSOURCE); "UNIFAC Fit with a Data Point"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(HENRYS_CONSTANT); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE); Tab(TABFULLCODE); phprop.ActivityCoefficient.BinaryInteractionParameterDatabase;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.HenrysConstant.fit.UNIFAC.error > 0 Then
             Print #1, ","; phprop.HenrysConstant.fit.UNIFAC.error
          Else
             Print #1, ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Print #1, Tab(TABFULLSOURCE); "UNIFAC Fit with a Data Point"; Tab(TABFULLVALUE); ValueString;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.HenrysConstant.fit.UNIFAC.error < 0 Then
             Print #1, Tab(TABFULLCODE); phprop.HenrysConstant.fit.UNIFAC.error
          Else
             Print #1, Tab(TABFULLCODE); ""
          End If
       End If
    End If

    'Print Henry's Constant from UNIFAC at Operating T
    If PROPAVAILABLE(HENRYS_CONSTANT_OPT_UNIFAC) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.HenrysConstant.operatingT.UNIFAC.Value, GetTheFormat(phprop.HenrysConstant.operatingT.UNIFAC.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.HenrysConstant.operatingT.UNIFAC.temperature, GetTheFormat(phprop.HenrysConstant.operatingT.UNIFAC.temperature))
       Print #1, Tab(TABFULLSOURCE); "UNIFAC at Operating T"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(HENRYS_CONSTANT); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE); Tab(TABFULLCODE); phprop.ActivityCoefficient.BinaryInteractionParameterDatabase;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.HenrysConstant.operatingT.UNIFAC.error > 0 Then
             Print #1, ","; phprop.HenrysConstant.operatingT.UNIFAC.error
          Else
             Print #1, ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Print #1, Tab(TABFULLSOURCE); "UNIFAC at Operating T"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLCODE);
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.HenrysConstant.operatingT.UNIFAC.error < 0 Then
             Print #1, Tab(TABFULLCODE); phprop.HenrysConstant.operatingT.UNIFAC.error
          Else
             Print #1, Tab(TABFULLCODE); ""
          End If
       End If
    End If

    'Print Henry's Constants from Database
    If PROPAVAILABLE(HENRYS_CONSTANT_DATABASE) Then
       For i = 1 To phprop.HenrysConstant.NumberOfDatabaseHenrysConstants
           ValueString = Space$(VALUELENGTH)
           RSet ValueString = Format$(phprop.HenrysConstant.database(i).Value, GetTheFormat(phprop.HenrysConstant.database(i).Value))
           TempString = Space$(TEMPLENGTH)
           RSet TempString = Format$(phprop.HenrysConstant.database(i).temperature, GetTheFormat(phprop.HenrysConstant.database(i).temperature))
           Print #1, Tab(TABFULLSOURCE); "Database (" & GetSource(phprop.HenrysConstant.database(i).Source.short) & ")"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(HENRYS_CONSTANT); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE)
       Next i
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Print #1, Tab(TABFULLSOURCE); "Database"; Tab(TABFULLVALUE); ValueString
    End If

    'Print Henry's Constants from UNIFAC corresponding to Database Temperatures
    If PROPAVAILABLE(HENRYS_CONSTANT_UNIFAC) Then
       For i = 1 To phprop.HenrysConstant.NumberOfDatabaseHenrysConstants
           If phprop.HenrysConstant.UNIFAC(i).error >= 0 Then
              ValueString = Space$(VALUELENGTH)
              RSet ValueString = Format$(phprop.HenrysConstant.UNIFAC(i).Value, GetTheFormat(phprop.HenrysConstant.UNIFAC(i).Value))
              TempString = Space$(TEMPLENGTH)
              RSet TempString = Format$(phprop.HenrysConstant.UNIFAC(i).temperature, GetTheFormat(phprop.HenrysConstant.UNIFAC(i).temperature))
              Print #1, Tab(TABFULLSOURCE); "UNIFAC at Database T"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(HENRYS_CONSTANT); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE); Tab(TABFULLCODE); phprop.ActivityCoefficient.BinaryInteractionParameterDatabase;
              If frmPrint!chkProperties(18).Value = 1 Then
                 If phprop.HenrysConstant.UNIFAC(i).error > 0 Then
                    Print #1, ","; phprop.HenrysConstant.UNIFAC(i).error
                 Else
                    Print #1, ""
                 End If
              End If
           Else
              ValueString = Space$(VALUELENGTH)
              RSet ValueString = "Not Available"
              TempString = Space$(TEMPLENGTH)
              RSet TempString = Format$(phprop.HenrysConstant.UNIFAC(i).temperature, GetTheFormat(phprop.HenrysConstant.UNIFAC(i).temperature))
              Print #1, Tab(TABFULLSOURCE); "UNIFAC at Database T"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE);
              Print #1, Tab(TABFULLCODE); phprop.HenrysConstant.UNIFAC(i).error
           End If
       Next i
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Print #1, Tab(TABFULLSOURCE); "UNIFAC at Database T"; Tab(TABFULLVALUE); ValueString
    End If


    'Print Henry's Constant from User Input
    If PROPAVAILABLE(HENRYS_CONSTANT_INPUT) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.HenrysConstant.input.Value, GetTheFormat(phprop.HenrysConstant.input.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.HenrysConstant.input.temperature, GetTheFormat(phprop.HenrysConstant.input.temperature))
       Print #1, Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(HENRYS_CONSTANT); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE)
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Print #1, Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString
    End If
    'Print relevant codes for this property
    If phprop.ActivityCoefficient.BinaryInteractionParameterDatabase > 0 Then
       Call PrintTheCodesToFile(phprop.ActivityCoefficient.BinaryInteractionParameterDatabase)
       Header = 0
    End If
    If optPrintProperties(1).Value = True Then
       If frmPrint!chkProperties(18).Value = 1 Then
          Call PrintTheErrorsToFile(Header, phprop.HenrysConstant.regress.error)
          Call PrintTheErrorsToFile(Header, phprop.HenrysConstant.fit.UNIFAC.error)
          Call PrintTheErrorsToFile(Header, phprop.HenrysConstant.operatingT.UNIFAC.error)
          For i = 1 To phprop.HenrysConstant.NumberOfDatabaseHenrysConstants
             Call PrintTheErrorsToFile(Header, phprop.HenrysConstant.UNIFAC(i).error)
          Next i
       End If
    Else
       Call PrintTheErrorsToFile(Header, phprop.HenrysConstant.regress.error)
       Call PrintTheErrorsToFile(Header, phprop.HenrysConstant.fit.UNIFAC.error)
       Call PrintTheErrorsToFile(Header, phprop.HenrysConstant.operatingT.UNIFAC.error)
       For i = 1 To phprop.HenrysConstant.NumberOfDatabaseHenrysConstants
          Call PrintTheErrorsToFile(Header, phprop.HenrysConstant.UNIFAC(i).error)
       Next i
    End If

resume_exit12:
Exit Sub

error_henryconstant:
Resume resume_exit12


End Sub

Private Sub FullyPrintHenrysConstantToPrinter()
    Dim ValueString As String
    Dim TempString As String
    Dim i As Integer
    Dim Header As Integer

On Error GoTo error_printhenryconstant

    'Set header flag
    Header = 1
    'Print Henry's Constant From Regression of Data Points
    If PROPAVAILABLE(HENRYS_CONSTANT_REGRESS) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.HenrysConstant.regress.Value, GetTheFormat(phprop.HenrysConstant.regress.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.HenrysConstant.regress.temperature, GetTheFormat(phprop.HenrysConstant.regress.temperature))
       Printer.Print Tab(TABFULLSOURCE); "Regression of Data Points"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(HENRYS_CONSTANT); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE);
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.HenrysConstant.regress.error > 0 Then
             Printer.Print Tab(TABFULLCODE); phprop.HenrysConstant.regress.error
          Else
             Printer.Print Tab(TABFULLCODE); ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Printer.Print Tab(TABFULLSOURCE); "Regression of Data Points"; Tab(TABFULLVALUE); ValueString;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.HenrysConstant.regress.error < 0 Then
             Printer.Print Tab(TABFULLCODE); phprop.HenrysConstant.regress.error
          Else
             Printer.Print Tab(TABFULLCODE); ""
          End If
       End If
    End If
    'Print Henry's Constant from UNIFAC Fit
    If PROPAVAILABLE(HENRYS_CONSTANT_FIT) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.HenrysConstant.fit.UNIFAC.Value, GetTheFormat(phprop.HenrysConstant.fit.UNIFAC.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.HenrysConstant.fit.UNIFAC.temperature, GetTheFormat(phprop.HenrysConstant.fit.UNIFAC.temperature))
       Printer.Print Tab(TABFULLSOURCE); "UNIFAC Fit with a Data Point"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(HENRYS_CONSTANT); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE); Tab(TABFULLCODE); phprop.ActivityCoefficient.BinaryInteractionParameterDatabase;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.HenrysConstant.fit.UNIFAC.error > 0 Then
             Printer.Print ","; phprop.HenrysConstant.fit.UNIFAC.error
          Else
             Printer.Print ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Printer.Print Tab(TABFULLSOURCE); "UNIFAC Fit with a Data Point"; Tab(TABFULLVALUE); ValueString;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.HenrysConstant.fit.UNIFAC.error < 0 Then
             Printer.Print Tab(TABFULLCODE); phprop.HenrysConstant.fit.UNIFAC.error
          Else
             Printer.Print ""
          End If
       End If
    End If
    'Print Henry's Constant from UNIFAC at Operating T
    If PROPAVAILABLE(HENRYS_CONSTANT_OPT_UNIFAC) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.HenrysConstant.operatingT.UNIFAC.Value, GetTheFormat(phprop.HenrysConstant.operatingT.UNIFAC.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.HenrysConstant.operatingT.UNIFAC.temperature, GetTheFormat(phprop.HenrysConstant.operatingT.UNIFAC.temperature))
       Printer.Print Tab(TABFULLSOURCE); "UNIFAC at Operating T"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(HENRYS_CONSTANT); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE); Tab(TABFULLCODE); phprop.ActivityCoefficient.BinaryInteractionParameterDatabase;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.HenrysConstant.operatingT.UNIFAC.error > 0 Then
             Printer.Print ","; phprop.HenrysConstant.operatingT.UNIFAC.error
          Else
             Printer.Print ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Printer.Print Tab(TABFULLSOURCE); "UNIFAC at Operating T"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLCODE);
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.HenrysConstant.operatingT.UNIFAC.error < 0 Then
             Printer.Print Tab(TABFULLCODE); phprop.HenrysConstant.operatingT.UNIFAC.error
          Else
             Printer.Print Tab(TABFULLCODE); ""
          End If
       End If
    End If
    'Print Henry's Constants from Database
    If PROPAVAILABLE(HENRYS_CONSTANT_DATABASE) Then
       For i = 1 To phprop.HenrysConstant.NumberOfDatabaseHenrysConstants
           ValueString = Space$(VALUELENGTH)
           RSet ValueString = Format$(phprop.HenrysConstant.database(i).Value, GetTheFormat(phprop.HenrysConstant.database(i).Value))
           TempString = Space$(TEMPLENGTH)
           RSet TempString = Format$(phprop.HenrysConstant.database(i).temperature, GetTheFormat(phprop.HenrysConstant.database(i).temperature))
           Printer.Print Tab(TABFULLSOURCE); "Database (" & GetSource(phprop.HenrysConstant.database(i).Source.short) & ")"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(HENRYS_CONSTANT); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE)
       Next i
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Printer.Print Tab(TABFULLSOURCE); "Database"; Tab(TABFULLVALUE); ValueString
    End If
    'Print Henry's Constants from UNIFAC corresponding to Database Temperatures
    If PROPAVAILABLE(HENRYS_CONSTANT_UNIFAC) Then
       For i = 1 To phprop.HenrysConstant.NumberOfDatabaseHenrysConstants
           If phprop.HenrysConstant.UNIFAC(i).error >= 0 Then
              ValueString = Space$(VALUELENGTH)
              RSet ValueString = Format$(phprop.HenrysConstant.UNIFAC(i).Value, GetTheFormat(phprop.HenrysConstant.UNIFAC(i).Value))
              TempString = Space$(TEMPLENGTH)
              RSet TempString = Format$(phprop.HenrysConstant.UNIFAC(i).temperature, GetTheFormat(phprop.HenrysConstant.UNIFAC(i).temperature))
              Printer.Print Tab(TABFULLSOURCE); "UNIFAC at Database T"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(HENRYS_CONSTANT); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE); Tab(TABFULLCODE); phprop.ActivityCoefficient.BinaryInteractionParameterDatabase;
              If frmPrint!chkProperties(18).Value = 1 Then
                 If phprop.HenrysConstant.UNIFAC(i).error > 0 Then
                    Printer.Print ","; phprop.HenrysConstant.UNIFAC(i).error
                 Else
                    Printer.Print ""
                 End If
              End If
           Else
              ValueString = Space$(VALUELENGTH)
              RSet ValueString = "Not Available"
              TempString = Space$(TEMPLENGTH)
              RSet TempString = Format$(phprop.HenrysConstant.UNIFAC(i).temperature, GetTheFormat(phprop.HenrysConstant.UNIFAC(i).temperature))
              Printer.Print Tab(TABFULLSOURCE); "UNIFAC at Database T"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE);
              Printer.Print Tab(TABFULLCODE); phprop.HenrysConstant.UNIFAC(i).error
           End If
       Next i
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Printer.Print Tab(TABFULLSOURCE); "UNIFAC at Database T"; Tab(TABFULLVALUE); ValueString
    End If
    'Print Henry's Constant from User Input
    If PROPAVAILABLE(HENRYS_CONSTANT_INPUT) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.HenrysConstant.input.Value, GetTheFormat(phprop.HenrysConstant.input.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.HenrysConstant.input.temperature, GetTheFormat(phprop.HenrysConstant.input.temperature))
       Printer.Print Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(HENRYS_CONSTANT); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE)
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Printer.Print Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString
    End If
    'Print relevant codes for this property
    If phprop.ActivityCoefficient.BinaryInteractionParameterDatabase > 0 Then
       Call PrintTheCodes(phprop.ActivityCoefficient.BinaryInteractionParameterDatabase)
       Header = 0
    End If
    If optPrintProperties(1).Value = True Then
       If frmPrint!chkProperties(18).Value = 1 Then
          Call PrintTheErrors(Header, phprop.HenrysConstant.regress.error)
          Call PrintTheErrors(0, phprop.HenrysConstant.fit.UNIFAC.error)
          Call PrintTheErrors(0, phprop.HenrysConstant.operatingT.UNIFAC.error)
          For i = 1 To phprop.HenrysConstant.NumberOfDatabaseHenrysConstants
             Call PrintTheErrors(0, phprop.HenrysConstant.UNIFAC(i).error)
          Next i
       End If
    Else
       Call PrintTheErrors(Header, phprop.HenrysConstant.regress.error)
       Call PrintTheErrors(0, phprop.HenrysConstant.fit.UNIFAC.error)
       Call PrintTheErrors(0, phprop.HenrysConstant.operatingT.UNIFAC.error)
       For i = 1 To phprop.HenrysConstant.NumberOfDatabaseHenrysConstants
          Call PrintTheErrors(0, phprop.HenrysConstant.UNIFAC(i).error)
       Next i
    End If

resume_exit13:
Exit Sub

error_printhenryconstant:
Resume resume_exit13


End Sub

Private Sub FullyPrintLiquidDensityToFile()
    Dim ValueString As String
    Dim TempString As String
    Dim Header As Integer

On Error GoTo error_liquiddensity

    Header = 1
    'Print Liquid Density From Database
    If PROPAVAILABLE(LIQUID_DENSITY_DATABASE) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.LiquidDensity.database.Value, GetTheFormat(phprop.LiquidDensity.database.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.LiquidDensity.database.temperature, GetTheFormat(phprop.LiquidDensity.database.temperature))
       Print #1, Tab(TABFULLSOURCE); "Database (" & GetSource(phprop.LiquidDensity.database.Source.short) & ")"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(LIQUID_DENSITY); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE);
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.LiquidDensity.database.error > 0 Then
             Print #1, Tab(TABFULLCODE); phprop.LiquidDensity.database.error
          Else
             Print #1, Tab(TABFULLCODE); ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Print #1, Tab(TABFULLSOURCE); "Database"; Tab(TABFULLVALUE); ValueString;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.LiquidDensity.database.error < 0 Then
             Print #1, Tab(TABFULLCODE); phprop.LiquidDensity.database.error
          Else
             Print #1, Tab(TABFULLCODE); ""
          End If
       End If
    End If

    'Print Liquid Density From Group Contribution Method
    If PROPAVAILABLE(LIQUID_DENSITY_UNIFAC) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.LiquidDensity.UNIFAC.Value, GetTheFormat(phprop.LiquidDensity.UNIFAC.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.LiquidDensity.UNIFAC.temperature, GetTheFormat(phprop.LiquidDensity.UNIFAC.temperature))
       Print #1, Tab(TABFULLSOURCE); "Group Contribution Method"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(LIQUID_DENSITY); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE);
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.LiquidDensity.UNIFAC.error > 0 Then
             Print #1, Tab(TABFULLCODE); phprop.LiquidDensity.UNIFAC.error
          Else
             Print #1, Tab(TABFULLCODE); ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Print #1, Tab(TABFULLSOURCE); "Group Contribution Method"; Tab(TABFULLVALUE); ValueString;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.LiquidDensity.UNIFAC.error < 0 Then
             Print #1, Tab(TABFULLCODE); phprop.LiquidDensity.UNIFAC.error
          Else
             Print #1, Tab(TABFULLCODE); ""
          End If
       End If
    End If

    'Print Liquid Density from User Input
    If PROPAVAILABLE(LIQUID_DENSITY_INPUT) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.LiquidDensity.input.Value, GetTheFormat(phprop.LiquidDensity.input.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.LiquidDensity.input.temperature, GetTheFormat(phprop.LiquidDensity.input.temperature))
       Print #1, Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(LIQUID_DENSITY); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE)
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Print #1, Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString
    End If
    'Print relevant codes for this property
    If optPrintProperties(1).Value = True Then
       If frmPrint!chkProperties(18).Value = 1 Then
          Call PrintTheErrorsToFile(Header, phprop.LiquidDensity.database.error)
          Call PrintTheErrorsToFile(Header, phprop.LiquidDensity.UNIFAC.error)
       End If
    Else
       Call PrintTheErrorsToFile(Header, phprop.LiquidDensity.database.error)
       Call PrintTheErrorsToFile(Header, phprop.LiquidDensity.UNIFAC.error)
    End If

resume_exit14:
Exit Sub

error_liquiddensity:
Resume resume_exit14


End Sub

Private Sub FullyPrintLiquidDensityToPrinter()
    Dim ValueString As String
    Dim TempString As String
    Dim Header As Integer

On Error GoTo error_printliquiddensity

    Header = 1
    'Print Liquid Density From Database
    If PROPAVAILABLE(LIQUID_DENSITY_DATABASE) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.LiquidDensity.database.Value, GetTheFormat(phprop.LiquidDensity.database.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.LiquidDensity.database.temperature, GetTheFormat(phprop.LiquidDensity.database.temperature))
       Printer.Print Tab(TABFULLSOURCE); "Database (" & GetSource(phprop.LiquidDensity.database.Source.short) & ")"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(LIQUID_DENSITY); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE);
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.LiquidDensity.database.error > 0 Then
             Printer.Print Tab(TABFULLCODE); phprop.LiquidDensity.database.error
          Else
             Printer.Print Tab(TABFULLCODE); ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Printer.Print Tab(TABFULLSOURCE); "Database"; Tab(TABFULLVALUE); ValueString;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.LiquidDensity.database.error < 0 Then
             Printer.Print Tab(TABFULLCODE); phprop.LiquidDensity.database.error
          Else
             Printer.Print Tab(TABFULLCODE); ""
          End If
       End If
    End If

    'Print Liquid Density From Group Contribution Method
    If PROPAVAILABLE(LIQUID_DENSITY_UNIFAC) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.LiquidDensity.UNIFAC.Value, GetTheFormat(phprop.LiquidDensity.UNIFAC.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.LiquidDensity.UNIFAC.temperature, GetTheFormat(phprop.LiquidDensity.UNIFAC.temperature))
       Printer.Print Tab(TABFULLSOURCE); "Group Contribution Method"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(LIQUID_DENSITY); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE);
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.LiquidDensity.UNIFAC.error > 0 Then
             Printer.Print Tab(TABFULLCODE); phprop.LiquidDensity.UNIFAC.error
          Else
             Printer.Print Tab(TABFULLCODE); ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Printer.Print Tab(TABFULLSOURCE); "Group Contribution Method"; Tab(TABFULLVALUE); ValueString;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.LiquidDensity.UNIFAC.error < 0 Then
             Printer.Print Tab(TABFULLCODE); phprop.LiquidDensity.UNIFAC.error
          Else
             Printer.Print Tab(TABFULLCODE); ""
          End If
       End If
    End If

    'Print Liquid Density from User Input
    If PROPAVAILABLE(LIQUID_DENSITY_INPUT) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.LiquidDensity.input.Value, GetTheFormat(phprop.LiquidDensity.input.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.LiquidDensity.input.temperature, GetTheFormat(phprop.LiquidDensity.input.temperature))
       Printer.Print Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(LIQUID_DENSITY); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE)
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Printer.Print Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString
    End If
    'Print relevant codes for this property
    If optPrintProperties(1).Value = True Then
       If frmPrint!chkProperties(18).Value = 1 Then
          Call PrintTheErrors(Header, phprop.LiquidDensity.database.error)
          Call PrintTheErrors(Header, phprop.LiquidDensity.UNIFAC.error)
       End If
    Else
       Call PrintTheErrors(Header, phprop.LiquidDensity.database.error)
       Call PrintTheErrors(Header, phprop.LiquidDensity.UNIFAC.error)
    End If

resume_exit15:
Exit Sub

error_printliquiddensity:
Resume resume_exit15


End Sub

Private Sub FullyPrintLiquidDiffusivityToFile()
    Dim ValueString As String
    Dim TempString As String
    Dim Header As Integer

On Error GoTo error_liquiddiffusivity

    Header = 1
    'Print Liquid Diffusivity from Hayduk & Laudie correlation
    If PROPAVAILABLE(LIQUID_DIFFUSIVITY_HAYDUKLAUDIE) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.LiquidDiffusivity.haydukLaudie.Value, GetTheFormat(phprop.LiquidDiffusivity.haydukLaudie.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.LiquidDiffusivity.haydukLaudie.temperature, GetTheFormat(phprop.LiquidDiffusivity.haydukLaudie.temperature))
       Print #1, Tab(TABFULLSOURCE); GetSource(phprop.LiquidDiffusivity.haydukLaudie.Source.short); Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(LIQUID_DIFFUSIVITY); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE);
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.LiquidDiffusivity.haydukLaudie.error > 0 Then
             Print #1, Tab(TABFULLCODE); phprop.LiquidDiffusivity.haydukLaudie.error
          Else
             Print #1, Tab(TABFULLCODE); ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Print #1, Tab(TABFULLSOURCE); GetSource(phprop.LiquidDiffusivity.haydukLaudie.Source.short); Tab(TABFULLVALUE); ValueString;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.LiquidDiffusivity.haydukLaudie.error < 0 Then
             Print #1, Tab(TABFULLCODE); phprop.LiquidDiffusivity.haydukLaudie.error
          Else
             Print #1, Tab(TABFULLCODE); ""
          End If
       End If
    End If

    'Print Liquid Diffusivity from method of Polson, 1950
    If PROPAVAILABLE(LIQUID_DIFFUSIVITY_POLSON) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.LiquidDiffusivity.polson.Value, GetTheFormat(phprop.LiquidDiffusivity.polson.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.LiquidDiffusivity.polson.temperature, GetTheFormat(phprop.LiquidDiffusivity.polson.temperature))
       Print #1, Tab(TABFULLSOURCE); GetSource(phprop.LiquidDiffusivity.polson.Source.short); Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(LIQUID_DIFFUSIVITY); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE);
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.LiquidDiffusivity.polson.error > 0 Then
             Print #1, Tab(TABFULLCODE); phprop.LiquidDiffusivity.polson.error
          Else
             Print #1, Tab(TABFULLCODE); ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Print #1, Tab(TABFULLSOURCE); GetSource(phprop.LiquidDiffusivity.polson.Source.short); Tab(TABFULLVALUE); ValueString;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.LiquidDiffusivity.polson.error < 0 Then
             Print #1, Tab(TABFULLCODE); phprop.LiquidDiffusivity.polson.error
          Else
             Print #1, Tab(TABFULLCODE); ""
          End If
       End If
    End If

    'Print Liquid Diffusivity from Wilke-Chang correlation
    If PROPAVAILABLE(LIQUID_DIFFUSIVITY_WILKECHANG) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.LiquidDiffusivity.wilkeChang.Value, GetTheFormat(phprop.LiquidDiffusivity.wilkeChang.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.LiquidDiffusivity.wilkeChang.temperature, GetTheFormat(phprop.LiquidDiffusivity.wilkeChang.temperature))
       Print #1, Tab(TABFULLSOURCE); GetSource(phprop.LiquidDiffusivity.wilkeChang.Source.short); Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(LIQUID_DIFFUSIVITY); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE);
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.LiquidDiffusivity.wilkeChang.error > 0 Then
             Print #1, Tab(TABFULLCODE); phprop.LiquidDiffusivity.wilkeChang.error
          Else
             Print #1, Tab(TABFULLCODE); ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Print #1, Tab(TABFULLSOURCE); GetSource(phprop.LiquidDiffusivity.wilkeChang.Source.short); Tab(TABFULLVALUE); ValueString;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.LiquidDiffusivity.wilkeChang.error < 0 Then
             Print #1, Tab(TABFULLCODE); phprop.LiquidDiffusivity.wilkeChang.error
          Else
             Print #1, Tab(TABFULLCODE); ""
          End If
       End If
    End If

    'Print Liquid Diffusivity from User Input
    If PROPAVAILABLE(LIQUID_DIFFUSIVITY_INPUT) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.LiquidDiffusivity.input.Value, GetTheFormat(phprop.LiquidDiffusivity.input.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.LiquidDiffusivity.input.temperature, GetTheFormat(phprop.LiquidDiffusivity.input.temperature))
       Print #1, Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(LIQUID_DIFFUSIVITY); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE)
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Print #1, Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString
    End If
    'Print relevant codes for this property
    If optPrintProperties(1).Value = True Then
       If frmPrint!chkProperties(18).Value = 1 Then
          Call PrintTheErrorsToFile(Header, phprop.LiquidDiffusivity.haydukLaudie.error)
          Call PrintTheErrorsToFile(Header, phprop.LiquidDiffusivity.polson.error)
          Call PrintTheErrorsToFile(Header, phprop.LiquidDiffusivity.wilkeChang.error)
       End If
    Else
       Call PrintTheErrorsToFile(Header, phprop.LiquidDiffusivity.haydukLaudie.error)
       Call PrintTheErrorsToFile(Header, phprop.LiquidDiffusivity.polson.error)
       Call PrintTheErrorsToFile(Header, phprop.LiquidDiffusivity.wilkeChang.error)
    End If

resume_exit16:
Exit Sub

error_liquiddiffusivity:
Resume resume_exit16


End Sub

Private Sub FullyPrintLiquidDiffusivityToPrinter()
    Dim ValueString As String
    Dim TempString As String
    Dim Header As Integer

On Error GoTo error_printliquiddiffusivity

    Header = 1
    'Print Liquid Diffusivity from Hayduk & Laudie correlation
    If PROPAVAILABLE(LIQUID_DIFFUSIVITY_HAYDUKLAUDIE) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.LiquidDiffusivity.haydukLaudie.Value, GetTheFormat(phprop.LiquidDiffusivity.haydukLaudie.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.LiquidDiffusivity.haydukLaudie.temperature, GetTheFormat(phprop.LiquidDiffusivity.haydukLaudie.temperature))
       Printer.Print Tab(TABFULLSOURCE); GetSource(phprop.LiquidDiffusivity.haydukLaudie.Source.short); Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(LIQUID_DIFFUSIVITY); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE);
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.LiquidDiffusivity.haydukLaudie.error > 0 Then
             Printer.Print Tab(TABFULLCODE); phprop.LiquidDiffusivity.haydukLaudie.error
          Else
             Printer.Print Tab(TABFULLCODE); ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Printer.Print Tab(TABFULLSOURCE); GetSource(phprop.LiquidDiffusivity.haydukLaudie.Source.short); Tab(TABFULLVALUE); ValueString;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.LiquidDiffusivity.haydukLaudie.error < 0 Then
             Printer.Print Tab(TABFULLCODE); phprop.LiquidDiffusivity.haydukLaudie.error
          Else
             Printer.Print Tab(TABFULLCODE); ""
          End If
       End If
    End If

    'Print Liquid Diffusivity from method of Polson, 1950
    If PROPAVAILABLE(LIQUID_DIFFUSIVITY_POLSON) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.LiquidDiffusivity.polson.Value, GetTheFormat(phprop.LiquidDiffusivity.polson.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.LiquidDiffusivity.polson.temperature, GetTheFormat(phprop.LiquidDiffusivity.polson.temperature))
       Printer.Print Tab(TABFULLSOURCE); GetSource(phprop.LiquidDiffusivity.polson.Source.short); Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(LIQUID_DIFFUSIVITY); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE);
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.LiquidDiffusivity.polson.error > 0 Then
             Printer.Print Tab(TABFULLCODE); phprop.LiquidDiffusivity.polson.error
          Else
             Printer.Print Tab(TABFULLCODE); ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Printer.Print Tab(TABFULLSOURCE); GetSource(phprop.LiquidDiffusivity.polson.Source.short); Tab(TABFULLVALUE); ValueString;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.LiquidDiffusivity.polson.error < 0 Then
             Printer.Print Tab(TABFULLCODE); phprop.LiquidDiffusivity.polson.error
          Else
             Printer.Print Tab(TABFULLCODE); ""
          End If
       End If
    End If

    'Print Liquid Diffusivity from Wilke-Chang correlation
    If PROPAVAILABLE(LIQUID_DIFFUSIVITY_WILKECHANG) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.LiquidDiffusivity.wilkeChang.Value, GetTheFormat(phprop.LiquidDiffusivity.wilkeChang.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.LiquidDiffusivity.wilkeChang.temperature, GetTheFormat(phprop.LiquidDiffusivity.wilkeChang.temperature))
       Printer.Print Tab(TABFULLSOURCE); GetSource(phprop.LiquidDiffusivity.wilkeChang.Source.short); Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(LIQUID_DIFFUSIVITY); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE);
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.LiquidDiffusivity.wilkeChang.error > 0 Then
             Printer.Print Tab(TABFULLCODE); phprop.LiquidDiffusivity.wilkeChang.error
          Else
             Printer.Print Tab(TABFULLCODE); ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Printer.Print Tab(TABFULLSOURCE); GetSource(phprop.LiquidDiffusivity.wilkeChang.Source.short); Tab(TABFULLVALUE); ValueString;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.LiquidDiffusivity.wilkeChang.error < 0 Then
             Printer.Print Tab(TABFULLCODE); phprop.LiquidDiffusivity.wilkeChang.error
          Else
             Printer.Print Tab(TABFULLCODE); ""
          End If
       End If
    End If

    'Print Liquid Diffusivity from User Input
    If PROPAVAILABLE(LIQUID_DIFFUSIVITY_INPUT) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.LiquidDiffusivity.input.Value, GetTheFormat(phprop.LiquidDiffusivity.input.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.LiquidDiffusivity.input.temperature, GetTheFormat(phprop.LiquidDiffusivity.input.temperature))
       Printer.Print Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(LIQUID_DIFFUSIVITY); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE)
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Printer.Print Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString
    End If
    'Print relevant codes for this property
    If optPrintProperties(1).Value = True Then
       If frmPrint!chkProperties(18).Value = 1 Then
          Call PrintTheErrors(Header, phprop.LiquidDiffusivity.haydukLaudie.error)
          Call PrintTheErrors(Header, phprop.LiquidDiffusivity.polson.error)
          Call PrintTheErrors(Header, phprop.LiquidDiffusivity.wilkeChang.error)
       End If
    Else
       Call PrintTheErrors(Header, phprop.LiquidDiffusivity.haydukLaudie.error)
       Call PrintTheErrors(Header, phprop.LiquidDiffusivity.polson.error)
       Call PrintTheErrors(Header, phprop.LiquidDiffusivity.wilkeChang.error)
    End If

resume_exit17:
Exit Sub

error_printliquiddiffusivity:
Resume resume_exit17


End Sub

Private Sub FullyPrintMolarVolumeNBPToFile()
    Dim ValueString As String
    Dim TempString As String

On Error GoTo error_molarvolume

    'Print Molar Volume at Normal Boiling Point From Schroeder's Method
    If PROPAVAILABLE(MOLAR_VOLUME_NBP_UNIFAC) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.MolarVolume.BoilingPoint.UNIFAC.Value, GetTheFormat(phprop.MolarVolume.BoilingPoint.UNIFAC.Value))
       TempString = Space$(TEMPLENGTH)
       If HaveProperty(BOILING_POINT) Then
          RSet TempString = Format$(phprop.MolarVolume.BoilingPoint.UNIFAC.temperature, GetTheFormat(phprop.MolarVolume.BoilingPoint.UNIFAC.temperature))
       Else
          RSet TempString = "N/A"
       End If
       Print #1, Tab(TABFULLSOURCE); GetSource(phprop.MolarVolume.BoilingPoint.UNIFAC.Source.short); Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(MOLAR_VOLUME_BOILING_POINT); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE);
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.MolarVolume.BoilingPoint.UNIFAC.error > 0 Then
             Print #1, Tab(TABFULLCODE); phprop.MolarVolume.BoilingPoint.UNIFAC.error
          Else
             Print #1, Tab(TABFULLCODE); ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Print #1, Tab(TABFULLSOURCE); GetSource(phprop.MolarVolume.BoilingPoint.UNIFAC.Source.short); Tab(TABFULLVALUE); ValueString;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.MolarVolume.BoilingPoint.UNIFAC.error < 0 Then
             Print #1, Tab(TABFULLCODE); phprop.MolarVolume.BoilingPoint.UNIFAC.error
          Else
             Print #1, Tab(TABFULLCODE); ""
          End If
       End If
    End If

    'Print Molar Volume at Normal Boiling Point from User Input
    If PROPAVAILABLE(MOLAR_VOLUME_NBP_INPUT) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.MolarVolume.BoilingPoint.input.Value, GetTheFormat(phprop.MolarVolume.BoilingPoint.input.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.MolarVolume.BoilingPoint.input.temperature, GetTheFormat(phprop.MolarVolume.BoilingPoint.input.temperature))
       Print #1, Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(MOLAR_VOLUME_BOILING_POINT); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE)
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Print #1, Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString
    End If
    'Print relevant codes for this property
    If optPrintProperties(1).Value = True Then
       If frmPrint!chkProperties(18).Value = 1 Then
          Call PrintTheErrorsToFile(1, phprop.MolarVolume.BoilingPoint.UNIFAC.error)
       End If
    Else
       Call PrintTheErrorsToFile(1, phprop.MolarVolume.BoilingPoint.UNIFAC.error)
    End If

resume_exit18:
Exit Sub

error_molarvolume:
Resume resume_exit18


End Sub

Private Sub FullyPrintMolarVolumeNBPToPrinter()
    Dim ValueString As String
    Dim TempString As String

On Error GoTo error_printmolarvolume

    'Print Molar Volume at Normal Boiling Point From Schroeder's Method
    If PROPAVAILABLE(MOLAR_VOLUME_NBP_UNIFAC) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.MolarVolume.BoilingPoint.UNIFAC.Value, GetTheFormat(phprop.MolarVolume.BoilingPoint.UNIFAC.Value))
       TempString = Space$(TEMPLENGTH)
       If HaveProperty(BOILING_POINT) Then
          RSet TempString = Format$(phprop.MolarVolume.BoilingPoint.UNIFAC.temperature, GetTheFormat(phprop.MolarVolume.BoilingPoint.UNIFAC.temperature))
       Else
          RSet TempString = "N/A"
       End If
       Printer.Print Tab(TABFULLSOURCE); GetSource(phprop.MolarVolume.BoilingPoint.UNIFAC.Source.short); Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(MOLAR_VOLUME_BOILING_POINT); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE);
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.MolarVolume.BoilingPoint.UNIFAC.error > 0 Then
             Printer.Print Tab(TABFULLCODE); phprop.MolarVolume.BoilingPoint.UNIFAC.error
          Else
             Printer.Print Tab(TABFULLCODE); ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Printer.Print Tab(TABFULLSOURCE); GetSource(phprop.MolarVolume.BoilingPoint.UNIFAC.Source.short); Tab(TABFULLVALUE); ValueString;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.MolarVolume.BoilingPoint.UNIFAC.error < 0 Then
             Printer.Print Tab(TABFULLCODE); phprop.MolarVolume.BoilingPoint.UNIFAC.error
          Else
             Printer.Print Tab(TABFULLCODE); ""
          End If
       End If
    End If

    'Print Molar Volume at Normal Boiling Point from User Input
    If PROPAVAILABLE(MOLAR_VOLUME_NBP_INPUT) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.MolarVolume.BoilingPoint.input.Value, GetTheFormat(phprop.MolarVolume.BoilingPoint.input.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.MolarVolume.BoilingPoint.input.temperature, GetTheFormat(phprop.MolarVolume.BoilingPoint.input.temperature))
       Printer.Print Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(MOLAR_VOLUME_BOILING_POINT); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE)
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Printer.Print Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString
    End If
    'Print relevant codes for this property
    If optPrintProperties(1).Value = True Then
       If frmPrint!chkProperties(18).Value = 1 Then
          Call PrintTheErrors(1, phprop.MolarVolume.BoilingPoint.UNIFAC.error)
       End If
    Else
       Call PrintTheErrors(1, phprop.MolarVolume.BoilingPoint.UNIFAC.error)
    End If

resume_exit19:
Exit Sub

error_printmolarvolume:
Resume resume_exit19


End Sub

Private Sub FullyPrintMolarVolumeOpTToFile()
    Dim ValueString As String
    Dim TempString As String
    Dim Header As Integer

On Error GoTo error_molarvolumeopt
    Header = 1
    'Print Molar Volume at Operating Temperature From Database
    If PROPAVAILABLE(MOLAR_VOLUME_OPT_DATABASE) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.MolarVolume.operatingT.database.Value, GetTheFormat(phprop.MolarVolume.operatingT.database.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.MolarVolume.operatingT.database.temperature, GetTheFormat(phprop.MolarVolume.operatingT.database.temperature))
       Print #1, Tab(TABFULLSOURCE); "Database (" & GetSource(phprop.MolarVolume.operatingT.database.Source.short) & ")"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(MOLAR_VOLUME_OPT); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE);
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.MolarVolume.operatingT.database.error > 0 Then
             Print #1, Tab(TABFULLCODE); phprop.MolarVolume.operatingT.database.error
          Else
             Print #1, Tab(TABFULLCODE); ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Print #1, Tab(TABFULLSOURCE); "Database"; Tab(TABFULLVALUE); ValueString;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.MolarVolume.operatingT.database.error < 0 Then
             Print #1, Tab(TABFULLCODE); phprop.MolarVolume.operatingT.database.error
          Else
             Print #1, Tab(TABFULLCODE); ""
          End If
       End If
    End If

    'Print Molar Volume at Operating Temperature From Group Contribution Method
    If PROPAVAILABLE(MOLAR_VOLUME_OPT_UNIFAC) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.MolarVolume.operatingT.UNIFAC.Value, GetTheFormat(phprop.MolarVolume.operatingT.UNIFAC.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.MolarVolume.operatingT.UNIFAC.temperature, GetTheFormat(phprop.MolarVolume.operatingT.UNIFAC.temperature))
       Print #1, Tab(TABFULLSOURCE); "Group Contribution Method"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(MOLAR_VOLUME_OPT); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE);
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.MolarVolume.operatingT.UNIFAC.error > 0 Then
             Print #1, Tab(TABFULLCODE); phprop.MolarVolume.operatingT.UNIFAC.error
          Else
             Print #1, Tab(TABFULLCODE); ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Print #1, Tab(TABFULLSOURCE); "Group Contribution Method"; Tab(TABFULLVALUE); ValueString;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.MolarVolume.operatingT.UNIFAC.error < 0 Then
             Print #1, Tab(TABFULLCODE); phprop.MolarVolume.operatingT.UNIFAC.error
          Else
             Print #1, Tab(TABFULLCODE); ""
          End If
       End If
    End If

    'Print Molar Volume at Operating Temperature from User Input
    If PROPAVAILABLE(MOLAR_VOLUME_OPT_INPUT) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.MolarVolume.operatingT.input.Value, GetTheFormat(phprop.MolarVolume.operatingT.input.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.MolarVolume.operatingT.input.temperature, GetTheFormat(phprop.MolarVolume.operatingT.input.temperature))
       Print #1, Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(MOLAR_VOLUME_OPT); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE)
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Print #1, Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString
    End If
    'Print relevant codes for this property
    If optPrintProperties(1).Value = True Then
       If frmPrint!chkProperties(18).Value = 1 Then
          Call PrintTheErrorsToFile(Header, phprop.MolarVolume.operatingT.database.error)
          Call PrintTheErrorsToFile(Header, phprop.MolarVolume.operatingT.UNIFAC.error)
       End If
    Else
       Call PrintTheErrorsToFile(Header, phprop.MolarVolume.operatingT.database.error)
       Call PrintTheErrorsToFile(Header, phprop.MolarVolume.operatingT.UNIFAC.error)
    End If

resume_exit20:
Exit Sub

error_molarvolumeopt:
Resume resume_exit20


End Sub

Private Sub FullyPrintMolarVolumeOpTToPrinter()
    Dim ValueString As String
    Dim TempString As String
    Dim Header As Integer

On Error GoTo error_printmolarVolumeOPt

    Header = 1
    'Print Molar Volume at Operating Temperature From Database
    If PROPAVAILABLE(MOLAR_VOLUME_OPT_DATABASE) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.MolarVolume.operatingT.database.Value, GetTheFormat(phprop.MolarVolume.operatingT.database.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.MolarVolume.operatingT.database.temperature, GetTheFormat(phprop.MolarVolume.operatingT.database.temperature))
       Printer.Print Tab(TABFULLSOURCE); "Database (" & GetSource(phprop.MolarVolume.operatingT.database.Source.short) & ")"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(MOLAR_VOLUME_OPT); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE);
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.MolarVolume.operatingT.database.error > 0 Then
             Printer.Print Tab(TABFULLCODE); phprop.MolarVolume.operatingT.database.error
          Else
             Printer.Print Tab(TABFULLCODE); ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Printer.Print Tab(TABFULLSOURCE); "Database"; Tab(TABFULLVALUE); ValueString;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.MolarVolume.operatingT.database.error < 0 Then
             Printer.Print Tab(TABFULLCODE); phprop.MolarVolume.operatingT.database.error
          Else
             Printer.Print Tab(TABFULLCODE); ""
          End If
       End If
    End If

    'Print Molar Volume at Operating Temperature From Group Contribution Method
    If PROPAVAILABLE(MOLAR_VOLUME_OPT_UNIFAC) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.MolarVolume.operatingT.UNIFAC.Value, GetTheFormat(phprop.MolarVolume.operatingT.UNIFAC.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.MolarVolume.operatingT.UNIFAC.temperature, GetTheFormat(phprop.MolarVolume.operatingT.UNIFAC.temperature))
       Printer.Print Tab(TABFULLSOURCE); "Group Contribution Method"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(MOLAR_VOLUME_OPT); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE);
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.MolarVolume.operatingT.UNIFAC.error > 0 Then
             Printer.Print Tab(TABFULLCODE); phprop.MolarVolume.operatingT.UNIFAC.error
          Else
             Printer.Print Tab(TABFULLCODE); ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Printer.Print Tab(TABFULLSOURCE); "Group Contribution Method"; Tab(TABFULLVALUE); ValueString;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.MolarVolume.operatingT.UNIFAC.error < 0 Then
             Printer.Print Tab(TABFULLCODE); phprop.MolarVolume.operatingT.UNIFAC.error
          Else
             Printer.Print Tab(TABFULLCODE); ""
          End If
       End If
    End If

    'Print Molar Volume at Operating Temperature from User Input
    If PROPAVAILABLE(MOLAR_VOLUME_OPT_INPUT) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.MolarVolume.operatingT.input.Value, GetTheFormat(phprop.MolarVolume.operatingT.input.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.MolarVolume.operatingT.input.temperature, GetTheFormat(phprop.MolarVolume.operatingT.input.temperature))
       Printer.Print Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(MOLAR_VOLUME_OPT); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE)
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Printer.Print Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString
    End If
    'Print relevant codes for this property
    If optPrintProperties(1).Value = True Then
       If frmPrint!chkProperties(18).Value = 1 Then
          Call PrintTheErrors(Header, phprop.MolarVolume.operatingT.database.error)
          Call PrintTheErrors(Header, phprop.MolarVolume.operatingT.UNIFAC.error)
       End If
    Else
       Call PrintTheErrors(Header, phprop.MolarVolume.operatingT.database.error)
       Call PrintTheErrors(Header, phprop.MolarVolume.operatingT.UNIFAC.error)
    End If

resume_exit21:
Exit Sub

error_printmolarVolumeOPt:
Resume resume_exit21

End Sub

Private Sub FullyPrintMolecularWeightToFile()
    Dim ValueString As String
    Dim Header As Integer

On Error GoTo error_molecularWeight

    Header = 1
    'Print Molecular Weight From Database
    If PROPAVAILABLE(MOLECULAR_WEIGHT_DATABASE) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.MolecularWeight.database.Value, MOLECULAR_WEIGHT_FORMAT)
       Print #1, Tab(TABFULLSOURCE); "Database (" & GetSource(phprop.MolecularWeight.database.Source.short) & ")"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(MOLECULAR_WEIGHT); "";
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.MolecularWeight.database.error > 0 Then
             Print #1, Tab(TABFULLCODE); phprop.MolecularWeight.database.error
          Else
             Print #1, Tab(TABFULLCODE); ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Print #1, Tab(TABFULLSOURCE); "Database"; Tab(TABFULLVALUE); ValueString;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.MolecularWeight.database.error < 0 Then
             Print #1, Tab(TABFULLCODE); phprop.MolecularWeight.database.error
          Else
             Print #1, Tab(TABFULLCODE); ""
          End If
       End If
    End If

    'Print Molecular Weight From Group Contribution Method
    If PROPAVAILABLE(MOLECULAR_WEIGHT_UNIFAC) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.MolecularWeight.UNIFAC.Value, MOLECULAR_WEIGHT_FORMAT)
       Print #1, Tab(TABFULLSOURCE); "Group Contribution Method"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(MOLECULAR_WEIGHT);
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.MolecularWeight.UNIFAC.error > 0 Then
             Print #1, Tab(TABFULLCODE); phprop.MolecularWeight.UNIFAC.error
          Else
             Print #1, Tab(TABFULLCODE); ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Print #1, Tab(TABFULLSOURCE); "Group Contribution Method"; Tab(TABFULLVALUE); ValueString;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.MolecularWeight.UNIFAC.error < 0 Then
             Print #1, Tab(TABFULLCODE); phprop.MolecularWeight.UNIFAC.error
          Else
             Print #1, Tab(TABFULLCODE); ""
          End If
       End If
    End If

    'Print Molecular Weight from User Input
    If PROPAVAILABLE(MOLECULAR_WEIGHT_INPUT) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.MolecularWeight.input.Value, MOLECULAR_WEIGHT_FORMAT)
       Print #1, Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(MOLECULAR_WEIGHT)
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Print #1, Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString
    End If
    'Print relevant codes for this property
    If optPrintProperties(1).Value = True Then
       If frmPrint!chkProperties(18).Value = 1 Then
          Call PrintTheErrorsToFile(Header, phprop.MolecularWeight.database.error)
          Call PrintTheErrorsToFile(Header, phprop.MolecularWeight.UNIFAC.error)
       End If
    Else
       Call PrintTheErrorsToFile(Header, phprop.MolecularWeight.database.error)
       Call PrintTheErrorsToFile(Header, phprop.MolecularWeight.UNIFAC.error)
    End If

resume_exit22:
Exit Sub

error_molecularWeight:
Resume resume_exit22


End Sub

Private Sub FullyPrintMolecularWeightToPrinter()
    Dim ValueString As String
    Dim Header As Integer

On Error GoTo error_printmolecularWeight

    Header = 1
    'Print Molecular Weight From Database
    If PROPAVAILABLE(MOLECULAR_WEIGHT_DATABASE) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.MolecularWeight.database.Value, MOLECULAR_WEIGHT_FORMAT)
       Printer.Print Tab(TABFULLSOURCE); "Database (" & GetSource(phprop.MolecularWeight.database.Source.short) & ")"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(MOLECULAR_WEIGHT); "";
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.MolecularWeight.database.error > 0 Then
             Printer.Print Tab(TABFULLCODE); phprop.MolecularWeight.database.error
          Else
             Printer.Print Tab(TABFULLCODE); ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Printer.Print Tab(TABFULLSOURCE); "Database"; Tab(TABFULLVALUE); ValueString;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.MolecularWeight.database.error > 0 Then
             Printer.Print Tab(TABFULLCODE); phprop.MolecularWeight.database.error
          Else
             Printer.Print Tab(TABFULLCODE); ""
          End If
       End If
    End If

    'Print Molecular Weight From Group Contribution Method
    If PROPAVAILABLE(MOLECULAR_WEIGHT_UNIFAC) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.MolecularWeight.UNIFAC.Value, MOLECULAR_WEIGHT_FORMAT)
       Printer.Print Tab(TABFULLSOURCE); "Group Contribution Method"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(MOLECULAR_WEIGHT);
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.MolecularWeight.UNIFAC.error > 0 Then
             Printer.Print Tab(TABFULLCODE); phprop.MolecularWeight.UNIFAC.error
          Else
             Printer.Print Tab(TABFULLCODE); ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Printer.Print Tab(TABFULLSOURCE); "Group Contribution Method"; Tab(TABFULLVALUE); ValueString;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.MolecularWeight.UNIFAC.error < 0 Then
             Printer.Print Tab(TABFULLCODE); phprop.MolecularWeight.UNIFAC.error
          Else
             Printer.Print Tab(TABFULLCODE); ""
          End If
       End If
    End If

    'Print Molecular Weight from User Input
    If PROPAVAILABLE(MOLECULAR_WEIGHT_INPUT) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.MolecularWeight.input.Value, MOLECULAR_WEIGHT_FORMAT)
       Printer.Print Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(MOLECULAR_WEIGHT)
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Printer.Print Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString
    End If
    'Print relevant codes for this property
    If optPrintProperties(1).Value = True Then
       If frmPrint!chkProperties(18).Value = 1 Then
          Call PrintTheErrors(Header, phprop.MolecularWeight.database.error)
          Call PrintTheErrors(Header, phprop.MolecularWeight.UNIFAC.error)
       End If
    Else
       Call PrintTheErrors(Header, phprop.MolecularWeight.database.error)
       Call PrintTheErrors(Header, phprop.MolecularWeight.UNIFAC.error)
    End If

resume_exit23:
Exit Sub

error_printmolecularWeight:
Resume resume_exit23


End Sub

Private Sub FullyPrintOctWaterPartCoeffToFile()
    Dim ValueString As String
    Dim TempString As String
    Dim Header As Integer

On Error GoTo error_octwater

    'Set header flag
    Header = 1
    'Print Octanol Water Partition Coefficient from UNIFAC at operating temperature
    If PROPAVAILABLE(OCT_WATER_PART_COEFF_OPT_UNIFAC) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.OctWaterPartCoeff.operatingT.UNIFAC.Value, GetTheFormat(phprop.OctWaterPartCoeff.operatingT.UNIFAC.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.OctWaterPartCoeff.operatingT.UNIFAC.temperature, GetTheFormat(phprop.OctWaterPartCoeff.operatingT.UNIFAC.temperature))
       Print #1, Tab(TABFULLSOURCE); "UNIFAC at Operating T"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(OCT_WATER_PART_COEFF); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE); Tab(TABFULLCODE); phprop.OctWaterPartCoeff.BinaryInteractionParameterDatabase;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.OctWaterPartCoeff.operatingT.UNIFAC.error > 0 Then
             Print #1, ","; phprop.OctWaterPartCoeff.operatingT.UNIFAC.error
          Else
             Print #1, ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Print #1, Tab(TABFULLSOURCE); "UNIFAC at Operating T"; Tab(TABFULLVALUE); ValueString;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.OctWaterPartCoeff.operatingT.UNIFAC.error < 0 Then
             Print #1, Tab(TABFULLCODE); phprop.OctWaterPartCoeff.operatingT.UNIFAC.error
          Else
             Print #1, Tab(TABFULLCODE); ""
          End If
       End If
    End If

    'Print Octanol Water Partition Coefficient from Database
    If PROPAVAILABLE(OCT_WATER_PART_COEFF_DB) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.OctWaterPartCoeff.database.Value, GetTheFormat(phprop.OctWaterPartCoeff.database.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.OctWaterPartCoeff.database.temperature, GetTheFormat(phprop.OctWaterPartCoeff.database.temperature))
       Print #1, Tab(TABFULLSOURCE); "Database (" & GetSource(phprop.OctWaterPartCoeff.database.Source.short) & ")"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(OCT_WATER_PART_COEFF); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE);
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.OctWaterPartCoeff.database.error > 0 Then
             Print #1, Tab(TABFULLCODE); phprop.OctWaterPartCoeff.database.error
          Else
             Print #1, Tab(TABFULLCODE); ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Print #1, Tab(TABFULLSOURCE); "Database"; Tab(TABFULLVALUE); ValueString;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.OctWaterPartCoeff.database.error < 0 Then
             Print #1, Tab(TABFULLCODE); phprop.OctWaterPartCoeff.database.error
          Else
             Print #1, Tab(TABFULLCODE); ""
          End If
       End If
    End If

    'Print Octanol Water Partition Coefficient from UNIFAC at database temperature
    If PROPAVAILABLE(OCT_WATER_PART_COEFF_DBT_UNIFAC) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.OctWaterPartCoeff.databaseT.UNIFAC.Value, GetTheFormat(phprop.OctWaterPartCoeff.databaseT.UNIFAC.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.OctWaterPartCoeff.databaseT.UNIFAC.temperature, GetTheFormat(phprop.OctWaterPartCoeff.databaseT.UNIFAC.temperature))
       Print #1, Tab(TABFULLSOURCE); "UNIFAC at Database T"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(OCT_WATER_PART_COEFF); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE); Tab(TABFULLCODE); phprop.OctWaterPartCoeff.BinaryInteractionParameterDatabase;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.OctWaterPartCoeff.databaseT.UNIFAC.error > 0 Then
             Print #1, ","; phprop.OctWaterPartCoeff.databaseT.UNIFAC.error
          Else
             Print #1, ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Print #1, Tab(TABFULLSOURCE); "UNIFAC at Database T"; Tab(TABFULLVALUE); ValueString;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.OctWaterPartCoeff.databaseT.UNIFAC.error < 0 Then
             Print #1, Tab(TABFULLCODE); phprop.OctWaterPartCoeff.databaseT.UNIFAC.error
          Else
             Print #1, Tab(TABFULLCODE); ""
          End If
       End If
    End If

    'Print Octanol Water Partition Coefficient from User Input
    If PROPAVAILABLE(OCT_WATER_PART_COEFF_INPUT) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.OctWaterPartCoeff.input.Value, GetTheFormat(phprop.OctWaterPartCoeff.input.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.OctWaterPartCoeff.input.temperature, GetTheFormat(phprop.OctWaterPartCoeff.input.temperature))
       Print #1, Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(OCT_WATER_PART_COEFF); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE)
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Print #1, Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString
    End If
    'Print relevant codes for this property
    If phprop.OctWaterPartCoeff.BinaryInteractionParameterDatabase > 0 Then
       Call PrintTheCodesToFile(phprop.OctWaterPartCoeff.BinaryInteractionParameterDatabase)
       Header = 0
    End If
    If optPrintProperties(1).Value = True Then
       If frmPrint!chkProperties(18).Value = 1 Then
          Call PrintTheErrorsToFile(Header, phprop.OctWaterPartCoeff.operatingT.UNIFAC.error)
          Call PrintTheErrorsToFile(Header, phprop.OctWaterPartCoeff.database.error)
          Call PrintTheErrorsToFile(Header, phprop.OctWaterPartCoeff.databaseT.UNIFAC.error)
       End If
    Else
       Call PrintTheErrorsToFile(Header, phprop.OctWaterPartCoeff.operatingT.UNIFAC.error)
       Call PrintTheErrorsToFile(Header, phprop.OctWaterPartCoeff.database.error)
       Call PrintTheErrorsToFile(Header, phprop.OctWaterPartCoeff.databaseT.UNIFAC.error)
    End If
resume_exit24:
Exit Sub

error_octwater:
Resume resume_exit24


End Sub

Private Sub FullyPrintOctWaterPartCoeffToPrinter()
    Dim ValueString As String
    Dim TempString As String
    Dim Header As Integer

On Error GoTo error_printoctwater

    'Set header flag
    Header = 1
    'Print Octanol Water Partition Coefficient from UNIFAC at operating temperature
    If PROPAVAILABLE(OCT_WATER_PART_COEFF_OPT_UNIFAC) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.OctWaterPartCoeff.operatingT.UNIFAC.Value, GetTheFormat(phprop.OctWaterPartCoeff.operatingT.UNIFAC.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.OctWaterPartCoeff.operatingT.UNIFAC.temperature, GetTheFormat(phprop.OctWaterPartCoeff.operatingT.UNIFAC.temperature))
       Printer.Print Tab(TABFULLSOURCE); "UNIFAC at Operating T"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(OCT_WATER_PART_COEFF); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE); Tab(TABFULLCODE); phprop.OctWaterPartCoeff.BinaryInteractionParameterDatabase;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.OctWaterPartCoeff.operatingT.UNIFAC.error > 0 Then
             Printer.Print ","; phprop.OctWaterPartCoeff.operatingT.UNIFAC.error
          Else
             Printer.Print ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Printer.Print Tab(TABFULLSOURCE); "UNIFAC at Operating T"; Tab(TABFULLVALUE); ValueString;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.OctWaterPartCoeff.operatingT.UNIFAC.error < 0 Then
             Printer.Print Tab(TABFULLCODE); phprop.OctWaterPartCoeff.operatingT.UNIFAC.error
          Else
             Printer.Print Tab(TABFULLCODE); ""
          End If
       End If
    End If

    'Print Octanol Water Partition Coefficient from Database
    If PROPAVAILABLE(OCT_WATER_PART_COEFF_DB) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.OctWaterPartCoeff.database.Value, GetTheFormat(phprop.OctWaterPartCoeff.database.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.OctWaterPartCoeff.database.temperature, GetTheFormat(phprop.OctWaterPartCoeff.database.temperature))
       Printer.Print Tab(TABFULLSOURCE); "Database (" & GetSource(phprop.OctWaterPartCoeff.database.Source.short) & ")"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(OCT_WATER_PART_COEFF); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE);
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.OctWaterPartCoeff.database.error > 0 Then
             Printer.Print Tab(TABFULLCODE); phprop.OctWaterPartCoeff.database.error
          Else
             Printer.Print Tab(TABFULLCODE); ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Printer.Print Tab(TABFULLSOURCE); "Database"; Tab(TABFULLVALUE); ValueString;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.OctWaterPartCoeff.database.error < 0 Then
             Printer.Print Tab(TABFULLCODE); phprop.OctWaterPartCoeff.database.error
          Else
             Printer.Print Tab(TABFULLCODE); ""
          End If
       End If
    End If

    'Print Octanol Water Partition Coefficient from UNIFAC at database temperature
    If PROPAVAILABLE(OCT_WATER_PART_COEFF_DBT_UNIFAC) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.OctWaterPartCoeff.databaseT.UNIFAC.Value, GetTheFormat(phprop.OctWaterPartCoeff.databaseT.UNIFAC.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.OctWaterPartCoeff.databaseT.UNIFAC.temperature, GetTheFormat(phprop.OctWaterPartCoeff.databaseT.UNIFAC.temperature))
       Printer.Print Tab(TABFULLSOURCE); "UNIFAC at Database T"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(OCT_WATER_PART_COEFF); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE); Tab(TABFULLCODE); phprop.OctWaterPartCoeff.BinaryInteractionParameterDatabase;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.OctWaterPartCoeff.databaseT.UNIFAC.error > 0 Then
             Printer.Print ","; phprop.OctWaterPartCoeff.databaseT.UNIFAC.error
          Else
             Printer.Print ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Printer.Print Tab(TABFULLSOURCE); "UNIFAC at Database T"; Tab(TABFULLVALUE); ValueString;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.OctWaterPartCoeff.databaseT.UNIFAC.error < 0 Then
             Printer.Print Tab(TABFULLCODE); phprop.OctWaterPartCoeff.databaseT.UNIFAC.error
          Else
             Printer.Print Tab(TABFULLCODE); ""
          End If
       End If
    End If

    'Print Octanol Water Partition Coefficient from User Input
    If PROPAVAILABLE(OCT_WATER_PART_COEFF_INPUT) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.OctWaterPartCoeff.input.Value, GetTheFormat(phprop.OctWaterPartCoeff.input.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.OctWaterPartCoeff.input.temperature, GetTheFormat(phprop.OctWaterPartCoeff.input.temperature))
       Printer.Print Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(OCT_WATER_PART_COEFF); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE)
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Printer.Print Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString
    End If
    'Print relevant codes for this property
    If phprop.OctWaterPartCoeff.BinaryInteractionParameterDatabase > 0 Then
       Call PrintTheCodes(phprop.OctWaterPartCoeff.BinaryInteractionParameterDatabase)
       Header = 0
    End If
    If optPrintProperties(1).Value = True Then
       If frmPrint!chkProperties(18).Value = 1 Then
          Call PrintTheErrors(Header, phprop.OctWaterPartCoeff.operatingT.UNIFAC.error)
          Call PrintTheErrors(Header, phprop.OctWaterPartCoeff.database.error)
          Call PrintTheErrors(Header, phprop.OctWaterPartCoeff.databaseT.UNIFAC.error)
       End If
    Else
       Call PrintTheErrors(Header, phprop.OctWaterPartCoeff.operatingT.UNIFAC.error)
       Call PrintTheErrors(Header, phprop.OctWaterPartCoeff.database.error)
       Call PrintTheErrors(Header, phprop.OctWaterPartCoeff.databaseT.UNIFAC.error)
    End If

resume_exit25:
Exit Sub

error_printoctwater:
Resume resume_exit25


End Sub

Private Sub FullyPrintRefractiveIndexToFile()
    Dim ValueString As String

On Error GoTo error_refractive

    'Print Refractive Index From Database
    If PROPAVAILABLE(REFRACTIVE_INDEX_DATABASE) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.RefractiveIndex.database.Value, REFRACTIVE_INDEX_FORMAT)
       Print #1, Tab(TABFULLSOURCE); "Database (" & GetSource(phprop.RefractiveIndex.database.Source.short) & ")"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(REFRACTIVE_INDEX);
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.RefractiveIndex.database.error > 0 Then
             Print #1, Tab(TABFULLCODE); phprop.RefractiveIndex.database.error
          Else
             Print #1, Tab(TABFULLCODE); ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Print #1, Tab(TABFULLSOURCE); "Database"; Tab(TABFULLVALUE); ValueString;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.RefractiveIndex.database.error < 0 Then
             Print #1, Tab(TABFULLCODE); phprop.RefractiveIndex.database.error
          Else
             Print #1, Tab(TABFULLCODE); ""
          End If
       End If
    End If
    'Print Refractive Index from User Input
    If PROPAVAILABLE(REFRACTIVE_INDEX_INPUT) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.RefractiveIndex.input.Value, REFRACTIVE_INDEX_FORMAT)
       Print #1, Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(REFRACTIVE_INDEX)
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Print #1, Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString
    End If
    'Print relevant codes for this property
    If optPrintProperties(1).Value = True Then
       If frmPrint!chkProperties(18).Value = 1 Then
          Call PrintTheErrorsToFile(1, phprop.RefractiveIndex.database.error)
       End If
    Else
       Call PrintTheErrorsToFile(1, phprop.RefractiveIndex.database.error)
    End If

resume_exit26:
Exit Sub

error_refractive:
Resume resume_exit26


End Sub

Private Sub FullyPrintRefractiveIndexToPrinter()
    Dim ValueString As String

On Error GoTo error_printrefractive

    'Print Refractive Index From Database
    If PROPAVAILABLE(REFRACTIVE_INDEX_DATABASE) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.RefractiveIndex.database.Value, REFRACTIVE_INDEX_FORMAT)
       Printer.Print Tab(TABFULLSOURCE); "Database (" & GetSource(phprop.RefractiveIndex.database.Source.short) & ")"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(REFRACTIVE_INDEX);
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.RefractiveIndex.database.error > 0 Then
             Printer.Print Tab(TABFULLCODE); phprop.RefractiveIndex.database.error
          Else
             Printer.Print Tab(TABFULLCODE); ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Printer.Print Tab(TABFULLSOURCE); "Database"; Tab(TABFULLVALUE); ValueString;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.RefractiveIndex.database.error > 0 Then
             Printer.Print Tab(TABFULLCODE); phprop.RefractiveIndex.database.error
          Else
             Printer.Print Tab(TABFULLCODE); ""
          End If
       End If
    End If

    'Print Refractive Index from User Input
    If PROPAVAILABLE(REFRACTIVE_INDEX_INPUT) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.RefractiveIndex.input.Value, REFRACTIVE_INDEX_FORMAT)
       Printer.Print Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(REFRACTIVE_INDEX)
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Printer.Print Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString
    End If
    'Print relevant codes for this property
    If optPrintProperties(1).Value = True Then
       If frmPrint!chkProperties(18).Value = 1 Then
          Call PrintTheErrors(1, phprop.RefractiveIndex.database.error)
       End If
    Else
       Call PrintTheErrors(1, phprop.RefractiveIndex.database.error)
    End If

resume_exit27:
Exit Sub

error_printrefractive:
Resume resume_exit27


End Sub

Private Sub FullyPrintVaporPressureToFile()
    Dim ValueString As String
    Dim TempString As String

On Error GoTo error_vaporpressure

    'Print Vapor Pressure From Database
    If PROPAVAILABLE(VAPOR_PRESSURE_DATABASE) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.VaporPressure.database.Value, GetTheFormat(phprop.VaporPressure.database.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.VaporPressure.database.temperature, GetTheFormat(phprop.VaporPressure.database.temperature))
       Print #1, Tab(TABFULLSOURCE); "Database (" & GetSource(phprop.VaporPressure.database.Source.short) & ")"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(VAPOR_PRESSURE); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE);
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.VaporPressure.database.error > 0 Then
             Print #1, Tab(TABFULLCODE); phprop.VaporPressure.database.error
          Else
             Print #1, Tab(TABFULLCODE); ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Print #1, Tab(TABFULLSOURCE); "Database"; Tab(TABFULLVALUE); ValueString;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.VaporPressure.database.error < 0 Then
             Print #1, Tab(TABFULLCODE); phprop.VaporPressure.database.error
          Else
             Print #1, Tab(TABFULLCODE); ""
          End If
       End If
    End If

    'Print Vapor Pressure from User Input
    If PROPAVAILABLE(VAPOR_PRESSURE_INPUT) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.VaporPressure.input.Value, GetTheFormat(phprop.VaporPressure.input.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.VaporPressure.input.temperature, GetTheFormat(phprop.VaporPressure.input.temperature))
       Print #1, Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(VAPOR_PRESSURE); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE)
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Print #1, Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString
    End If
    'Print relevant codes for this property
    If optPrintProperties(1).Value = True Then
       If frmPrint!chkProperties(18).Value = 1 Then
          Call PrintTheErrorsToFile(1, phprop.VaporPressure.database.error)
       End If
    Else
       Call PrintTheErrorsToFile(1, phprop.VaporPressure.database.error)
    End If

resume_exit28:
Exit Sub

error_vaporpressure:
Resume resume_exit28


End Sub

Private Sub FullyPrintVaporPressureToPrinter()
    Dim ValueString As String
    Dim TempString As String

On Error GoTo error_printvaporpressure

    'Print Vapor Pressure From Database
    If PROPAVAILABLE(VAPOR_PRESSURE_DATABASE) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.VaporPressure.database.Value, GetTheFormat(phprop.VaporPressure.database.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.VaporPressure.database.temperature, GetTheFormat(phprop.VaporPressure.database.temperature))
       Printer.Print Tab(TABFULLSOURCE); "Database (" & GetSource(phprop.VaporPressure.database.Source.short) & ")"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(VAPOR_PRESSURE); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE);
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.VaporPressure.database.error > 0 Then
             Printer.Print Tab(TABFULLCODE); phprop.VaporPressure.database.error
          Else
             Printer.Print Tab(TABFULLCODE); ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Printer.Print Tab(TABFULLSOURCE); "Database"; Tab(TABFULLVALUE); ValueString;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.VaporPressure.database.error < 0 Then
             Printer.Print Tab(TABFULLCODE); phprop.VaporPressure.database.error
          Else
             Printer.Print Tab(TABFULLCODE); ""
          End If
       End If
    End If

    'Print Vapor Pressure from User Input
    If PROPAVAILABLE(VAPOR_PRESSURE_INPUT) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.VaporPressure.input.Value, GetTheFormat(phprop.VaporPressure.input.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.VaporPressure.input.temperature, GetTheFormat(phprop.VaporPressure.input.temperature))
       Printer.Print Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(VAPOR_PRESSURE); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE)
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Printer.Print Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString
    End If
    'Print relevant codes for this property
    If optPrintProperties(1).Value = True Then
       If frmPrint!chkProperties(18).Value = 1 Then
          Call PrintTheErrors(1, phprop.VaporPressure.database.error)
       End If
    Else
       Call PrintTheErrors(1, phprop.VaporPressure.database.error)
    End If

resume_exit29:
Exit Sub

error_printvaporpressure:
Resume resume_exit29


End Sub

Private Sub FullyPrintWaterDensityToFile()
    Dim ValueString As String
    Dim TempString As String

On Error GoTo error_waterdensity
    'Print Water Density from Correlation
    If PROPAVAILABLE(WATER_DENSITY_CORRELATION) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.WaterDensity.correlation.Value, WATER_DENSITY_FORMAT)
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.WaterDensity.correlation.temperature, GetTheFormat(phprop.WaterDensity.correlation.temperature))
       Print #1, Tab(TABFULLSOURCE); GetSource(phprop.WaterDensity.correlation.Source.short); Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(WATER_DENSITY); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE);
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.WaterDensity.correlation.error > 0 Then
             Print #1, Tab(TABFULLCODE); phprop.WaterDensity.correlation.error
          Else
             Print #1, Tab(TABFULLCODE); ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Print #1, Tab(TABFULLSOURCE); GetSource(phprop.WaterDensity.correlation.Source.short); Tab(TABFULLVALUE); ValueString;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.WaterDensity.correlation.error < 0 Then
             Print #1, Tab(TABFULLCODE); phprop.WaterDensity.correlation.error
          Else
             Print #1, Tab(TABFULLCODE); ""
          End If
       End If
    End If

    'Print Water Density from User Input
    If PROPAVAILABLE(WATER_DENSITY_INPUT) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.WaterDensity.input.Value, WATER_DENSITY_FORMAT)
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.WaterDensity.input.temperature, GetTheFormat(phprop.WaterDensity.input.temperature))
       Print #1, Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(WATER_DENSITY); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE)
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Print #1, Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString
    End If
    'Print relevant codes for this property
    If optPrintProperties(1).Value = True Then
       If frmPrint!chkProperties(18).Value = 1 Then
          Call PrintTheErrorsToFile(1, phprop.WaterDensity.correlation.error)
       End If
    Else
       Call PrintTheErrorsToFile(1, phprop.WaterDensity.correlation.error)
    End If

resume_exit30:
Exit Sub

error_waterdensity:
Resume resume_exit30


End Sub

Private Sub FullyPrintWaterDensityToPrinter()
    Dim ValueString As String
    Dim TempString As String

On Error GoTo error_printwaterdensity

    'Print Water Density from Correlation
    If PROPAVAILABLE(WATER_DENSITY_CORRELATION) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.WaterDensity.correlation.Value, WATER_DENSITY_FORMAT)
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.WaterDensity.correlation.temperature, GetTheFormat(phprop.WaterDensity.correlation.temperature))
       Printer.Print Tab(TABFULLSOURCE); GetSource(phprop.WaterDensity.correlation.Source.short); Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(WATER_DENSITY); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE);
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.WaterDensity.correlation.error > 0 Then
             Printer.Print Tab(TABFULLCODE); phprop.WaterDensity.correlation.error
          Else
             Printer.Print Tab(TABFULLCODE); ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Printer.Print Tab(TABFULLSOURCE); GetSource(phprop.WaterDensity.correlation.Source.short); Tab(TABFULLVALUE); ValueString;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.WaterDensity.correlation.error < 0 Then
             Printer.Print Tab(TABFULLCODE); phprop.WaterDensity.correlation.error
          Else
             Printer.Print Tab(TABFULLCODE); ""
          End If
       End If
    End If

    'Print Water Density from User Input
    If PROPAVAILABLE(WATER_DENSITY_INPUT) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.WaterDensity.input.Value, WATER_DENSITY_FORMAT)
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.WaterDensity.input.temperature, GetTheFormat(phprop.WaterDensity.input.temperature))
       Printer.Print Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(WATER_DENSITY); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE)
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Printer.Print Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString
    End If
    'Print relevant codes for this property
    If optPrintProperties(1).Value = True Then
       If frmPrint!chkProperties(18).Value = 1 Then
          Call PrintTheErrors(1, phprop.WaterDensity.correlation.error)
       End If
    Else
       Call PrintTheErrors(1, phprop.WaterDensity.correlation.error)
    End If

resume_exit31:
Exit Sub

error_printwaterdensity:
Resume resume_exit31


End Sub

Private Sub FullyPrintWaterSurfaceTensionToFile()
    Dim ValueString As String
    Dim TempString As String

On Error GoTo error_watersurfacetension

    'Print Water Surface Tension from Correlation
    If PROPAVAILABLE(WATER_SURF_TENSION_CORRELATION) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.WaterSurfaceTension.correlation.Value, GetTheFormat(phprop.WaterSurfaceTension.correlation.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.WaterSurfaceTension.correlation.temperature, GetTheFormat(phprop.WaterSurfaceTension.correlation.temperature))
       Print #1, Tab(TABFULLSOURCE); GetSource(phprop.WaterSurfaceTension.correlation.Source.short); Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(WATER_SURFACE_TENSION); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE);
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.WaterSurfaceTension.correlation.error > 0 Then
             Print #1, Tab(TABFULLCODE); phprop.WaterSurfaceTension.correlation.error
          Else
             Print #1, Tab(TABFULLCODE); ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Print #1, Tab(TABFULLSOURCE); GetSource(phprop.WaterSurfaceTension.correlation.Source.short); Tab(TABFULLVALUE); ValueString;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.WaterSurfaceTension.correlation.error < 0 Then
             Print #1, Tab(TABFULLCODE); phprop.WaterSurfaceTension.correlation.error
          Else
             Print #1, Tab(TABFULLCODE); ""
          End If
       End If
    End If

    'Print Water Surface Tension from User Input
    If PROPAVAILABLE(WATER_SURF_TENSION_INPUT) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.WaterSurfaceTension.input.Value, GetTheFormat(phprop.WaterSurfaceTension.input.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.WaterSurfaceTension.input.temperature, GetTheFormat(phprop.WaterSurfaceTension.input.temperature))
       Print #1, Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(WATER_SURFACE_TENSION); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE)
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Print #1, Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString
    End If
    'Print relevant codes for this property
    If optPrintProperties(1).Value = True Then
       If frmPrint!chkProperties(18).Value = 1 Then
          Call PrintTheErrorsToFile(1, phprop.WaterSurfaceTension.correlation.error)
       End If
    Else
       Call PrintTheErrorsToFile(1, phprop.WaterSurfaceTension.correlation.error)
    End If

resume_exit32:
Exit Sub

error_watersurfacetension:
Resume resume_exit32


End Sub

Private Sub FullyPrintWaterSurfaceTensionToPrinter()
    Dim ValueString As String
    Dim TempString As String

On Error GoTo error_printwatersurfacetension

    'Print Water Surface Tension from Correlation
    If PROPAVAILABLE(WATER_SURF_TENSION_CORRELATION) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.WaterSurfaceTension.correlation.Value, GetTheFormat(phprop.WaterSurfaceTension.correlation.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.WaterSurfaceTension.correlation.temperature, GetTheFormat(phprop.WaterSurfaceTension.correlation.temperature))
       Printer.Print Tab(TABFULLSOURCE); GetSource(phprop.WaterSurfaceTension.correlation.Source.short); Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(WATER_SURFACE_TENSION); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE);
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.WaterSurfaceTension.correlation.error > 0 Then
             Printer.Print Tab(TABFULLCODE); phprop.WaterSurfaceTension.correlation.error
          Else
             Printer.Print Tab(TABFULLCODE); ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Printer.Print Tab(TABFULLSOURCE); GetSource(phprop.WaterSurfaceTension.correlation.Source.short); Tab(TABFULLVALUE); ValueString;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.WaterSurfaceTension.correlation.error < 0 Then
             Printer.Print Tab(TABFULLCODE); phprop.WaterSurfaceTension.correlation.error
          Else
             Printer.Print Tab(TABFULLCODE); ""
          End If
       End If
    End If

    'Print Water Surface Tension from User Input
    If PROPAVAILABLE(WATER_SURF_TENSION_INPUT) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.WaterSurfaceTension.input.Value, GetTheFormat(phprop.WaterSurfaceTension.input.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.WaterSurfaceTension.input.temperature, GetTheFormat(phprop.WaterSurfaceTension.input.temperature))
       Printer.Print Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(WATER_SURFACE_TENSION); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE)
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Printer.Print Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString
    End If
    'Print relevant codes for this property
    If optPrintProperties(1).Value = True Then
       If frmPrint!chkProperties(18).Value = 1 Then
          Call PrintTheErrors(1, phprop.WaterSurfaceTension.correlation.error)
       End If
    Else
       Call PrintTheErrors(1, phprop.WaterSurfaceTension.correlation.error)
    End If

resume_exit33:
Exit Sub

error_printwatersurfacetension:
Resume resume_exit33


End Sub

Private Sub FullyPrintWaterViscosityToFile()
    Dim ValueString As String
    Dim TempString As String

On Error GoTo error_waterviscosity

    'Print Water Viscosity from Correlation
    If PROPAVAILABLE(WATER_VISCOSITY_CORRELATION) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.WaterViscosity.correlation.Value, GetTheFormat(phprop.WaterViscosity.correlation.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.WaterViscosity.correlation.temperature, GetTheFormat(phprop.WaterViscosity.correlation.temperature))
       Print #1, Tab(TABFULLSOURCE); GetSource(phprop.WaterViscosity.correlation.Source.short); Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(WATER_VISCOSITY); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE);
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.WaterViscosity.correlation.error > 0 Then
             Print #1, Tab(TABFULLCODE); phprop.WaterViscosity.correlation.error
          Else
             Print #1, Tab(TABFULLCODE); ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Print #1, Tab(TABFULLSOURCE); GetSource(phprop.WaterViscosity.correlation.Source.short); Tab(TABFULLVALUE); ValueString;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.WaterViscosity.correlation.error < 0 Then
             Print #1, Tab(TABFULLCODE); phprop.WaterViscosity.correlation.error
          Else
             Print #1, Tab(TABFULLCODE); ""
          End If
       End If
    End If

    'Print Water Viscosity from User Input
    If PROPAVAILABLE(WATER_VISCOSITY_INPUT) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.WaterViscosity.input.Value, GetTheFormat(phprop.WaterViscosity.input.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.WaterViscosity.input.temperature, GetTheFormat(phprop.WaterViscosity.input.temperature))
       Print #1, Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(WATER_VISCOSITY); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE)
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Print #1, Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString
    End If
    'Print relevant codes for this property
    If optPrintProperties(1).Value = True Then
       If frmPrint!chkProperties(18).Value = 1 Then
          Call PrintTheErrorsToFile(1, phprop.WaterViscosity.correlation.error)
       End If
    Else
       Call PrintTheErrorsToFile(1, phprop.WaterViscosity.correlation.error)
    End If

resume_exit34:
Exit Sub

error_waterviscosity:
Resume resume_exit34

End Sub

Private Sub FullyPrintWaterViscosityToPrinter()
    Dim ValueString As String
    Dim TempString As String

On Error GoTo error_printwaterviscosity

    'Print Water Viscosity from Correlation
    If PROPAVAILABLE(WATER_VISCOSITY_CORRELATION) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.WaterViscosity.correlation.Value, GetTheFormat(phprop.WaterViscosity.correlation.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.WaterViscosity.correlation.temperature, GetTheFormat(phprop.WaterViscosity.correlation.temperature))
       Printer.Print Tab(TABFULLSOURCE); GetSource(phprop.WaterViscosity.correlation.Source.short); Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(WATER_VISCOSITY); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE);
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.WaterViscosity.correlation.error > 0 Then
             Printer.Print Tab(TABFULLCODE); phprop.WaterViscosity.correlation.error
          Else
             Printer.Print Tab(TABFULLCODE); ""
          End If
       End If
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Printer.Print Tab(TABFULLSOURCE); GetSource(phprop.WaterViscosity.correlation.Source.short); Tab(TABFULLVALUE); ValueString;
       If frmPrint!chkProperties(18).Value = 1 Then
          If phprop.WaterViscosity.correlation.error < 0 Then
             Printer.Print Tab(TABFULLCODE); phprop.WaterViscosity.correlation.error
          Else
             Printer.Print Tab(TABFULLCODE); ""
          End If
       End If
    End If

    'Print Water Viscosity from User Input
    If PROPAVAILABLE(WATER_VISCOSITY_INPUT) Then
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = Format$(phprop.WaterViscosity.input.Value, GetTheFormat(phprop.WaterViscosity.input.Value))
       TempString = Space$(TEMPLENGTH)
       RSet TempString = Format$(phprop.WaterViscosity.input.temperature, GetTheFormat(phprop.WaterViscosity.input.temperature))
       Printer.Print Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString; Tab(TABFULLUNITS); Units(WATER_VISCOSITY); Tab(TABFULLTEMPERATURE); TempString; "  "; Units(OPERATING_TEMPERATURE)
    Else
       ValueString = Space$(VALUELENGTH)
       RSet ValueString = "Not Available"
       Printer.Print Tab(TABFULLSOURCE); "User Input"; Tab(TABFULLVALUE); ValueString
    End If
    'Print relevant codes for this property
    If optPrintProperties(1).Value = True Then
       If frmPrint!chkProperties(18).Value = 1 Then
          Call PrintTheErrors(1, phprop.WaterViscosity.correlation.error)
       End If
    Else
       Call PrintTheErrors(1, phprop.WaterViscosity.correlation.error)
    End If

resume_exit35:
Exit Sub

error_printwaterviscosity:
Resume resume_exit35


End Sub

Private Sub GetPrintFileName()

    On Error Resume Next
    frmPrint!CMDialog1.DefaultExt = "out"
    frmPrint!CMDialog1.Filter = "StEPP Output Files (*.out)|*.out"
    frmPrint!CMDialog1.DialogTitle = "Print StEPP Output to File"
    frmPrint!CMDialog1.flags = OFN_OVERWRITEPROMPT Or OFN_PATHMUSTEXIST
    frmPrint!CMDialog1.CancelError = True
    frmPrint!CMDialog1.Action = 2
    PrintFileName$ = frmPrint!CMDialog1.FileName
    If Err = 32755 Then   'Cancel selected by user
       PrintFileName$ = ""
    End If

End Sub

Private Function GetSource(Source As Long) As String

   Select Case Source
      Case 1
         GetSource = "YAWS"
      Case 2
         GetSource = "SUPERFUND"
      Case 3
         GetSource = "RTI"
      Case 4
         GetSource = "DIPPR801"
      Case 5
         GetSource = "BRI/NASA"
      Case 6
         GetSource = "Clean Air Act"
      Case 7
         GetSource = "UNIFAC"
      Case 8
         GetSource = "Schroeder's Method"
      Case 9
         GetSource = "Group Contribution Method"
      Case 10
         GetSource = "Hayduk & Laudie"
      Case 11
         GetSource = "Polson"
      Case 12
         GetSource = "Wilke-Chang"
      Case 13
         GetSource = "Wilke-Lee"
      Case 14
         GetSource = "Data Correlation"
      Case 15
         GetSource = "Cummins"
      Case 16
         GetSource = "Ideal Gas Law"
      Case 17
         GetSource = "HC Regression"
      Case 18
         GetSource = "HC UNIFAC Fit"
      Case 19
         GetSource = "Solubility UNIFAC Fit"
   End Select

End Function

Private Sub optPrintContaminants_Click(Index As Integer)
    
    Select Case Index
       Case 0   'All Contaminants
          lblCurrentContaminant.ForeColor = &H80000008
          lblCurrentContaminant.BackColor = &HC0C0C0
       Case 1   'Currently selected contaminant
          lblCurrentContaminant.ForeColor = &HFFFFFF
          lblCurrentContaminant.BackColor = &H800000
    End Select
End Sub

Private Sub optPrintProperties_Click(Index As Integer)
    Dim i As Integer

    Select Case Index
       Case 0   'All Properties
          For i = 0 To 18
              chkProperties(i).BackColor = &HC0C0C0
              chkProperties(i).ForeColor = &H80000008
              chkProperties(i).Enabled = False
          Next i
       Case 1   'Selected Properties
          For i = 0 To 18
              If chkProperties(i).Value Then
                 chkProperties(i).BackColor = &H800000
                 chkProperties(i).ForeColor = &HFFFFFF
              End If
              chkProperties(i).Enabled = True
          Next i
    End Select
    
    If frmPrint!cboPropertyDescription.ListIndex = 0 Then
       chkProperties(18).Enabled = False
    End If

End Sub

Private Sub PrintActivityCoefficientPrinter()
    Dim ValueString As String


    Select Case frmPrint!cboPropertyDescription.ListIndex
       Case 0   'Print Selected Value Only
          If HaveProperty(ACTIVITY_COEFFICIENT) Then
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = Format$(phprop.ActivityCoefficient.CurrentSelection.Value, GetTheFormat(phprop.ActivityCoefficient.CurrentSelection.Value))
             Printer.Print "Activity Coefficient"; Tab(TABVALUE); ValueString; Tab(TABUNITS); Units(ACTIVITY_COEFFICIENT); Tab(TABSOURCE); GetSource(phprop.ActivityCoefficient.CurrentSelection.Source)
          Else
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = "Not Available"
             Printer.Print "Activity Coefficient"; Tab(TABVALUE); ValueString
          End If
       Case 1   'Print Full Description of Activity Coefficient
          HeightActivityCoefficient = 0
          Printer.FontSize = 12
          PrintMsg = ""
          HeightActivityCoefficient = HeightActivityCoefficient + NUMLINES_PROPERTY_NAME * Printer.TextHeight(PrintMsg)
          Printer.FontSize = 10
          HeightActivityCoefficient = HeightActivityCoefficient + NUMLINES_ACTIVITY_COEFFICIENT * Printer.TextHeight(PrintMsg)
          TotalHeightThisPage = Printer.CurrentY
          If TotalHeightThisPage + HeightActivityCoefficient + BOTTOM_MARGIN_SAFETY_FACTOR > Printer.Height Then
             Printer.NewPage
             Call PrintTitleContinuation
          End If

          Printer.FontBold = True
          Printer.FontSize = 12
          Printer.Print "Property:  INFINITE DILUTION ACTIVITY COEFFICIENT"
          Printer.Print
          Printer.FontBold = False
          Printer.FontUnderline = True
          Printer.FontSize = 10
          Printer.Print Tab(TABFULLSOURCE); "Source:"; Tab(TABFULLVALUE); "Value:"; Tab(TABFULLUNITS); "Units:"; Tab(TABFULLTEMPERATURE); "Temp.:"; Tab(TABFULLCODE); "Code:"
          Printer.FontUnderline = False
          Printer.Print
          Call FullyPrintActivityCoefficientToPrinter
          Printer.Print
          Printer.Print

    End Select

End Sub

Private Sub PrintActivityCoefficientToFile()
    Dim ValueString As String

    Select Case frmPrint!cboPropertyDescription.ListIndex
       Case 0   'Print Selected Value Only
          If HaveProperty(ACTIVITY_COEFFICIENT) Then
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = Format$(phprop.ActivityCoefficient.CurrentSelection.Value, GetTheFormat(phprop.ActivityCoefficient.CurrentSelection.Value))
             Print #1, "Activity Coefficient"; Tab(TABVALUE); ValueString; Tab(TABUNITS); Units(ACTIVITY_COEFFICIENT); Tab(TABSOURCE); GetSource(phprop.ActivityCoefficient.CurrentSelection.Source)
          Else
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = "Not Available"
             Print #1, "Activity Coefficient"; Tab(TABVALUE); ValueString
          End If
       Case 1   'Print Full Description of Activity Coefficient
          Print #1, "Property:  INFINITE DILUTION ACTIVITY COEFFICIENT"
          Print #1,
          Print #1, Tab(TABFULLSOURCE); "Source:"; Tab(TABFULLVALUE); "Value:"; Tab(TABFULLUNITS); "Units:"; Tab(TABFULLTEMPERATURE); "Temp.:"; Tab(TABFULLCODE); "Code:"
          Print #1,
          Call FullyPrintActivityCoefficientToFile
          Print #1,
          Print #1,

    End Select

End Sub

Private Sub PrintAirDensityPrinter()
    Dim ValueString As String

    Select Case frmPrint!cboPropertyDescription.ListIndex
       Case 0   'Print Selected Value Only
          If HaveProperty(AIR_DENSITY) Then
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = Format$(phprop.AirDensity.CurrentSelection.Value, GetTheFormat(phprop.AirDensity.CurrentSelection.Value))
             Printer.Print "Air Density"; Tab(TABVALUE); ValueString; Tab(TABUNITS); Units(AIR_DENSITY); Tab(TABSOURCE); GetSource(phprop.AirDensity.CurrentSelection.Source)
          Else
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = "Not Available"
             Printer.Print "Air Density"; Tab(TABVALUE); ValueString
          End If
       Case 1   'Print Full Description of Air Density
          HeightAirDensity = 0
          Printer.FontSize = 12
          PrintMsg = ""
          HeightAirDensity = HeightAirDensity + NUMLINES_PROPERTY_NAME * Printer.TextHeight(PrintMsg)
          Printer.FontSize = 10
          HeightAirDensity = HeightAirDensity + NUMLINES_AIR_DENSITY * Printer.TextHeight(PrintMsg)
          TotalHeightThisPage = Printer.CurrentY
          If TotalHeightThisPage + HeightAirDensity + BOTTOM_MARGIN_SAFETY_FACTOR > Printer.Height Then
             Printer.NewPage
             Call PrintAirWaterTitleContinuation
          End If

          Printer.FontBold = True
          Printer.FontSize = 12
          Printer.Print "Property:  AIR DENSITY"
          Printer.Print
          Printer.FontBold = False
          Printer.FontUnderline = True
          Printer.FontSize = 10
          Printer.Print Tab(TABFULLSOURCE); "Source:"; Tab(TABFULLVALUE); "Value:"; Tab(TABFULLUNITS); "Units:"; Tab(TABFULLTEMPERATURE); "Temp.:"; Tab(TABFULLCODE); "Code:"
          Printer.FontUnderline = False
          Printer.Print
          Call FullyPrintAirDensityToPrinter
          Printer.Print
          Printer.Print

    End Select

End Sub

Private Sub PrintAirDensityToFile()
    Dim ValueString As String

    Select Case frmPrint!cboPropertyDescription.ListIndex
       Case 0   'Print Selected Value Only
          If HaveProperty(AIR_DENSITY) Then
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = Format$(phprop.AirDensity.CurrentSelection.Value, GetTheFormat(phprop.AirDensity.CurrentSelection.Value))
             Print #1, "Air Density"; Tab(TABVALUE); ValueString; Tab(TABUNITS); Units(AIR_DENSITY); Tab(TABSOURCE); GetSource(phprop.AirDensity.CurrentSelection.Source)
          Else
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = "Not Available"
             Print #1, "Air Density"; Tab(TABVALUE); ValueString
          End If
       Case 1   'Print Full Description of Air Density
          Print #1, "Property:  AIR DENSITY"
          Print #1,
          Print #1, Tab(TABFULLSOURCE); "Source:"; Tab(TABFULLVALUE); "Value:"; Tab(TABFULLUNITS); "Units:"; Tab(TABFULLTEMPERATURE); "Temp.:"; Tab(TABFULLCODE); "Code:"
          Print #1,
          Call FullyPrintAirDensityToFile
          Print #1,
          Print #1,

    End Select

End Sub

Private Sub PrintAirViscosityPrinter()
    Dim ValueString As String

    Select Case frmPrint!cboPropertyDescription.ListIndex
       Case 0   'Print Selected Value Only
          If HaveProperty(AIR_VISCOSITY) Then
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = Format$(phprop.AirViscosity.CurrentSelection.Value, GetTheFormat(phprop.AirViscosity.CurrentSelection.Value))
             Printer.Print "Air Viscosity"; Tab(TABVALUE); ValueString; Tab(TABUNITS); Units(AIR_VISCOSITY); Tab(TABSOURCE); GetSource(phprop.AirViscosity.CurrentSelection.Source)
          Else
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = "Not Available"
             Printer.Print "Air Viscosity"; Tab(TABVALUE); ValueString
          End If

       Case 1   'Print Full Description of Air Density
          HeightAirViscosity = 0
          Printer.FontSize = 12
          PrintMsg = ""
          HeightAirViscosity = HeightAirViscosity + NUMLINES_PROPERTY_NAME * Printer.TextHeight(PrintMsg)
          Printer.FontSize = 10
          HeightAirViscosity = HeightAirViscosity + NUMLINES_AIR_VISCOSITY * Printer.TextHeight(PrintMsg)
          TotalHeightThisPage = Printer.CurrentY
          If TotalHeightThisPage + HeightAirViscosity + BOTTOM_MARGIN_SAFETY_FACTOR > Printer.Height Then
             Printer.NewPage
             Call PrintAirWaterTitleContinuation
          End If

          Printer.FontBold = True
          Printer.FontSize = 12
          Printer.Print "Property:  AIR VISCOSITY"
          Printer.Print
          Printer.FontBold = False
          Printer.FontUnderline = True
          Printer.FontSize = 10
          Printer.Print Tab(TABFULLSOURCE); "Source:"; Tab(TABFULLVALUE); "Value:"; Tab(TABFULLUNITS); "Units:"; Tab(TABFULLTEMPERATURE); "Temp.:"; Tab(TABFULLCODE); "Code:"
          Printer.FontUnderline = False
          Printer.Print
          Call FullyPrintAirViscosityToPrinter
          Printer.Print
          Printer.Print

    End Select

End Sub

Private Sub PrintAirViscosityToFile()
    Dim ValueString As String

    Select Case frmPrint!cboPropertyDescription.ListIndex
       Case 0   'Print Selected Value Only
          If HaveProperty(AIR_VISCOSITY) Then
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = Format$(phprop.AirViscosity.CurrentSelection.Value, GetTheFormat(phprop.AirViscosity.CurrentSelection.Value))
             Print #1, "Air Viscosity"; Tab(TABVALUE); ValueString; Tab(TABUNITS); Units(AIR_VISCOSITY); Tab(TABSOURCE); GetSource(phprop.AirViscosity.CurrentSelection.Source)
          Else
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = "Not Available"
             Print #1, "Air Viscosity"; Tab(TABVALUE); ValueString
          End If

       Case 1   'Print Full Description of Air Density
          Print #1, "Property:  AIR VISCOSITY"
          Print #1,
          Print #1, Tab(TABFULLSOURCE); "Source:"; Tab(TABFULLVALUE); "Value:"; Tab(TABFULLUNITS); "Units:"; Tab(TABFULLTEMPERATURE); "Temp.:"; Tab(TABFULLCODE); "Code:"
          Print #1,
          Call FullyPrintAirViscosityToFile
          Print #1,
          Print #1,

    End Select

End Sub

Private Sub PrintAirWaterProperties()

       Printer.FontSize = 14
       Printer.FontBold = True
       Printer.Print "PROPERTIES OF AIR AND WATER"
       Printer.Print
       Printer.Print

       Select Case frmPrint!cboPropertyDescription.ListIndex
          Case 0   'Print Selected Value Only
             Printer.FontBold = False
             Printer.Print "Summary of Selected Values"
             Printer.Print
             Printer.Print
             Printer.FontSize = 10
             Printer.FontBold = True
             Printer.FontUnderline = True
             Printer.Print "Property:"; Tab(TABVALUE); "Value:"; Tab(TABUNITS); "Units:"; Tab(TABSOURCE); "Source:"
             Printer.Print
             Printer.FontUnderline = False
             Printer.FontBold = False
             Printer.FontSize = 10
             Call PrintOperatingPressure
             Call PrintOperatingTemperature
             Printer.Print

             If frmPrint!optPrintProperties(0).Value Then  'Print all properties
                Call PrintAllAirWaterPropertiesToPrinter
             ElseIf frmPrint!optPrintProperties(1).Value Then  'Print chosen properties only
                Call PrintChosenAirWaterPropertiesToPrinter
             End If
          Case 1   'Print Full Description of Properties
             Printer.FontBold = False
             Printer.Print "Full Description of Properties"
             Printer.Print
             Printer.Print
             HeightTitle = Printer.CurrentY
             Printer.FontSize = 10
             Printer.FontBold = True
             Printer.FontUnderline = True
             Printer.Print "Operating Conditions:"; Tab(TABVALUE); "Value:"; Tab(TABUNITS); "Units:"; Tab(TABSOURCE); "Source:"
             Printer.Print
             Printer.FontUnderline = False
             Printer.FontBold = False
             Printer.FontSize = 10
             Call PrintOperatingPressure
             Call PrintOperatingTemperature
             Printer.Print
             Printer.Print
             HeightOperatingConditions = Printer.CurrentY - HeightTitle

             If frmPrint!optPrintProperties(0).Value Then  'Print all properties
                Call PrintAllAirWaterPropertiesToPrinter
             ElseIf frmPrint!optPrintProperties(1).Value Then  'Print chosen properties only
                Call PrintChosenAirWaterPropertiesToPrinter
             End If

       End Select

End Sub

Private Sub PrintAirWaterPropertiesToFile()

       Print #1, "PROPERTIES OF AIR AND WATER"
       Print #1,
       Print #1,
       Print #1,

       Select Case frmPrint!cboPropertyDescription.ListIndex
          Case 0   'Print Selected Value Only
             Print #1, "Summary of Selected Values"
             Print #1,
             Print #1,
             Print #1, "Property:"; Tab(TABVALUE); "Value:"; Tab(TABUNITS); "Units:"; Tab(TABSOURCE); "Source:"
             Print #1,
             Call PrintOperatingPressureToFile
             Call PrintOperatingTemperatureToFile
             Print #1,

             If frmPrint!optPrintProperties(0).Value Then  'Print all properties
                Call PrintAllAirWaterPropertiesToFile
             ElseIf frmPrint!optPrintProperties(1).Value Then  'Print chosen properties only
                Call PrintChosenAirWaterPropertiesToFile
             End If
          Case 1   'Print Full Description of Properties
             Print #1, "Full Description of Properties"
             Print #1,
             Print #1,
             Print #1, "Property:"; Tab(TABVALUE); "Value:"; Tab(TABUNITS); "Units:"; Tab(TABSOURCE); "Source:"
             Print #1,
             Call PrintOperatingPressureToFile
             Call PrintOperatingTemperatureToFile
             Print #1,
             Print #1,

             If frmPrint!optPrintProperties(0).Value Then  'Print all properties
                Call PrintAllAirWaterPropertiesToFile
             ElseIf frmPrint!optPrintProperties(1).Value Then  'Print chosen properties only
                Call PrintChosenAirWaterPropertiesToFile
             End If

       End Select
    
End Sub

Private Sub PrintAirWaterTitleContinuation()

       Printer.FontSize = 14
       Printer.FontBold = True
       Printer.Print "PROPERTIES OF AIR AND WATER";
       Printer.FontBold = False
       Printer.Print " (continued)"
       Printer.Print
       Printer.Print
       Printer.FontBold = False
       Printer.Print "Full Description of Properties"
       Printer.Print
       Printer.Print
       HeightTitle = Printer.CurrentY
       Printer.FontSize = 10
       Printer.FontBold = True
       Printer.FontUnderline = True
       Printer.Print "Operating Conditions:"; Tab(TABVALUE); "Value:"; Tab(TABUNITS); "Units:"; Tab(TABSOURCE); "Source:"
       Printer.Print
       Printer.FontUnderline = False
       Printer.FontBold = False
       Printer.FontSize = 10
       Call PrintOperatingPressure
       Call PrintOperatingTemperature
       Printer.Print
       Printer.Print
       HeightOperatingConditions = Printer.CurrentY - HeightTitle

End Sub

Private Sub PrintAllAirWaterPropertiesToFile()

       Call PrintWaterDensityToFile
       Call PrintWaterViscosityToFile
       Call PrintWaterSurfaceTensionToFile
       Call PrintAirDensityToFile
       Call PrintAirViscosityToFile

End Sub

Private Sub PrintAllAirWaterPropertiesToPrinter()

       Call PrintWaterDensityPrinter
       Call PrintWaterViscosityPrinter
       Call PrintWaterSurfaceTensionPrinter
       Call PrintAirDensityPrinter
       Call PrintAirViscosityPrinter

End Sub

Private Sub PrintAllPropertiesToFile()

       Call PrintVaporPressureToFile
       Call PrintActivityCoefficientToFile
       Call PrintHenrysConstantToFile
       Call PrintMolecularWeightToFile
       Call PrintBoilingPointToFile
       Call PrintLiquidDensityToFile
       Call PrintMolarVolumeOpTToFile
       Call PrintMolarVolumeNBPToFile
       Call PrintRefractiveIndexToFile
       Call PrintAqueousSolubilityToFile
       Call PrintOctWaterPartCoeffToFile
       Call PrintLiquidDiffusivityToFile
       Call PrintGasDiffusivityToFile

End Sub

Private Sub PrintAllPropertiesToPrinter()

       Call PrintVaporPressurePrinter
       Call PrintActivityCoefficientPrinter
       Call PrintHenrysConstantPrinter
       Call PrintMolecularWeightPrinter
       Call PrintBoilingPointPrinter
       Call PrintLiquidDensityPrinter
       Call PrintMolarVolumeOpTPrinter
       Call PrintMolarVolumeNBPPrinter
       Call PrintRefractiveIndexPrinter
       Call PrintAqueousSolubilityPrinter
       Call PrintOctWaterPartCoeffPrinter
       Call PrintLiquidDiffusivityPrinter
       Call PrintGasDiffusivityPrinter

End Sub

Private Sub PrintAqueousSolubilityPrinter()
    Dim ValueString As String

    Select Case frmPrint!cboPropertyDescription.ListIndex
       Case 0   'Print Selected Value Only
          If HaveProperty(AQUEOUS_SOLUBILITY) Then
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = Format$(phprop.AqueousSolubility.CurrentSelection.Value, GetTheFormat(phprop.AqueousSolubility.CurrentSelection.Value))
             Printer.Print "Aqueous Solubility"; Tab(TABVALUE); ValueString; Tab(TABUNITS); Units(AQUEOUS_SOLUBILITY); Tab(TABSOURCE); GetSource(phprop.AqueousSolubility.CurrentSelection.Source)
          Else
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = "Not Available"
             Printer.Print "Aqueous Solubility"; Tab(TABVALUE); ValueString
          End If
       Case 1   'Print Full Description of Aqueous Solubility
          HeightAqueousSolubility = 0
          Printer.FontSize = 12
          PrintMsg = ""
          HeightAqueousSolubility = HeightAqueousSolubility + NUMLINES_PROPERTY_NAME * Printer.TextHeight(PrintMsg)
          Printer.FontSize = 10
          HeightAqueousSolubility = HeightAqueousSolubility + NUMLINES_AQUEOUS_SOLUBILITY * Printer.TextHeight(PrintMsg)
          TotalHeightThisPage = Printer.CurrentY
          If TotalHeightThisPage + HeightAqueousSolubility + BOTTOM_MARGIN_SAFETY_FACTOR > Printer.Height Then
             Printer.NewPage
             Call PrintTitleContinuation
          End If

          Printer.FontBold = True
          Printer.FontSize = 12
          Printer.Print "Property:  AQUEOUS SOLUBILITY"
          Printer.Print
          Printer.FontBold = False
          Printer.FontUnderline = True
          Printer.FontSize = 10
          Printer.Print Tab(TABFULLSOURCE); "Source:"; Tab(TABFULLVALUE); "Value:"; Tab(TABFULLUNITS); "Units:"; Tab(TABFULLTEMPERATURE); "Temp.:"; Tab(TABFULLCODE); "Code:"
          Printer.FontUnderline = False
          Printer.Print
          Call FullyPrintAqueousSolubilityToPrinter
          Printer.Print
          Printer.Print

    End Select

End Sub

Private Sub PrintAqueousSolubilityToFile()
    Dim ValueString As String

    Select Case frmPrint!cboPropertyDescription.ListIndex
       Case 0   'Print Selected Value Only
          If HaveProperty(AQUEOUS_SOLUBILITY) Then
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = Format$(phprop.AqueousSolubility.CurrentSelection.Value, GetTheFormat(phprop.AqueousSolubility.CurrentSelection.Value))
             Print #1, "Aqueous Solubility"; Tab(TABVALUE); ValueString; Tab(TABUNITS); Units(AQUEOUS_SOLUBILITY); Tab(TABSOURCE); GetSource(phprop.AqueousSolubility.CurrentSelection.Source)
          Else
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = "Not Available"
             Print #1, "Aqueous Solubility"; Tab(TABVALUE); ValueString
          End If
       Case 1   'Print Full Description of Aqueous Solubility
          Print #1, "Property:  AQUEOUS SOLUBILITY"
          Print #1,
          Print #1, Tab(TABFULLSOURCE); "Source:"; Tab(TABFULLVALUE); "Value:"; Tab(TABFULLUNITS); "Units:"; Tab(TABFULLTEMPERATURE); "Temp.:"; Tab(TABFULLCODE); "Code:"
          Print #1,
          Call FullyPrintAqueousSolubilityToFile
          Print #1,
          Print #1,

    End Select

End Sub

Private Sub PrintBoilingPointPrinter()
    Dim ValueString As String

    Select Case frmPrint!cboPropertyDescription.ListIndex
       Case 0   'Print Selected Value Only
          If HaveProperty(BOILING_POINT) Then
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = Format$(phprop.BoilingPoint.CurrentSelection.Value, GetTheFormat(phprop.BoilingPoint.CurrentSelection.Value))
             Printer.Print "Normal Boiling Point (NBP)"; Tab(TABVALUE); ValueString; Tab(TABUNITS); Units(BOILING_POINT); Tab(TABSOURCE); GetSource(phprop.BoilingPoint.CurrentSelection.Source)
          Else
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = "Not Available"
             Printer.Print "Normal Boiling Point (NBP)"; Tab(TABVALUE); ValueString
          End If
       Case 1   'Print Full Description of Normal Boiling Point
          HeightBoilingPoint = 0
          Printer.FontSize = 12
          PrintMsg = ""
          HeightBoilingPoint = HeightBoilingPoint + NUMLINES_PROPERTY_NAME * Printer.TextHeight(PrintMsg)
          Printer.FontSize = 10
          HeightBoilingPoint = HeightBoilingPoint + NUMLINES_BOILING_POINT * Printer.TextHeight(PrintMsg)
          TotalHeightThisPage = Printer.CurrentY
          If TotalHeightThisPage + HeightBoilingPoint + BOTTOM_MARGIN_SAFETY_FACTOR > Printer.Height Then
             Printer.NewPage
             Call PrintTitleContinuation
          End If

          Printer.FontBold = True
          Printer.FontSize = 12
          Printer.Print "Property:  NORMAL BOILING POINT"
          Printer.Print
          Printer.FontBold = False
          Printer.FontUnderline = True
          Printer.FontSize = 10
          Printer.Print Tab(TABFULLSOURCE); "Source:"; Tab(TABFULLVALUE); "Value:"; Tab(TABFULLUNITS); "Units:"; Tab(TABFULLCODE); "Code:"
          Printer.FontUnderline = False
          Printer.Print
          Call FullyPrintBoilingPointToPrinter
          Printer.Print
          Printer.Print

    End Select

End Sub

Private Sub PrintBoilingPointToFile()
    Dim ValueString As String

    Select Case frmPrint!cboPropertyDescription.ListIndex
       Case 0   'Print Selected Value Only
          If HaveProperty(BOILING_POINT) Then
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = Format$(phprop.BoilingPoint.CurrentSelection.Value, GetTheFormat(phprop.BoilingPoint.CurrentSelection.Value))
             Print #1, "Normal Boiling Point (NBP)"; Tab(TABVALUE); ValueString; Tab(TABUNITS); Units(BOILING_POINT); Tab(TABSOURCE); GetSource(phprop.BoilingPoint.CurrentSelection.Source)
          Else
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = "Not Available"
             Print #1, "Normal Boiling Point (NBP)"; Tab(TABVALUE); ValueString
          End If
       Case 1   'Print Full Description of Normal Boiling Point
          Print #1, "Property:  NORMAL BOILING POINT"
          Print #1,
          Print #1, Tab(TABFULLSOURCE); "Source:"; Tab(TABFULLVALUE); "Value:"; Tab(TABFULLUNITS); "Units:"; Tab(TABFULLCODE); "Code:"
          Print #1,
          Call FullyPrintBoilingPointToFile
          Print #1,
          Print #1,

    End Select

End Sub

Private Sub PrintChosenAirWaterPropertiesToFile()

    If frmPrint!chkProperties(13).Value = 1 Then
       Call PrintWaterDensityToFile
    End If

    If frmPrint!chkProperties(14).Value = 1 Then
       Call PrintWaterViscosityToFile
    End If

    If frmPrint!chkProperties(15).Value = 1 Then
       Call PrintWaterSurfaceTensionToFile
    End If

    If frmPrint!chkProperties(16).Value = 1 Then
       Call PrintAirDensityToFile
    End If

    If frmPrint!chkProperties(17).Value = 1 Then
       Call PrintAirViscosityToFile
    End If


End Sub

Private Sub PrintChosenAirWaterPropertiesToPrinter()

    If frmPrint!chkProperties(13).Value = 1 Then
       Call PrintWaterDensityPrinter
    End If

    If frmPrint!chkProperties(14).Value = 1 Then
       Call PrintWaterViscosityPrinter
    End If

    If frmPrint!chkProperties(15).Value = 1 Then
       Call PrintWaterSurfaceTensionPrinter
    End If

    If frmPrint!chkProperties(16).Value = 1 Then
       Call PrintAirDensityPrinter
    End If

    If frmPrint!chkProperties(17).Value = 1 Then
       Call PrintAirViscosityPrinter
    End If

End Sub

Private Sub PrintChosenPropertiesToFile()

    If frmPrint!chkProperties(0).Value = 1 Then
       Call PrintVaporPressureToFile
    End If

    If frmPrint!chkProperties(1).Value = 1 Then
       Call PrintActivityCoefficientToFile
    End If

    If frmPrint!chkProperties(2).Value = 1 Then
       Call PrintHenrysConstantToFile
    End If

    If frmPrint!chkProperties(3).Value = 1 Then
       Call PrintMolecularWeightToFile
    End If

    If frmPrint!chkProperties(4).Value = 1 Then
       Call PrintBoilingPointToFile
    End If

    If frmPrint!chkProperties(5).Value = 1 Then
       Call PrintLiquidDensityToFile
    End If

    If frmPrint!chkProperties(6).Value = 1 Then
       Call PrintMolarVolumeOpTToFile
    End If

    If frmPrint!chkProperties(7).Value = 1 Then
       Call PrintMolarVolumeNBPToFile
    End If

    If frmPrint!chkProperties(8).Value = 1 Then
       Call PrintRefractiveIndexToFile
    End If

    If frmPrint!chkProperties(9).Value = 1 Then
       Call PrintAqueousSolubilityToFile
    End If

    If frmPrint!chkProperties(10).Value = 1 Then
       Call PrintOctWaterPartCoeffToFile
    End If

    If frmPrint!chkProperties(11).Value = 1 Then
       Call PrintLiquidDiffusivityToFile
    End If

    If frmPrint!chkProperties(12).Value = 1 Then
       Call PrintGasDiffusivityToFile
    End If

End Sub

Private Sub PrintChosenPropertiesToPrinter()

    If frmPrint!chkProperties(0).Value = 1 Then
       Call PrintVaporPressurePrinter
    End If

    If frmPrint!chkProperties(1).Value = 1 Then
       Call PrintActivityCoefficientPrinter
    End If

    If frmPrint!chkProperties(2).Value = 1 Then
       Call PrintHenrysConstantPrinter
    End If

    If frmPrint!chkProperties(3).Value = 1 Then
       Call PrintMolecularWeightPrinter
    End If

    If frmPrint!chkProperties(4).Value = 1 Then
       Call PrintBoilingPointPrinter
    End If

    If frmPrint!chkProperties(5).Value = 1 Then
       Call PrintLiquidDensityPrinter
    End If

    If frmPrint!chkProperties(6).Value = 1 Then
       Call PrintMolarVolumeOpTPrinter
    End If

    If frmPrint!chkProperties(7).Value = 1 Then
       Call PrintMolarVolumeNBPPrinter
    End If

    If frmPrint!chkProperties(8).Value = 1 Then
       Call PrintRefractiveIndexPrinter
    End If

    If frmPrint!chkProperties(9).Value = 1 Then
       Call PrintAqueousSolubilityPrinter
    End If

    If frmPrint!chkProperties(10).Value = 1 Then
       Call PrintOctWaterPartCoeffPrinter
    End If

    If frmPrint!chkProperties(11).Value = 1 Then
       Call PrintLiquidDiffusivityPrinter
    End If

    If frmPrint!chkProperties(12).Value = 1 Then
       Call PrintGasDiffusivityPrinter
    End If

End Sub

Private Sub PrintGasDiffusivityPrinter()
    Dim ValueString As String

    Select Case frmPrint!cboPropertyDescription.ListIndex
       Case 0   'Print Selected Value Only
          If HaveProperty(GAS_DIFFUSIVITY) Then
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = Format$(phprop.GasDiffusivity.CurrentSelection.Value, GetTheFormat(phprop.GasDiffusivity.CurrentSelection.Value))
             Printer.Print "Gas Diffusivity"; Tab(TABVALUE); ValueString; Tab(TABUNITS); Units(GAS_DIFFUSIVITY); Tab(TABSOURCE); GetSource(phprop.GasDiffusivity.CurrentSelection.Source)
          Else
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = "Not Available"
             Printer.Print "Gas Diffusivity"; Tab(TABVALUE); ValueString
          End If
       Case 1   'Print Full Description of Gas Diffusivity
          HeightGasDiffusivity = 0
          Printer.FontSize = 12
          PrintMsg = ""
          HeightGasDiffusivity = HeightGasDiffusivity + NUMLINES_PROPERTY_NAME * Printer.TextHeight(PrintMsg)
          Printer.FontSize = 10
          HeightGasDiffusivity = HeightGasDiffusivity + NUMLINES_GAS_DIFFUSIVITY * Printer.TextHeight(PrintMsg)
          TotalHeightThisPage = Printer.CurrentY
          If TotalHeightThisPage + HeightGasDiffusivity + BOTTOM_MARGIN_SAFETY_FACTOR > Printer.Height Then
             Printer.NewPage
             Call PrintTitleContinuation
          End If

          Printer.FontBold = True
          Printer.FontSize = 12
          Printer.Print "Property:  GAS DIFFUSIVITY"
          Printer.Print
          Printer.FontBold = False
          Printer.FontUnderline = True
          Printer.FontSize = 10
          Printer.Print Tab(TABFULLSOURCE); "Source:"; Tab(TABFULLVALUE); "Value:"; Tab(TABFULLUNITS); "Units:"; Tab(TABFULLTEMPERATURE); "Temp.:"; Tab(TABFULLCODE); "Code:"
          Printer.FontUnderline = False
          Printer.Print
          Call FullyPrintGasDiffusivityToPrinter
          Printer.Print
          Printer.Print

    End Select

End Sub

Private Sub PrintGasDiffusivityToFile()
    Dim ValueString As String

    Select Case frmPrint!cboPropertyDescription.ListIndex
       Case 0   'Print Selected Value Only
          If HaveProperty(GAS_DIFFUSIVITY) Then
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = Format$(phprop.GasDiffusivity.CurrentSelection.Value, GetTheFormat(phprop.GasDiffusivity.CurrentSelection.Value))
             Print #1, "Gas Diffusivity"; Tab(TABVALUE); ValueString; Tab(TABUNITS); Units(GAS_DIFFUSIVITY); Tab(TABSOURCE); GetSource(phprop.GasDiffusivity.CurrentSelection.Source)
          Else
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = "Not Available"
             Print #1, "Gas Diffusivity"; Tab(TABVALUE); ValueString
          End If
       Case 1   'Print Full Description of Gas Diffusivity
          Print #1, "Property:  GAS DIFFUSIVITY"
          Print #1,
          Print #1, Tab(TABFULLSOURCE); "Source:"; Tab(TABFULLVALUE); "Value:"; Tab(TABFULLUNITS); "Units:"; Tab(TABFULLTEMPERATURE); "Temperature:"; Tab(TABFULLCODE); "Code:"
          Print #1,
          Call FullyPrintGasDiffusivityToFile
          Print #1,
          Print #1,

    End Select

End Sub

Private Sub PrintHenrysConstantPrinter()
    Dim ValueString As String


    Select Case frmPrint!cboPropertyDescription.ListIndex
       Case 0   'Print Selected Value Only
          If HaveProperty(HENRYS_CONSTANT) Then
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = Format$(phprop.HenrysConstant.CurrentSelection.Value, GetTheFormat(phprop.HenrysConstant.CurrentSelection.Value))
             Printer.Print "Henry's Constant"; Tab(TABVALUE); ValueString; Tab(TABUNITS); Units(HENRYS_CONSTANT); Tab(TABSOURCE); GetSource(phprop.HenrysConstant.CurrentSelection.Source)
          Else
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = "Not Available"
             Printer.Print "Henry's Constant"; Tab(TABVALUE); ValueString
          End If
       Case 1   'Print Full Description of Henry's Constant
          HeightHenrysConstant = 0
          Printer.FontSize = 12
          PrintMsg = ""
          HeightHenrysConstant = HeightHenrysConstant + NUMLINES_PROPERTY_NAME * Printer.TextHeight(PrintMsg)
          Printer.FontSize = 10
          HeightHenrysConstant = HeightHenrysConstant + NUMLINES_HENRYS_CONSTANT * Printer.TextHeight(PrintMsg)

          If phprop.HenrysConstant.NumberOfDatabaseHenrysConstants > 1 Then 'Account for more than one Henry's constant in database in height determination
             HeightHenrysConstant = HeightHenrysConstant + 2 * (phprop.HenrysConstant.NumberOfDatabaseHenrysConstants - 1) * Printer.TextHeight(PrintMsg)
          End If
          TotalHeightThisPage = Printer.CurrentY
          If TotalHeightThisPage + HeightHenrysConstant + BOTTOM_MARGIN_SAFETY_FACTOR > Printer.Height Then
             Printer.NewPage
             Call PrintTitleContinuation
          End If

          Printer.FontBold = True
          Printer.FontSize = 12
          Printer.Print "Property:  HENRY'S CONSTANT"
          Printer.Print
          Printer.FontBold = False
          Printer.FontUnderline = True
          Printer.FontSize = 10
          Printer.Print Tab(TABFULLSOURCE); "Source:"; Tab(TABFULLVALUE); "Value:"; Tab(TABFULLUNITS); "Units:"; Tab(TABFULLTEMPERATURE); "Temp.:"; Tab(TABFULLCODE); "Code:"
          Printer.FontUnderline = False
          Printer.Print
          Call FullyPrintHenrysConstantToPrinter
          Printer.Print
          Printer.Print

    End Select

End Sub

Private Sub PrintHenrysConstantToFile()
    Dim ValueString As String

    Select Case frmPrint!cboPropertyDescription.ListIndex
       Case 0   'Print Selected Value Only
          If HaveProperty(HENRYS_CONSTANT) Then
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = Format$(phprop.HenrysConstant.CurrentSelection.Value, GetTheFormat(phprop.HenrysConstant.CurrentSelection.Value))
             Print #1, "Henry's Constant"; Tab(TABVALUE); ValueString; Tab(TABUNITS); Units(HENRYS_CONSTANT); Tab(TABSOURCE); GetSource(phprop.HenrysConstant.CurrentSelection.Source)
          Else
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = "Not Available"
             Print #1, "Henry's Constant"; Tab(TABVALUE); ValueString
          End If
       Case 1   'Print Full Description of Henry's Constant
          Print #1, "Property:  HENRY'S CONSTANT"
          Print #1,
          Print #1, Tab(TABFULLSOURCE); "Source:"; Tab(TABFULLVALUE); "Value:"; Tab(TABFULLUNITS); "Units:"; Tab(TABFULLTEMPERATURE); "Temp.:"; Tab(TABFULLCODE); "Code:"
          Print #1,
          Call FullyPrintHenrysConstantToFile
          Print #1,
          Print #1,

    End Select

End Sub

Private Sub PrintLiquidDensityPrinter()
    Dim ValueString As String

    Select Case frmPrint!cboPropertyDescription.ListIndex
       Case 0   'Print Selected Value Only
          If HaveProperty(LIQUID_DENSITY) Then
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = Format$(phprop.LiquidDensity.CurrentSelection.Value, GetTheFormat(phprop.LiquidDensity.CurrentSelection.Value))
             Printer.Print "Liquid Density"; Tab(TABVALUE); ValueString; Tab(TABUNITS); Units(LIQUID_DENSITY); Tab(TABSOURCE); GetSource(phprop.LiquidDensity.CurrentSelection.Source)
          Else
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = "Not Available"
             Printer.Print "Liquid Density"; Tab(TABVALUE); ValueString
          End If
       Case 1   'Print Full Description of Liquid Density
          HeightLiquidDensity = 0
          Printer.FontSize = 12
          PrintMsg = ""
          HeightLiquidDensity = HeightLiquidDensity + NUMLINES_PROPERTY_NAME * Printer.TextHeight(PrintMsg)
          Printer.FontSize = 10
          HeightLiquidDensity = HeightLiquidDensity + NUMLINES_LIQUID_DENSITY * Printer.TextHeight(PrintMsg)
          TotalHeightThisPage = Printer.CurrentY
          If TotalHeightThisPage + HeightLiquidDensity + BOTTOM_MARGIN_SAFETY_FACTOR > Printer.Height Then
             Printer.NewPage
             Call PrintTitleContinuation
          End If

          Printer.FontBold = True
          Printer.FontSize = 12
          Printer.Print "Property:  LIQUID DENSITY"
          Printer.Print
          Printer.FontBold = False
          Printer.FontUnderline = True
          Printer.FontSize = 10
          Printer.Print Tab(TABFULLSOURCE); "Source:"; Tab(TABFULLVALUE); "Value:"; Tab(TABFULLUNITS); "Units:"; Tab(TABFULLTEMPERATURE); "Temp.:"; Tab(TABFULLCODE); "Code:"
          Printer.FontUnderline = False
          Printer.Print
          Call FullyPrintLiquidDensityToPrinter
          Printer.Print
          Printer.Print

    End Select

End Sub

Private Sub PrintLiquidDensityToFile()
    Dim ValueString As String

    Select Case frmPrint!cboPropertyDescription.ListIndex
       Case 0   'Print Selected Value Only
          If HaveProperty(LIQUID_DENSITY) Then
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = Format$(phprop.LiquidDensity.CurrentSelection.Value, GetTheFormat(phprop.LiquidDensity.CurrentSelection.Value))
             Print #1, "Liquid Density"; Tab(TABVALUE); ValueString; Tab(TABUNITS); Units(LIQUID_DENSITY); Tab(TABSOURCE); GetSource(phprop.LiquidDensity.CurrentSelection.Source)
          Else
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = "Not Available"
             Print #1, "Liquid Density"; Tab(TABVALUE); ValueString
          End If
       Case 1   'Print Full Description of Liquid Density
          Print #1, "Property:  LIQUID DENSITY"
          Print #1,
          Print #1, Tab(TABFULLSOURCE); "Source:"; Tab(TABFULLVALUE); "Value:"; Tab(TABFULLUNITS); "Units:"; Tab(TABFULLTEMPERATURE); "Temp.:"; Tab(TABFULLCODE); "Code:"
          Print #1,
          Call FullyPrintLiquidDensityToFile
          Print #1,
          Print #1,

    End Select

End Sub

Private Sub PrintLiquidDiffusivityPrinter()
    Dim ValueString As String

    Select Case frmPrint!cboPropertyDescription.ListIndex
       Case 0   'Print Selected Value Only
          If HaveProperty(LIQUID_DIFFUSIVITY) Then
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = Format$(phprop.LiquidDiffusivity.CurrentSelection.Value, GetTheFormat(phprop.LiquidDiffusivity.CurrentSelection.Value))
             Printer.Print "Liquid Diffusivity"; Tab(TABVALUE); ValueString; Tab(TABUNITS); Units(LIQUID_DIFFUSIVITY); Tab(TABSOURCE); GetSource(phprop.LiquidDiffusivity.CurrentSelection.Source)
          Else
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = "Not Available"
             Printer.Print "Liquid Diffusivity"; Tab(TABVALUE); ValueString
          End If
       Case 1   'Print Full Description of Liquid Diffusivity
          HeightLiquidDiffusivity = 0
          Printer.FontSize = 12
          PrintMsg = ""
          HeightLiquidDiffusivity = HeightLiquidDiffusivity + NUMLINES_PROPERTY_NAME * Printer.TextHeight(PrintMsg)
          Printer.FontSize = 10
          HeightLiquidDiffusivity = HeightLiquidDiffusivity + NUMLINES_LIQUID_DIFFUSIVITY * Printer.TextHeight(PrintMsg)
          TotalHeightThisPage = Printer.CurrentY
          If TotalHeightThisPage + HeightLiquidDiffusivity + BOTTOM_MARGIN_SAFETY_FACTOR > Printer.Height Then
             Printer.NewPage
             Call PrintTitleContinuation
          End If

          Printer.FontBold = True
          Printer.FontSize = 12
          Printer.Print "Property:  LIQUID DIFFUSIVITY"
          Printer.Print
          Printer.FontBold = False
          Printer.FontUnderline = True
          Printer.FontSize = 10
          Printer.Print Tab(TABFULLSOURCE); "Source:"; Tab(TABFULLVALUE); "Value:"; Tab(TABFULLUNITS); "Units:"; Tab(TABFULLTEMPERATURE); "Temp.:"; Tab(TABFULLCODE); "Code:"
          Printer.FontUnderline = False
          Printer.Print
          Call FullyPrintLiquidDiffusivityToPrinter
          Printer.Print
          Printer.Print

    End Select

End Sub

Private Sub PrintLiquidDiffusivityToFile()
    Dim ValueString As String

    Select Case frmPrint!cboPropertyDescription.ListIndex
       Case 0   'Print Selected Value Only
          If HaveProperty(LIQUID_DIFFUSIVITY) Then
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = Format$(phprop.LiquidDiffusivity.CurrentSelection.Value, GetTheFormat(phprop.LiquidDiffusivity.CurrentSelection.Value))
             Print #1, "Liquid Diffusivity"; Tab(TABVALUE); ValueString; Tab(TABUNITS); Units(LIQUID_DIFFUSIVITY); Tab(TABSOURCE); GetSource(phprop.LiquidDiffusivity.CurrentSelection.Source)
          Else
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = "Not Available"
             Print #1, "Liquid Diffusivity"; Tab(TABVALUE); ValueString
          End If
       Case 1   'Print Full Description of Liquid Diffusivity
          Print #1, "Property:  LIQUID DIFFUSIVITY"
          Print #1,
          Print #1, Tab(TABFULLSOURCE); "Source:"; Tab(TABFULLVALUE); "Value:"; Tab(TABFULLUNITS); "Units:"; Tab(TABFULLTEMPERATURE); "Temperature:"; Tab(TABFULLCODE); "Code:"
          Print #1,
          Call FullyPrintLiquidDiffusivityToFile
          Print #1,
          Print #1,

    End Select

End Sub

Private Sub PrintMolarVolumeNBPPrinter()
    Dim ValueString As String

    Select Case frmPrint!cboPropertyDescription.ListIndex
       Case 0   'Print Selected Value Only
          If HaveProperty(MOLAR_VOLUME_BOILING_POINT) Then
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = Format$(phprop.MolarVolume.BoilingPoint.CurrentSelection.Value, GetTheFormat(phprop.MolarVolume.BoilingPoint.CurrentSelection.Value))
             Printer.Print "Molar Volume @ NBP"; Tab(TABVALUE); ValueString; Tab(TABUNITS); Units(MOLAR_VOLUME_BOILING_POINT); Tab(TABSOURCE); GetSource(phprop.MolarVolume.BoilingPoint.CurrentSelection.Source)
          Else
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = "Not Available"
             Printer.Print "Molar Volume @ NBP"; Tab(TABVALUE); ValueString
          End If
       Case 1   'Print Full Description of Molar Volume at Normal Boiling Point
          HeightMolarVolumeNBP = 0
          Printer.FontSize = 12
          PrintMsg = ""
          HeightMolarVolumeNBP = HeightMolarVolumeNBP + NUMLINES_PROPERTY_NAME * Printer.TextHeight(PrintMsg)
          Printer.FontSize = 10
          HeightMolarVolumeNBP = HeightMolarVolumeNBP + NUMLINES_MOLAR_VOLUME_NBP * Printer.TextHeight(PrintMsg)
          TotalHeightThisPage = Printer.CurrentY
          If TotalHeightThisPage + HeightMolarVolumeNBP + BOTTOM_MARGIN_SAFETY_FACTOR > Printer.Height Then
             Printer.NewPage
             Call PrintTitleContinuation
          End If

          Printer.FontBold = True
          Printer.FontSize = 12
          Printer.Print "Property:  MOLAR VOLUME AT NORMAL BOILING POINT"
          Printer.Print
          Printer.FontBold = False
          Printer.FontUnderline = True
          Printer.FontSize = 10
          Printer.Print Tab(TABFULLSOURCE); "Source:"; Tab(TABFULLVALUE); "Value:"; Tab(TABFULLUNITS); "Units:"; Tab(TABFULLTEMPERATURE); "Temp.:"; Tab(TABFULLCODE); "Code:"
          Printer.FontUnderline = False
          Printer.Print
          Call FullyPrintMolarVolumeNBPToPrinter
          Printer.Print
          Printer.Print

    End Select

End Sub

Private Sub PrintMolarVolumeNBPToFile()
    Dim ValueString As String

    Select Case frmPrint!cboPropertyDescription.ListIndex
       Case 0   'Print Selected Value Only
          If HaveProperty(MOLAR_VOLUME_BOILING_POINT) Then
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = Format$(phprop.MolarVolume.BoilingPoint.CurrentSelection.Value, GetTheFormat(phprop.MolarVolume.BoilingPoint.CurrentSelection.Value))
             Print #1, "Molar Volume @ NBP"; Tab(TABVALUE); ValueString; Tab(TABUNITS); Units(MOLAR_VOLUME_BOILING_POINT); Tab(TABSOURCE); GetSource(phprop.MolarVolume.BoilingPoint.CurrentSelection.Source)
          Else
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = "Not Available"
             Print #1, "Molar Volume @ NBP"; Tab(TABVALUE); ValueString
          End If
       Case 1   'Print Full Description of Molar Volume at Normal Boiling Point
          Print #1, "Property:  MOLAR VOLUME AT THE NORMAL BOILING POINT"
          Print #1,
          Print #1, Tab(TABFULLSOURCE); "Source:"; Tab(TABFULLVALUE); "Value:"; Tab(TABFULLUNITS); "Units:"; Tab(TABFULLTEMPERATURE); "Temperature:"; Tab(TABFULLCODE); "Code:"
          Print #1,
          Call FullyPrintMolarVolumeNBPToFile
          Print #1,
          Print #1,

    End Select

End Sub

Private Sub PrintMolarVolumeOpTPrinter()
    Dim ValueString As String

    Select Case frmPrint!cboPropertyDescription.ListIndex
       Case 0   'Print Selected Value Only
          If HaveProperty(MOLAR_VOLUME_OPT) Then
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = Format$(phprop.MolarVolume.operatingT.CurrentSelection.Value, GetTheFormat(phprop.MolarVolume.operatingT.CurrentSelection.Value))
             Printer.Print "Molar Volume @ Operating T"; Tab(TABVALUE); ValueString; Tab(TABUNITS); Units(MOLAR_VOLUME_OPT); Tab(TABSOURCE); GetSource(phprop.MolarVolume.operatingT.CurrentSelection.Source)
          Else
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = "Not Available"
             Printer.Print "Molar Volume @ Operating T"; Tab(TABVALUE); ValueString
          End If
       Case 1   'Print Full Description of Molar Volume at Operating T
          HeightMolarVolumeOpT = 0
          Printer.FontSize = 12
          PrintMsg = ""
          HeightMolarVolumeOpT = HeightMolarVolumeOpT + NUMLINES_PROPERTY_NAME * Printer.TextHeight(PrintMsg)
          Printer.FontSize = 10
          HeightMolarVolumeOpT = HeightMolarVolumeOpT + NUMLINES_MOLAR_VOLUME_OPT * Printer.TextHeight(PrintMsg)
          TotalHeightThisPage = Printer.CurrentY
          If TotalHeightThisPage + HeightMolarVolumeOpT + BOTTOM_MARGIN_SAFETY_FACTOR > Printer.Height Then
             Printer.NewPage
             Call PrintTitleContinuation
          End If

          Printer.FontBold = True
          Printer.FontSize = 12
          Printer.Print "Property:  MOLAR VOLUME AT OPERATING TEMPERATURE"
          Printer.Print
          Printer.FontBold = False
          Printer.FontUnderline = True
          Printer.FontSize = 10
          Printer.Print Tab(TABFULLSOURCE); "Source:"; Tab(TABFULLVALUE); "Value:"; Tab(TABFULLUNITS); "Units:"; Tab(TABFULLTEMPERATURE); "Temp.:"; Tab(TABFULLCODE); "Code:"
          Printer.FontUnderline = False
          Printer.Print
          Call FullyPrintMolarVolumeOpTToPrinter
          Printer.Print
          Printer.Print

    End Select

End Sub

Private Sub PrintMolarVolumeOpTToFile()
    Dim ValueString As String

    Select Case frmPrint!cboPropertyDescription.ListIndex
       Case 0   'Print Selected Value Only
          If HaveProperty(MOLAR_VOLUME_OPT) Then
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = Format$(phprop.MolarVolume.operatingT.CurrentSelection.Value, GetTheFormat(phprop.MolarVolume.operatingT.CurrentSelection.Value))
             Print #1, "Molar Volume @ Operating T"; Tab(TABVALUE); ValueString; Tab(TABUNITS); Units(MOLAR_VOLUME_OPT); Tab(TABSOURCE); GetSource(phprop.MolarVolume.operatingT.CurrentSelection.Source)
          Else
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = "Not Available"
             Print #1, "Molar Volume @ Operating T"; Tab(TABVALUE); ValueString
          End If
       Case 1   'Print Full Description of Molar Volume at Operating T
          Print #1, "Property:  MOLAR VOLUME AT THE OPERATING TEMPERATURE"
          Print #1,
          Print #1, Tab(TABFULLSOURCE); "Source:"; Tab(TABFULLVALUE); "Value:"; Tab(TABFULLUNITS); "Units:"; Tab(TABFULLTEMPERATURE); "Temp.:"; Tab(TABFULLCODE); "Code:"
          Print #1,
          Call FullyPrintMolarVolumeOpTToFile
          Print #1,
          Print #1,

    End Select

End Sub

Private Sub PrintMolecularWeightPrinter()
    Dim ValueString As String

    Select Case frmPrint!cboPropertyDescription.ListIndex
       Case 0   'Print Selected Value Only
          If HaveProperty(MOLECULAR_WEIGHT) Then
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = Format$(phprop.MolecularWeight.CurrentSelection.Value, MOLECULAR_WEIGHT_FORMAT)
             Printer.Print "Molecular Weight"; Tab(TABVALUE); ValueString; Tab(TABUNITS); Units(MOLECULAR_WEIGHT); Tab(TABSOURCE); GetSource(phprop.MolecularWeight.CurrentSelection.Source)
          Else
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = "Not Available"
             Printer.Print "Molecular Weight"; Tab(TABVALUE); ValueString
          End If
       Case 1   'Print Full Description of Molecular Weight
          HeightMolecularWeight = 0
          Printer.FontSize = 12
          PrintMsg = ""
          HeightMolecularWeight = HeightMolecularWeight + NUMLINES_PROPERTY_NAME * Printer.TextHeight(PrintMsg)
          Printer.FontSize = 10
          HeightMolecularWeight = HeightMolecularWeight + NUMLINES_MOLECULAR_WEIGHT * Printer.TextHeight(PrintMsg)
          TotalHeightThisPage = Printer.CurrentY
          If TotalHeightThisPage + HeightMolecularWeight + BOTTOM_MARGIN_SAFETY_FACTOR > Printer.Height Then
             Printer.NewPage
             Call PrintTitleContinuation
          End If

          Printer.FontBold = True
          Printer.FontSize = 12
          Printer.Print "Property:  MOLECULAR WEIGHT"
          Printer.Print
          Printer.FontBold = False
          Printer.FontUnderline = True
          Printer.FontSize = 10
          Printer.Print Tab(TABFULLSOURCE); "Source:"; Tab(TABFULLVALUE); "Value:"; Tab(TABFULLUNITS); "Units:"; Tab(TABFULLCODE); "Code:"
          Printer.FontUnderline = False
          Printer.Print
          Call FullyPrintMolecularWeightToPrinter
          Printer.Print
          Printer.Print

    End Select

End Sub

Private Sub PrintMolecularWeightToFile()
    Dim ValueString As String

    Select Case frmPrint!cboPropertyDescription.ListIndex
       Case 0   'Print Selected Value Only
          If HaveProperty(MOLECULAR_WEIGHT) Then
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = Format$(phprop.MolecularWeight.CurrentSelection.Value, MOLECULAR_WEIGHT_FORMAT)
             Print #1, "Molecular Weight"; Tab(TABVALUE); ValueString; Tab(TABUNITS); Units(MOLECULAR_WEIGHT); Tab(TABSOURCE); GetSource(phprop.MolecularWeight.CurrentSelection.Source)
          Else
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = "Not Available"
             Print #1, "Molecular Weight"; Tab(TABVALUE); ValueString
          End If
       Case 1   'Print Full Description of Molecular Weight
          Print #1, "Property:  MOLECULAR WEIGHT"
          Print #1,
          Print #1, Tab(TABFULLSOURCE); "Source:"; Tab(TABFULLVALUE); "Value:"; Tab(TABFULLUNITS); "Units:"; Tab(TABFULLCODE); "Code:"
          Print #1,
          Call FullyPrintMolecularWeightToFile
          Print #1,
          Print #1,

    End Select

End Sub

Private Sub PrintOctWaterPartCoeffPrinter()
    Dim ValueString As String

    Select Case frmPrint!cboPropertyDescription.ListIndex
       Case 0   'Print Selected Value Only
          If HaveProperty(OCT_WATER_PART_COEFF) Then
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = Format$(phprop.OctWaterPartCoeff.CurrentSelection.Value, GetTheFormat(phprop.OctWaterPartCoeff.CurrentSelection.Value))
             Printer.Print "log Octanol Water Part. Coeff."; Tab(TABVALUE); ValueString; Tab(TABUNITS); Units(OCT_WATER_PART_COEFF); Tab(TABSOURCE); GetSource(phprop.OctWaterPartCoeff.CurrentSelection.Source)
          Else
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = "Not Available"
             Printer.Print "log Octanol Water Part. Coeff."; Tab(TABVALUE); ValueString
          End If
       Case 1   'Print Full Description of Octanol Water Partition Coefficient
          HeightOctWaterPartCoeff = 0
          Printer.FontSize = 12
          PrintMsg = ""
          HeightOctWaterPartCoeff = HeightOctWaterPartCoeff + NUMLINES_PROPERTY_NAME * Printer.TextHeight(PrintMsg)
          Printer.FontSize = 10
          HeightOctWaterPartCoeff = HeightOctWaterPartCoeff + NUMLINES_OCT_WATER_PART_COEFF * Printer.TextHeight(PrintMsg)
          TotalHeightThisPage = Printer.CurrentY
          If TotalHeightThisPage + HeightOctWaterPartCoeff + BOTTOM_MARGIN_SAFETY_FACTOR > Printer.Height Then
             Printer.NewPage
             Call PrintTitleContinuation
          End If

          Printer.FontBold = True
          Printer.FontSize = 12
          Printer.Print "Property:  OCTANOL WATER PARTITION COEFFICIENT"
          Printer.Print
          Printer.FontBold = False
          Printer.FontUnderline = True
          Printer.FontSize = 10
          Printer.Print Tab(TABFULLSOURCE); "Source:"; Tab(TABFULLVALUE); "Value:"; Tab(TABFULLUNITS); "Units:"; Tab(TABFULLTEMPERATURE); "Temp.:"; Tab(TABFULLCODE); "Code:"
          Printer.FontUnderline = False
          Printer.Print
          Call FullyPrintOctWaterPartCoeffToPrinter
          Printer.Print
          Printer.Print

    End Select

End Sub

Private Sub PrintOctWaterPartCoeffToFile()
    Dim ValueString As String

    Select Case frmPrint!cboPropertyDescription.ListIndex
       Case 0   'Print Selected Value Only
          If HaveProperty(OCT_WATER_PART_COEFF) Then
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = Format$(phprop.OctWaterPartCoeff.CurrentSelection.Value, GetTheFormat(phprop.OctWaterPartCoeff.CurrentSelection.Value))
             Print #1, "log Octanol Water Part. Coeff."; Tab(TABVALUE); ValueString; Tab(TABUNITS); Units(OCT_WATER_PART_COEFF); Tab(TABSOURCE); GetSource(phprop.OctWaterPartCoeff.CurrentSelection.Source)
          Else
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = "Not Available"
             Print #1, "log Octanol Water Part. Coeff."; Tab(TABVALUE); ValueString
          End If
       Case 1   'Print Full Description of Octanol Water Partition Coefficient
          Print #1, "Property:  OCTANOL WATER PARTITION COEFFICIENT"
          Print #1,
          Print #1, Tab(TABFULLSOURCE); "Source:"; Tab(TABFULLVALUE); "Value:"; Tab(TABFULLUNITS); "Units:"; Tab(TABFULLTEMPERATURE); "Temp.:"; Tab(TABFULLCODE); "Code:"
          Print #1,
          Call FullyPrintOctWaterPartCoeffToFile
          Print #1,
          Print #1,

    End Select

End Sub

Private Sub PrintOneContaminant()

       Printer.FontSize = 14
       Printer.FontBold = True
       Printer.Print phprop.CasNumber; " "; phprop.Name
       Printer.Print
       Printer.Print

       Select Case frmPrint!cboPropertyDescription.ListIndex
          Case 0   'Print Selected Value Only
             Printer.FontBold = False
             Printer.Print "Summary of Selected Values"
             Printer.Print
             Printer.Print
             Printer.FontSize = 10
             Printer.FontBold = True
             Printer.FontUnderline = True
             Printer.Print "Property:"; Tab(TABVALUE); "Value:"; Tab(TABUNITS); "Units:"; Tab(TABSOURCE); "Source:"
             Printer.Print
             Printer.FontUnderline = False
             Printer.FontBold = False
             Printer.FontSize = 10
             Call PrintOperatingPressure
             Call PrintOperatingTemperature
             Printer.Print

             If frmPrint!optPrintProperties(0).Value Then  'Print all properties
                Call PrintAllPropertiesToPrinter
             ElseIf frmPrint!optPrintProperties(1).Value Then  'Print chosen properties only
                Call PrintChosenPropertiesToPrinter
             End If
          Case 1   'Print Full Description of Properties
             Printer.FontBold = False
             Printer.Print "Full Description of Properties"
             Printer.Print
             Printer.Print
             HeightTitle = Printer.CurrentY
             Printer.FontSize = 10
             Printer.FontBold = True
             Printer.FontUnderline = True
             Printer.Print "Operating Conditions:"; Tab(TABVALUE); "Value:"; Tab(TABUNITS); "Units:"; Tab(TABSOURCE); "Source:"
             Printer.Print
             Printer.FontUnderline = False
             Printer.FontBold = False
             Printer.FontSize = 10
             Call PrintOperatingPressure
             Call PrintOperatingTemperature
             Printer.Print
             Printer.Print
             HeightOperatingConditions = Printer.CurrentY - HeightTitle

             If frmPrint!optPrintProperties(0).Value Then  'Print all properties
                Call PrintAllPropertiesToPrinter
             ElseIf frmPrint!optPrintProperties(1).Value Then  'Print chosen properties only
                Call PrintChosenPropertiesToPrinter
             End If

       End Select

End Sub

Private Sub PrintOneContaminantToFile()

       Print #1, phprop.CasNumber; " "; phprop.Name
       Print #1,
       Print #1,
       Print #1,

       Select Case frmPrint!cboPropertyDescription.ListIndex
          Case 0   'Print Selected Value Only
             Print #1, "Summary of Selected Values"
             Print #1,
             Print #1,
             Print #1, "Property:"; Tab(TABVALUE); "Value:"; Tab(TABUNITS); "Units:"; Tab(TABSOURCE); "Source:"
             Print #1,
             Call PrintOperatingPressureToFile
             Call PrintOperatingTemperatureToFile
             Print #1,

             If frmPrint!optPrintProperties(0).Value Then  'Print all properties
                Call PrintAllPropertiesToFile
             ElseIf frmPrint!optPrintProperties(1).Value Then  'Print chosen properties only
                Call PrintChosenPropertiesToFile
             End If
          Case 1   'Print Full Description of Properties
             Print #1, "Full Description of Properties"
             Print #1,
             Print #1,
             Print #1, "Property:"; Tab(TABVALUE); "Value:"; Tab(TABUNITS); "Units:"; Tab(TABSOURCE); "Source:"
             Print #1,
             Call PrintOperatingPressureToFile
             Call PrintOperatingTemperatureToFile
             Print #1,
             Print #1,

             If frmPrint!optPrintProperties(0).Value Then  'Print all properties
                Call PrintAllPropertiesToFile
             ElseIf frmPrint!optPrintProperties(1).Value Then  'Print chosen properties only
                Call PrintChosenPropertiesToFile
             End If

       End Select

End Sub

Private Sub PrintOperatingPressure()
    Dim ValueString As String

    ValueString = Space$(VALUELENGTH)
    RSet ValueString = Format$(phprop.OperatingPressure, GetTheFormat(phprop.OperatingPressure))
    Printer.Print "Operating Pressure"; Tab(TABVALUE); ValueString; Tab(TABUNITS); Units(OPERATING_PRESSURE); Tab(TABSOURCE); "User Input"

End Sub

Private Sub PrintOperatingPressureToFile()
    Dim ValueString As String

    ValueString = Space$(VALUELENGTH)
    RSet ValueString = Format$(phprop.OperatingPressure, GetTheFormat(phprop.OperatingPressure))
    Print #1, "Operating Pressure"; Tab(TABVALUE); ValueString; Tab(TABUNITS); Units(OPERATING_PRESSURE); Tab(TABSOURCE); "User Input"

End Sub

Private Sub PrintOperatingTemperature()
    Dim ValueString As String

    ValueString = Space$(VALUELENGTH)
    RSet ValueString = Format$(phprop.OperatingTemperature, GetTheFormat(phprop.OperatingTemperature))
    Printer.Print "Operating Temperature"; Tab(TABVALUE); ValueString; Tab(TABUNITS); Units(OPERATING_TEMPERATURE); Tab(TABSOURCE); "User Input"

End Sub

Private Sub PrintOperatingTemperatureToFile()
    Dim ValueString As String

    ValueString = Space$(VALUELENGTH)
    RSet ValueString = Format$(phprop.OperatingTemperature, GetTheFormat(phprop.OperatingTemperature))
    Print #1, "Operating Temperature"; Tab(TABVALUE); ValueString; Tab(TABUNITS); Units(OPERATING_TEMPERATURE); Tab(TABSOURCE); "User Input"

End Sub

Private Sub PrintRefractiveIndexPrinter()
    Dim ValueString As String

    Select Case frmPrint!cboPropertyDescription.ListIndex
       Case 0   'Print Selected Value Only
          If HaveProperty(REFRACTIVE_INDEX) Then
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = Format$(phprop.RefractiveIndex.CurrentSelection.Value, REFRACTIVE_INDEX_FORMAT)
             Printer.Print "Refractive Index"; Tab(TABVALUE); ValueString; Tab(TABUNITS); Units(REFRACTIVE_INDEX); Tab(TABSOURCE); GetSource(phprop.RefractiveIndex.CurrentSelection.Source)
          Else
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = "Not Available"
             Printer.Print "Refractive Index"; Tab(TABVALUE); ValueString
          End If
       Case 1   'Print Full Description of Refractive Index
          HeightRefractiveIndex = 0
          Printer.FontSize = 12
          PrintMsg = ""
          HeightRefractiveIndex = HeightRefractiveIndex + NUMLINES_PROPERTY_NAME * Printer.TextHeight(PrintMsg)
          Printer.FontSize = 10
          HeightRefractiveIndex = HeightRefractiveIndex + NUMLINES_REFRACTIVE_INDEX * Printer.TextHeight(PrintMsg)
          TotalHeightThisPage = Printer.CurrentY
          If TotalHeightThisPage + HeightRefractiveIndex + BOTTOM_MARGIN_SAFETY_FACTOR > Printer.Height Then
             Printer.NewPage
             Call PrintTitleContinuation
          End If

          Printer.FontBold = True
          Printer.FontSize = 12
          Printer.Print "Property:  REFRACTIVE INDEX"
          Printer.Print
          Printer.FontBold = False
          Printer.FontUnderline = True
          Printer.FontSize = 10
          Printer.Print Tab(TABFULLSOURCE); "Source:"; Tab(TABFULLVALUE); "Value:"; Tab(TABFULLUNITS); "Units:"; Tab(TABFULLCODE); "Code:"
          Printer.FontUnderline = False
          Printer.Print
          Call FullyPrintRefractiveIndexToPrinter
          Printer.Print
          Printer.Print

    End Select

End Sub

Private Sub PrintRefractiveIndexToFile()
    Dim ValueString As String

    Select Case frmPrint!cboPropertyDescription.ListIndex
       Case 0   'Print Selected Value Only
          If HaveProperty(REFRACTIVE_INDEX) Then
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = Format$(phprop.RefractiveIndex.CurrentSelection.Value, REFRACTIVE_INDEX_FORMAT)
             Print #1, "Refractive Index"; Tab(TABVALUE); ValueString; Tab(TABUNITS); Units(REFRACTIVE_INDEX); Tab(TABSOURCE); GetSource(phprop.RefractiveIndex.CurrentSelection.Source)
          Else
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = "Not Available"
             Print #1, "Refractive Index"; Tab(TABVALUE); ValueString
          End If
       Case 1   'Print Full Description of Refractive Index
          Print #1, "Property:  REFRACTIVE INDEX"
          Print #1,
          Print #1, Tab(TABFULLSOURCE); "Source:"; Tab(TABFULLVALUE); "Value:"; Tab(TABFULLUNITS); "Units:"; Tab(TABFULLCODE); "Code:"
          Print #1,
          Call FullyPrintRefractiveIndexToFile
          Print #1,
          Print #1,

    End Select

End Sub

Private Sub PrintTheCodes(code As Long)

    Printer.Print
    Printer.FontUnderline = True
    Printer.Print Tab(TABCODE); "Code:"; Tab(TABCODEDESCRIPTION); "Description:"
    Printer.FontUnderline = False
    Printer.Print
    Printer.Print Tab(TABCODE); code;
    Select Case code
       Case 1   'UNIFAC Parameter Set = Original UNIFAC VLE
            Printer.Print Tab(TABCODEDESCRIPTION); "UNIFAC Parameter Set = Original UNIFAC VLE"
       Case 2   'UNIFAC Parameter Set = UNIFAC LLE
            Printer.Print Tab(TABCODEDESCRIPTION); "UNIFAC Parameter Set = UNIFAC LLE"
       Case 3   'UNIFAC Parameter Set = Environmental VLE
            Printer.Print Tab(TABCODEDESCRIPTION); "UNIFAC Parameter Set = Environmental VLE"
'       Case 0   'UNIFAC Calculation Not Possible
'            Printer.Print Tab(TABCODEDESCRIPTION); "UNIFAC Calculation Not Possible"
    End Select
End Sub

Private Sub PrintTheCodesToFile(code As Long)

    Print #1,
    Print #1,
    Printer.FontUnderline = True
    Print #1, Tab(TABCODE); "Code:"; Tab(TABCODEDESCRIPTION); "Description:"
    Printer.FontUnderline = False
    Print #1,
    Print #1, Tab(TABCODE); code;
    Select Case code
       Case 1   'UNIFAC Parameter Set = Original UNIFAC VLE
            Print #1, Tab(TABCODEDESCRIPTION); "UNIFAC Parameter Set = Original UNIFAC VLE"
       Case 2   'UNIFAC Parameter Set = UNIFAC LLE
            Print #1, Tab(TABCODEDESCRIPTION); "UNIFAC Parameter Set = UNIFAC LLE"
       Case 3   'UNIFAC Parameter Set = Environmental VLE
            Print #1, Tab(TABCODEDESCRIPTION); "UNIFAC Parameter Set = Environmental VLE"
'       Case 0   'UNIFAC Calculation Not Possible
'            Print #1, Tab(TABCODEDESCRIPTION); "UNIFAC Calculation Not Possible"
    End Select

End Sub

Private Sub PrintTheErrors(Header As Integer, code As Long)
    
    Dim cut As String
    
    If code = 0 Then Exit Sub

    If Header = 1 Then
       Header = 0
       Printer.Print
       Printer.FontUnderline = True
       Printer.Print Tab(TABCODE); "Code:"; Tab(TABCODEDESCRIPTION); "Description:"
       Printer.FontUnderline = False
       Printer.Print
    End If
    
    Printer.Print Tab(TABCODE); code;
    curri = 1
    currl = 1
' This routine breaks up long error messages
iter1:
    i = curri + 59
    If Len(ErrorMsg(code)) <= i Then
       cut = Mid$(ErrorMsg(code), curri, i)
       Printer.Print Tab(TABCODEDESCRIPTION); Trim$(cut)
       Exit Sub
    End If
    If Mid$(ErrorMsg(code), i, 1) = " " Then
        cut = Mid$(ErrorMsg(code), curri, i - curri)
        curri = i
        Printer.Print Tab(TABCODEDESCRIPTION); Trim$(cut)
        GoTo iter1:
    Else
        curri = i
iter2:
        curri = curri + 1
        If Mid$(ErrorMsg(code), curri, 1) = "." Then
            cut = Mid$(ErrorMsg(code), currl, curri - currl + 1)
            Printer.Print Tab(TABCODEDESCRIPTION); Trim$(cut)
            Exit Sub
        End If
        If Mid$(ErrorMsg(code), curri, 1) = " " Then
            cut = Mid$(ErrorMsg(code), currl, curri - currl)
            currl = curri
            Printer.Print Tab(TABCODEDESCRIPTION); Trim$(cut)
            GoTo iter1:
        End If
        GoTo iter2:
    End If

End Sub

Private Sub PrintTheErrorsToFile(Header As Integer, code As Long)
    
    Dim cut As String

    If code = 0 Then Exit Sub

    If Header = 1 Then
       Header = 0
       Print #1,
       Print #1,
       Printer.FontUnderline = True
       Print #1, Tab(TABCODE); "Code:"; Tab(TABCODEDESCRIPTION); "Description:"
       Printer.FontUnderline = False
       Print #1,
    End If

    Print #1, Tab(TABCODE); code;
    curri = 1
    currl = 1
iterf1:
    i = curri + 59
    If Len(ErrorMsg(code)) < i Then
        cut$ = Mid$(ErrorMsg(code), curri, i)
        Print #1, Tab(TABCODEDESCRIPTION); Trim$(cut)
        Exit Sub
    End If
    If Mid$(ErrorMsg(code), i, 1) = " " Then
        cut$ = Mid$(ErrorMsg(code), curri, i)
        curri = i
    Else
        curri = i
iterf2:
        curri = curri + 1
        '////////////////////////////////////////////////////////////////////
        '////  CODE ADDED BY ERIC OMAN (25-MAR-1999) BEGINS:
        If (curri > Len(ErrorMsg(code))) Then
            '
            ' IN CASE SOMETHING WEIRD HAPPENS, PRINT THE ERROR
            ' MESSAGE WITHOUT EXTRA FORMATTING, AND EXIT OUT.
            '
            Print #1, ErrorMsg(code)
            Exit Sub
        End If
        '////  CODE ADDED BY ERIC OMAN (25-MAR-1999) ENDS.
        '////////////////////////////////////////////////////////////////////
        If Mid$(ErrorMsg(code), curri, 1) = "." Then
            cut = Mid$(ErrorMsg(code), currl, curri - currl + 1)
            Print #1, Tab(TABCODEDESCRIPTION); Trim$(cut)
            Exit Sub
        End If
        If Mid$(ErrorMsg(code), curri, 1) = " " Then
            cut$ = Mid$(ErrorMsg(code), currl, curri)
            currl = curri
            Print #1, Tab(TABCODEDESCRIPTION); Trim$(cut)
            GoTo iterf1:
        End If
        GoTo iterf2:
    End If

End Sub

Private Sub PrintTitleContinuation()

       Printer.FontSize = 14
       Printer.FontBold = True
       Printer.Print phprop.CasNumber; " "; Trim$(phprop.Name);
       Printer.FontBold = False
       Printer.Print " (continued)"
       Printer.Print
       Printer.Print
       Printer.FontBold = False
       Printer.Print "Full Description of Properties"
       Printer.Print
       Printer.Print
       HeightTitle = Printer.CurrentY
       Printer.FontSize = 10
       Printer.FontBold = True
       Printer.FontUnderline = True
       Printer.Print "Operating Conditions:"; Tab(TABVALUE); "Value:"; Tab(TABUNITS); "Units:"; Tab(TABSOURCE); "Source:"
       Printer.Print
       Printer.FontUnderline = False
       Printer.FontBold = False
       Printer.FontSize = 10
       Call PrintOperatingPressure
       Call PrintOperatingTemperature
       Printer.Print
       Printer.Print
       HeightOperatingConditions = Printer.CurrentY - HeightTitle

End Sub

Private Sub PrintToFile()
    Dim CurrentlySelectedContaminant As Integer
    Dim i As Integer, J As Integer
    Dim ChosenAtLeastOneContaminantProperty As Integer
    Dim ChosenAtLeastOneAirWaterProperty As Integer
    Dim EnglishValue As Double
    Static temphdbttmp(20) As Double
    Static temphunttmp(20) As Double

On Error GoTo error_printtofile

    ChosenAtLeastOneContaminantProperty = False
    ChosenAtLeastOneAirWaterProperty = False
    If optPrintProperties(0).Value Then
       ChosenAtLeastOneContaminantProperty = True
       ChosenAtLeastOneAirWaterProperty = True
    Else
       For i = 0 To 12
           If (frmPrint!chkProperties(i).Value = 1) Then
              ChosenAtLeastOneContaminantProperty = True
              Exit For
           End If
       Next i

       For i = 13 To 17
           If (frmPrint!chkProperties(i).Value = 1) Then
              ChosenAtLeastOneAirWaterProperty = True
              Exit For
           End If
       Next i
    End If

    If frmPrint!optPrintContaminants(0).Value Then   'Print All Contaminants
       If ChosenAtLeastOneContaminantProperty Then
          CurrentlySelectedContaminant = contam_prop_form!cboSelectContaminant.ListIndex + 1
          For i = 1 To NumSelectedChemicals
              phprop = PropContaminant(i)
              
              'If English units are desired convert them here
              If cboUnits.ListIndex = 1 Then
              
                 temppress = phprop.OperatingPressure
                 Call PRESSCNV(EnglishValue, phprop.OperatingPressure)
                 phprop.OperatingPressure = EnglishValue
              
                 tempt = phprop.OperatingTemperature
                 Call TEMPCNV(EnglishValue, phprop.OperatingTemperature)
                 phprop.OperatingTemperature = EnglishValue
                 
                 'Change all temperatures
                 tempvptmp = phprop.VaporPressure.database.temperature
                 Call TEMPCNV(EnglishValue, phprop.VaporPressure.database.temperature)
                 phprop.VaporPressure.database.temperature = EnglishValue
                 
                 tempvptmpi = phprop.VaporPressure.input.temperature
                 Call TEMPCNV(EnglishValue, phprop.VaporPressure.input.temperature)
                 phprop.VaporPressure.input.temperature = EnglishValue
                 
                 tempactmp = phprop.ActivityCoefficient.UNIFAC.temperature
                 Call TEMPCNV(EnglishValue, phprop.ActivityCoefficient.UNIFAC.temperature)
                 phprop.ActivityCoefficient.UNIFAC.temperature = EnglishValue
                 
                 temphregtmp = phprop.HenrysConstant.regress.temperature
                 Call TEMPCNV(EnglishValue, phprop.HenrysConstant.regress.temperature)
                 phprop.HenrysConstant.regress.temperature = EnglishValue
                 
                 temphfittmp = phprop.HenrysConstant.fit.UNIFAC.temperature
                 Call TEMPCNV(EnglishValue, phprop.HenrysConstant.fit.UNIFAC.temperature)
                 phprop.HenrysConstant.fit.UNIFAC.temperature = EnglishValue
                 
                 temphopttmp = phprop.HenrysConstant.operatingT.UNIFAC.temperature
                 Call TEMPCNV(EnglishValue, phprop.HenrysConstant.operatingT.UNIFAC.temperature)
                 phprop.HenrysConstant.operatingT.UNIFAC.temperature = EnglishValue
       
                 For J = 1 To phprop.HenrysConstant.NumberOfDatabaseHenrysConstants
                    temphdbttmp(J) = phprop.HenrysConstant.database(J).temperature
                    temphunttmp(J) = phprop.HenrysConstant.UNIFAC(J).temperature
                    Call TEMPCNV(EnglishValue, phprop.HenrysConstant.database(J).temperature)
                    phprop.HenrysConstant.database(J).temperature = EnglishValue
                    Call TEMPCNV(EnglishValue, phprop.HenrysConstant.UNIFAC(J).temperature)
                    phprop.HenrysConstant.UNIFAC(J).temperature = EnglishValue
                 Next J
                 
                 temphtmpi = phprop.HenrysConstant.input.temperature
                 Call TEMPCNV(EnglishValue, phprop.HenrysConstant.input.temperature)
                 phprop.HenrysConstant.input.temperature = EnglishValue
                 
                 templdtmp = phprop.LiquidDensity.database.temperature
                 Call TEMPCNV(EnglishValue, phprop.LiquidDensity.database.temperature)
                 phprop.LiquidDensity.database.temperature = EnglishValue
                 
                 templdutmp = phprop.LiquidDensity.UNIFAC.temperature
                 Call TEMPCNV(EnglishValue, phprop.LiquidDensity.UNIFAC.temperature)
                 phprop.LiquidDensity.UNIFAC.temperature = EnglishValue
                 
                 templdtmpi = phprop.LiquidDensity.input.temperature
                 Call TEMPCNV(EnglishValue, phprop.LiquidDensity.input.temperature)
                 phprop.LiquidDensity.input.temperature = EnglishValue
                 
                 tempmvopttmp = phprop.MolarVolume.operatingT.database.temperature
                 Call TEMPCNV(EnglishValue, phprop.MolarVolume.operatingT.database.temperature)
                 phprop.MolarVolume.operatingT.database.temperature = EnglishValue
                 
                 tempmvoptutmp = phprop.MolarVolume.operatingT.UNIFAC.temperature
                 Call TEMPCNV(EnglishValue, phprop.MolarVolume.operatingT.UNIFAC.temperature)
                 phprop.MolarVolume.operatingT.UNIFAC.temperature = EnglishValue
                 
                 tempmvopttmpi = phprop.MolarVolume.operatingT.input.temperature
                 Call TEMPCNV(EnglishValue, phprop.MolarVolume.operatingT.input.temperature)
                 phprop.MolarVolume.operatingT.input.temperature = EnglishValue
                 
                 tempmvtmp = phprop.MolarVolume.BoilingPoint.UNIFAC.temperature
                 Call TEMPCNV(EnglishValue, phprop.MolarVolume.BoilingPoint.UNIFAC.temperature)
                 phprop.MolarVolume.BoilingPoint.UNIFAC.temperature = EnglishValue
                 
                 tempmvtmpi = phprop.MolarVolume.BoilingPoint.input.temperature
                 Call TEMPCNV(EnglishValue, phprop.MolarVolume.BoilingPoint.input.temperature)
                 phprop.MolarVolume.BoilingPoint.input.temperature = EnglishValue
                 
                 tempaqfittmp = phprop.AqueousSolubility.fit.UNIFAC.temperature
                 Call TEMPCNV(EnglishValue, phprop.AqueousSolubility.fit.UNIFAC.temperature)
                 phprop.AqueousSolubility.fit.UNIFAC.temperature = EnglishValue
                 
                 tempaqopttmp = phprop.AqueousSolubility.operatingT.UNIFAC.temperature
                 Call TEMPCNV(EnglishValue, phprop.AqueousSolubility.operatingT.UNIFAC.temperature)
                 phprop.AqueousSolubility.operatingT.UNIFAC.temperature = EnglishValue
       
                 tempaqdbtmp = phprop.AqueousSolubility.database.temperature
                 Call TEMPCNV(EnglishValue, phprop.AqueousSolubility.database.temperature)
                 phprop.AqueousSolubility.database.temperature = EnglishValue
                 
                 tempaquntmp = phprop.AqueousSolubility.UNIFAC.temperature
                 Call TEMPCNV(EnglishValue, phprop.AqueousSolubility.UNIFAC.temperature)
                 phprop.AqueousSolubility.UNIFAC.temperature = EnglishValue
                 
                 tempaqtmpi = phprop.AqueousSolubility.input.temperature
                 Call TEMPCNV(EnglishValue, phprop.AqueousSolubility.input.temperature)
                 phprop.AqueousSolubility.input.temperature = EnglishValue
                 
                 tempoctopttmp = phprop.OctWaterPartCoeff.operatingT.UNIFAC.temperature
                 Call TEMPCNV(EnglishValue, phprop.OctWaterPartCoeff.operatingT.UNIFAC.temperature)
                 phprop.OctWaterPartCoeff.operatingT.UNIFAC.temperature = EnglishValue
              
                 tempoctdbtmp = phprop.OctWaterPartCoeff.database.temperature
                 Call TEMPCNV(EnglishValue, phprop.OctWaterPartCoeff.database.temperature)
                 phprop.OctWaterPartCoeff.database.temperature = EnglishValue

                 tempoctuntmp = phprop.OctWaterPartCoeff.databaseT.UNIFAC.temperature
                 Call TEMPCNV(EnglishValue, phprop.OctWaterPartCoeff.databaseT.UNIFAC.temperature)
                 phprop.OctWaterPartCoeff.databaseT.UNIFAC.temperature = EnglishValue

                 tempocttmpi = phprop.OctWaterPartCoeff.input.temperature
                 Call TEMPCNV(EnglishValue, phprop.OctWaterPartCoeff.input.temperature)
                 phprop.OctWaterPartCoeff.input.temperature = EnglishValue

                 templdhltmp = phprop.LiquidDiffusivity.haydukLaudie.temperature
                 Call TEMPCNV(EnglishValue, phprop.LiquidDiffusivity.haydukLaudie.temperature)
                 phprop.LiquidDiffusivity.haydukLaudie.temperature = EnglishValue

                 templdptmp = phprop.LiquidDiffusivity.polson.temperature
                 Call TEMPCNV(EnglishValue, phprop.LiquidDiffusivity.polson.temperature)
                 phprop.LiquidDiffusivity.polson.temperature = EnglishValue
                 
                 templdwctmp = phprop.LiquidDiffusivity.wilkeChang.temperature
                 Call TEMPCNV(EnglishValue, phprop.LiquidDiffusivity.wilkeChang.temperature)
                 phprop.LiquidDiffusivity.wilkeChang.temperature = EnglishValue
                 
                 templdtmpi = phprop.LiquidDiffusivity.input.temperature
                 Call TEMPCNV(EnglishValue, phprop.LiquidDiffusivity.input.temperature)
                 phprop.LiquidDiffusivity.input.temperature = EnglishValue
                 
                 tempgdwltmp = phprop.GasDiffusivity.wilkeLee.temperature
                 Call TEMPCNV(EnglishValue, phprop.GasDiffusivity.wilkeLee.temperature)
                 phprop.GasDiffusivity.wilkeLee.temperature = EnglishValue
                 
                 tempgdtmpi = phprop.GasDiffusivity.input.temperature
                 Call TEMPCNV(EnglishValue, phprop.GasDiffusivity.input.temperature)
                 phprop.GasDiffusivity.input.temperature = EnglishValue

                 'Convert values
                 tempvp = phprop.VaporPressure.CurrentSelection.Value
                 tempvpi = phprop.VaporPressure.input.Value
                 Call VPCONV(EnglishValue, phprop.VaporPressure.CurrentSelection.Value)
                 phprop.VaporPressure.CurrentSelection.Value = EnglishValue
                 phprop.VaporPressure.database.Value = EnglishValue
                 Call VPCONV(EnglishValue, phprop.VaporPressure.input.Value)
                 phprop.VaporPressure.input.Value = EnglishValue
                 
                 tempmw = phprop.MolecularWeight.CurrentSelection.Value
                 tempmwdb = phprop.MolecularWeight.database.Value
                 tempmwun = phprop.MolecularWeight.UNIFAC.Value
                 tempmwi = phprop.MolecularWeight.input.Value
                 Call MWCONV(EnglishValue, phprop.MolecularWeight.CurrentSelection.Value)
                 phprop.MolecularWeight.CurrentSelection.Value = EnglishValue
                 Call MWCONV(EnglishValue, phprop.MolecularWeight.database.Value)
                 phprop.MolecularWeight.database.Value = EnglishValue
                 Call MWCONV(EnglishValue, phprop.MolecularWeight.UNIFAC.Value)
                 phprop.MolecularWeight.UNIFAC.Value = EnglishValue
                 Call MWCONV(EnglishValue, phprop.MolecularWeight.input.Value)
                 phprop.MolecularWeight.input.Value = EnglishValue
              
                 tempbp = phprop.BoilingPoint.CurrentSelection.Value
                 tempbpdb = phprop.BoilingPoint.database.Value
                 tempbpi = phprop.BoilingPoint.input.Value
                 Call NBPCONV(EnglishValue, phprop.BoilingPoint.CurrentSelection.Value)
                 phprop.BoilingPoint.CurrentSelection.Value = EnglishValue
                 Call NBPCONV(EnglishValue, phprop.BoilingPoint.database.Value)
                 phprop.BoilingPoint.database.Value = EnglishValue
                 Call NBPCONV(EnglishValue, phprop.BoilingPoint.input.Value)
                 phprop.BoilingPoint.input.Value = EnglishValue
                 
                 templd = phprop.LiquidDensity.CurrentSelection.Value
                 templddb = phprop.LiquidDensity.database.Value
                 templdun = phprop.LiquidDensity.UNIFAC.Value
                 templdi = phprop.LiquidDensity.input.Value
                 Call LDENSCNV(EnglishValue, phprop.LiquidDensity.CurrentSelection.Value)
                 phprop.LiquidDensity.CurrentSelection.Value = EnglishValue
                 Call LDENSCNV(EnglishValue, phprop.LiquidDensity.database.Value)
                 phprop.LiquidDensity.database.Value = EnglishValue
                 Call LDENSCNV(EnglishValue, phprop.LiquidDensity.UNIFAC.Value)
                 phprop.LiquidDensity.UNIFAC.Value = EnglishValue
                 Call LDENSCNV(EnglishValue, phprop.LiquidDensity.input.Value)
                 phprop.LiquidDensity.input.Value = EnglishValue
              
                 tempmvopt = phprop.MolarVolume.operatingT.CurrentSelection.Value
                 tempmvoptdb = phprop.MolarVolume.operatingT.database.Value
                 tempmvoptun = phprop.MolarVolume.operatingT.UNIFAC.Value
                 tempmvopti = phprop.MolarVolume.operatingT.input.Value
                 Call MVOPTCNV(EnglishValue, phprop.MolarVolume.operatingT.CurrentSelection.Value)
                 phprop.MolarVolume.operatingT.CurrentSelection.Value = EnglishValue
                 Call MVOPTCNV(EnglishValue, phprop.MolarVolume.operatingT.database.Value)
                 phprop.MolarVolume.operatingT.database.Value = EnglishValue
                 Call MVOPTCNV(EnglishValue, phprop.MolarVolume.operatingT.UNIFAC.Value)
                 phprop.MolarVolume.operatingT.UNIFAC.Value = EnglishValue
                 Call MVOPTCNV(EnglishValue, phprop.MolarVolume.operatingT.input.Value)
                 phprop.MolarVolume.operatingT.input.Value = EnglishValue
              
                 tempmv = phprop.MolarVolume.BoilingPoint.CurrentSelection.Value
                 tempmvsc = phprop.MolarVolume.BoilingPoint.UNIFAC.Value
                 tempmvi = phprop.MolarVolume.BoilingPoint.input.Value
                 Call MVNBPCNV(EnglishValue, phprop.MolarVolume.BoilingPoint.CurrentSelection.Value)
                 phprop.MolarVolume.BoilingPoint.CurrentSelection.Value = EnglishValue
                 Call MVNBPCNV(EnglishValue, phprop.MolarVolume.BoilingPoint.UNIFAC.Value)
                 phprop.MolarVolume.BoilingPoint.UNIFAC.Value = EnglishValue
                 Call MVNBPCNV(EnglishValue, phprop.MolarVolume.BoilingPoint.input.Value)
                 phprop.MolarVolume.BoilingPoint.input.Value = EnglishValue
              
                 templdiff = phprop.LiquidDiffusivity.CurrentSelection.Value
                 templhldiff = phprop.LiquidDiffusivity.haydukLaudie.Value
                 templpldiff = phprop.LiquidDiffusivity.polson.Value
                 templwcdiff = phprop.LiquidDiffusivity.wilkeChang.Value
                 templdiffi = phprop.LiquidDiffusivity.input.Value
                 Call LDIFFCNV(EnglishValue, phprop.LiquidDiffusivity.CurrentSelection.Value)
                 phprop.LiquidDiffusivity.CurrentSelection.Value = EnglishValue
                 Call LDIFFCNV(EnglishValue, phprop.LiquidDiffusivity.haydukLaudie.Value)
                 phprop.LiquidDiffusivity.haydukLaudie.Value = EnglishValue
                 Call LDIFFCNV(EnglishValue, phprop.LiquidDiffusivity.polson.Value)
                 phprop.LiquidDiffusivity.polson.Value = EnglishValue
                 Call LDIFFCNV(EnglishValue, phprop.LiquidDiffusivity.wilkeChang.Value)
                 phprop.LiquidDiffusivity.wilkeChang.Value = EnglishValue
                 Call LDIFFCNV(EnglishValue, phprop.LiquidDiffusivity.input.Value)
                 phprop.LiquidDiffusivity.input.Value = EnglishValue
              
                 tempgdiff = phprop.GasDiffusivity.CurrentSelection.Value
                 tempgdiffwl = phprop.GasDiffusivity.wilkeLee.Value
                 tempgdiffi = phprop.GasDiffusivity.input.Value
                 Call GDIFFCNV(EnglishValue, phprop.GasDiffusivity.CurrentSelection.Value)
                 phprop.GasDiffusivity.CurrentSelection.Value = EnglishValue
                 Call GDIFFCNV(EnglishValue, phprop.GasDiffusivity.wilkeLee.Value)
                 phprop.GasDiffusivity.wilkeLee.Value = EnglishValue
                 Call GDIFFCNV(EnglishValue, phprop.GasDiffusivity.input.Value)
                 phprop.GasDiffusivity.input.Value = EnglishValue
              
              End If

              For J = 1 To NUMBER_OF_PROPERTIES
                  HaveProperty(J) = phprop.HaveProperty(J)
              Next J
              For J = 1 To NUMBER_OF_PROPERTIES_AVAILABLE
                  PROPAVAILABLE(J) = phprop.PROPAVAILABLE(J)
              Next J

              Call PrintOneContaminantToFile  'Prints the contaminant currently stored in structure phprop

              'If English units were selected convert them back
              If cboUnits.ListIndex = 1 Then
                 
                 'Convert temperatures back
                 phprop.VaporPressure.database.temperature = tempvptmp
                 
                 phprop.VaporPressure.input.temperature = tempvptmpi
                 
                 phprop.ActivityCoefficient.UNIFAC.temperature = tempactmp
                 
                 phprop.HenrysConstant.regress.temperature = temphregtmp
                 
                 phprop.HenrysConstant.fit.UNIFAC.temperature = temphfittmp
                 
                 phprop.HenrysConstant.operatingT.UNIFAC.temperature = temphopttmp
       
                 For J = 1 To phprop.HenrysConstant.NumberOfDatabaseHenrysConstants
                    phprop.HenrysConstant.database(J).temperature = temphdbttmp(J)
                    phprop.HenrysConstant.UNIFAC(J).temperature = temphunttmp(J)
                 Next J
                 
                 phprop.HenrysConstant.input.temperature = temphtmpi
                 
                 phprop.LiquidDensity.database.temperature = templdtmp
                 
                 phprop.LiquidDensity.UNIFAC.temperature = templdutmp
                 
                 phprop.LiquidDensity.input.temperature = templdtmpi
                 
                 phprop.MolarVolume.operatingT.database.temperature = tempmvopttmp
                 
                 phprop.MolarVolume.operatingT.UNIFAC.temperature = tempmvoptutmp
                 
                 phprop.MolarVolume.operatingT.input.temperature = tempmvopttmpi
                 
                 phprop.MolarVolume.BoilingPoint.UNIFAC.temperature = tempmvtmp
                 
                 phprop.MolarVolume.BoilingPoint.input.temperature = tempmvtmpi
                 
                 phprop.AqueousSolubility.fit.UNIFAC.temperature = tempaqfittmp
                 
                 phprop.AqueousSolubility.operatingT.UNIFAC.temperature = tempaqopttmp
       
                 phprop.AqueousSolubility.database.temperature = tempaqdbtmp
                 
                 phprop.AqueousSolubility.UNIFAC.temperature = tempaquntmp
                 
                 phprop.AqueousSolubility.input.temperature = tempaqtmpi
                 
                 phprop.OctWaterPartCoeff.operatingT.UNIFAC.temperature = tempoctopttmp
              
                 phprop.OctWaterPartCoeff.database.temperature = tempoctdbtmp

                 phprop.OctWaterPartCoeff.databaseT.UNIFAC.temperature = tempoctuntmp

                 phprop.OctWaterPartCoeff.input.temperature = tempocttmpi

                 phprop.LiquidDiffusivity.haydukLaudie.temperature = templdhltmp

                 phprop.LiquidDiffusivity.polson.temperature = templdptmp
                 
                 phprop.LiquidDiffusivity.wilkeChang.temperature = templdwctmp
                 
                 phprop.LiquidDiffusivity.input.temperature = templdtmpi
                 
                 phprop.GasDiffusivity.wilkeLee.temperature = tempgdwltmp
                 
                 phprop.GasDiffusivity.input.temperature = tempgdtmpi
              
                 'Covert values back
                 phprop.OperatingPressure = tempress
              
                 phprop.OperatingTemperature = tempt
              
                 phprop.VaporPressure.CurrentSelection.Value = tempvp
                 phprop.VaporPressure.database.Value = tempvp
                 phprop.VaporPressure.input.Value = tempvpi
              
                 phprop.MolecularWeight.CurrentSelection.Value = tempmw
                 phprop.MolecularWeight.database.Value = tempmwdb
                 phprop.MolecularWeight.UNIFAC.Value = tempmwun
                 phprop.MolecularWeight.input.Value = tempmwi
                 
                 phprop.BoilingPoint.CurrentSelection.Value = tempbp
                 phprop.BoilingPoint.database.Value = tempbpdb
                 phprop.BoilingPoint.input.Value = tempbpi
                 
                 phprop.LiquidDensity.CurrentSelection.Value = templd
                 phprop.LiquidDensity.database.Value = templddb
                 phprop.LiquidDensity.UNIFAC.Value = templdun
                 phprop.LiquidDensity.input.Value = templdi
                 
                 phprop.MolarVolume.operatingT.CurrentSelection.Value = tempmvopt
                 phprop.MolarVolume.operatingT.database.Value = tempmvoptdb
                 phprop.MolarVolume.operatingT.UNIFAC.Value = tempmvoptun
                 phprop.MolarVolume.operatingT.input.Value = tempmvopti

                 phprop.MolarVolume.BoilingPoint.CurrentSelection.Value = tempmv
                 phprop.MolarVolume.BoilingPoint.UNIFAC.Value = tempmvsc
                 phprop.MolarVolume.BoilingPoint.input.Value = tempmvi
                 
                 phprop.LiquidDiffusivity.CurrentSelection.Value = templdiff
                 phprop.LiquidDiffusivity.haydukLaudie.Value = templhldiff
                 phprop.LiquidDiffusivity.polson.Value = templpldiff
                 phprop.LiquidDiffusivity.wilkeChang.Value = templwcdiff
                 phprop.LiquidDiffusivity.input.Value = templdiffi
              
                 phprop.GasDiffusivity.CurrentSelection.Value = tempgdiff
                 phprop.GasDiffusivity.wilkeLee.Value = tempgdiffwl
                 phprop.GasDiffusivity.input.Value = tempgdiffi
              
              End If
              
              If i = NumSelectedChemicals Then Exit For
           
              Print #1,
              Print #1,
              Print #1,
              Print #1,
           
          Next i

          If ChosenAtLeastOneAirWaterProperty Then
             Print #1,
             Print #1,
             Print #1,
             Print #1,
             Print #1,
             Print #1,
          End If

          phprop = PropContaminant(CurrentlySelectedContaminant)

          For i = 1 To NUMBER_OF_PROPERTIES
              HaveProperty(i) = phprop.HaveProperty(i)
          Next i
          For i = 1 To NUMBER_OF_PROPERTIES_AVAILABLE
              PROPAVAILABLE(i) = phprop.PROPAVAILABLE(i)
          Next i
       End If

       If ChosenAtLeastOneAirWaterProperty Then
          
          'If English units are desired convert them here
          If cboUnits.ListIndex = 1 Then
              
             temppress = phprop.OperatingPressure
             Call PRESSCNV(EnglishValue, phprop.OperatingPressure)
             phprop.OperatingPressure = EnglishValue
              
             tempt = phprop.OperatingTemperature
             Call TEMPCNV(EnglishValue, phprop.OperatingTemperature)
             phprop.OperatingTemperature = EnglishValue
             
             'Convert Temperatures
             tempwdtmp = phprop.WaterDensity.correlation.temperature
             tempwdtmpi = phprop.WaterDensity.input.temperature
             Call TEMPCNV(EnglishValue, phprop.WaterDensity.correlation.temperature)
             phprop.WaterDensity.correlation.temperature = EnglishValue
             Call TEMPCNV(EnglishValue, phprop.WaterDensity.input.temperature)
             phprop.WaterDensity.input.temperature = EnglishValue
             
             tempwvtmp = phprop.WaterViscosity.correlation.temperature
             tempwvtmpi = phprop.WaterViscosity.input.temperature
             Call TEMPCNV(EnglishValue, phprop.WaterViscosity.correlation.temperature)
             phprop.WaterViscosity.correlation.temperature = EnglishValue
             Call TEMPCNV(EnglishValue, phprop.WaterViscosity.input.temperature)
             phprop.WaterViscosity.input.temperature = EnglishValue
             
             tempwsttmp = phprop.WaterSurfaceTension.correlation.temperature
             tempwsttmpi = phprop.WaterSurfaceTension.input.temperature
             Call TEMPCNV(EnglishValue, phprop.WaterSurfaceTension.correlation.temperature)
             phprop.WaterSurfaceTension.correlation.temperature = EnglishValue
             Call TEMPCNV(EnglishValue, phprop.WaterSurfaceTension.input.temperature)
             phprop.WaterSurfaceTension.input.temperature = EnglishValue
             
             tempadtmp = phprop.AirDensity.correlation.temperature
             tempadtmpi = phprop.AirDensity.input.temperature
             Call TEMPCNV(EnglishValue, phprop.AirDensity.correlation.temperature)
             phprop.AirDensity.correlation.temperature = EnglishValue
             Call TEMPCNV(EnglishValue, phprop.AirDensity.input.temperature)
             phprop.AirDensity.input.temperature = EnglishValue
             
             tempavtmp = phprop.AirViscosity.correlation.temperature
             tempavtmpi = phprop.AirViscosity.input.temperature
             Call TEMPCNV(EnglishValue, phprop.AirViscosity.correlation.temperature)
             phprop.AirViscosity.correlation.temperature = EnglishValue
             Call TEMPCNV(EnglishValue, phprop.AirViscosity.input.temperature)
             phprop.AirViscosity.input.temperature = EnglishValue
              
             'Convert Values
             tempwd = phprop.WaterDensity.CurrentSelection.Value
             tempwdcor = phprop.WaterDensity.correlation.Value
             tempwdi = phprop.WaterDensity.input.Value
             Call WDENSCNV(EnglishValue, phprop.WaterDensity.CurrentSelection.Value)
             phprop.WaterDensity.CurrentSelection.Value = EnglishValue
             Call WDENSCNV(EnglishValue, phprop.WaterDensity.correlation.Value)
             phprop.WaterDensity.correlation.Value = EnglishValue
             Call WDENSCNV(EnglishValue, phprop.WaterDensity.input.Value)
             phprop.WaterDensity.input.Value = EnglishValue
             
             tempwv = phprop.WaterViscosity.CurrentSelection.Value
             tempwvcor = phprop.WaterViscosity.correlation.Value
             tempwvi = phprop.WaterViscosity.input.Value
             Call WVISCCNV(EnglishValue, phprop.WaterViscosity.CurrentSelection.Value)
             phprop.WaterViscosity.CurrentSelection.Value = EnglishValue
             Call WVISCCNV(EnglishValue, phprop.WaterViscosity.correlation.Value)
             phprop.WaterViscosity.correlation.Value = EnglishValue
             Call WVISCCNV(EnglishValue, phprop.WaterViscosity.input.Value)
             phprop.WaterViscosity.input.Value = EnglishValue
             
             tempwst = phprop.WaterSurfaceTension.CurrentSelection.Value
             tempwstcor = phprop.WaterSurfaceTension.correlation.Value
             tempwsti = phprop.WaterSurfaceTension.input.Value
             Call H2OSTCNV(EnglishValue, phprop.WaterSurfaceTension.CurrentSelection.Value)
             phprop.WaterSurfaceTension.CurrentSelection.Value = EnglishValue
             Call H2OSTCNV(EnglishValue, phprop.WaterSurfaceTension.correlation.Value)
             phprop.WaterSurfaceTension.correlation.Value = EnglishValue
             Call H2OSTCNV(EnglishValue, phprop.WaterSurfaceTension.input.Value)
             phprop.WaterSurfaceTension.input.Value = EnglishValue
             
             tempad = phprop.AirDensity.CurrentSelection.Value
             tempadcor = phprop.AirDensity.correlation.Value
             tempadi = phprop.AirDensity.input.Value
             Call ADENSCNV(EnglishValue, phprop.AirDensity.CurrentSelection.Value)
             phprop.AirDensity.CurrentSelection.Value = EnglishValue
             Call ADENSCNV(EnglishValue, phprop.AirDensity.correlation.Value)
             phprop.AirDensity.correlation.Value = EnglishValue
             Call ADENSCNV(EnglishValue, phprop.AirDensity.input.Value)
             phprop.AirDensity.input.Value = EnglishValue
             
             tempav = phprop.AirViscosity.CurrentSelection.Value
             tempavcor = phprop.AirViscosity.correlation.Value
             tempavi = phprop.AirViscosity.input.Value
             Call AVISCCNV(EnglishValue, phprop.AirViscosity.CurrentSelection.Value)
             phprop.AirViscosity.CurrentSelection.Value = EnglishValue
             Call AVISCCNV(EnglishValue, phprop.AirViscosity.correlation.Value)
             phprop.AirViscosity.correlation.Value = EnglishValue
             Call AVISCCNV(EnglishValue, phprop.AirViscosity.input.Value)
             phprop.AirViscosity.input.Value = EnglishValue

          End If

          Call PrintAirWaterPropertiesToFile
          
          'If English units were selected convert them back
          If cboUnits.ListIndex = 1 Then
              
             phprop.OperatingPressure = temppress
              
             phprop.OperatingTemperature = tempt
             
             'Convert temperatures back
             
             phprop.WaterDensity.correlation.temperature = tempwdtmp
             phprop.WaterDensity.input.temperature = tempwdtmpi
             
             phprop.WaterViscosity.correlation.temperature = tempwvtmp
             phprop.WaterViscosity.input.temperature = tempwvtmpi
             
             phprop.WaterSurfaceTension.correlation.temperature = tempwsttmp
             phprop.WaterSurfaceTension.input.temperature = tempwsttmpi
             
             phprop.AirDensity.correlation.temperature = tempadtmp
             phprop.AirDensity.input.temperature = tempadtmpi
             
             phprop.AirViscosity.correlation.temperature = tempavtmp
             phprop.AirViscosity.input.temperature = tempavtmpi
              
             'Convert values back
             phprop.WaterDensity.CurrentSelection.Value = tempwd
             phprop.WaterDensity.correlation.Value = tempwdcor
             phprop.WaterDensity.input.Value = tempwdi
             
             phprop.WaterViscosity.CurrentSelection.Value = tempwv
             phprop.WaterViscosity.correlation.Value = tempwvcor
             phprop.WaterViscosity.input.Value = tempwvi
             
             phprop.WaterSurfaceTension.CurrentSelection.Value = tempwst
             phprop.WaterSurfaceTension.correlation.Value = tempwstcor
             phprop.WaterSurfaceTension.input.Value = tempwsti
             
             phprop.AirDensity.CurrentSelection.Value = tempad
             phprop.AirDensity.correlation.Value = tempadcor
             phprop.AirDensity.input.Value = tempadi
             
             phprop.AirViscosity.CurrentSelection.Value = tempav
             phprop.AirViscosity.correlation.Value = tempavcor
             phprop.AirViscosity.input.Value = tempavi

          End If
       
       End If

       ElseIf frmPrint!optPrintContaminants(1).Value Then   'Print Currently Selected Contaminant
          If ChosenAtLeastOneContaminantProperty Then
             
              'If English units are desired convert them here
              If cboUnits.ListIndex = 1 Then
              
                 temppress = phprop.OperatingPressure
                 Call PRESSCNV(EnglishValue, phprop.OperatingPressure)
                 phprop.OperatingPressure = EnglishValue
              
                 tempt = phprop.OperatingTemperature
                 Call TEMPCNV(EnglishValue, phprop.OperatingTemperature)
                 phprop.OperatingTemperature = EnglishValue
                 
                 'Change all temperatures
                 tempvptmp = phprop.VaporPressure.database.temperature
                 Call TEMPCNV(EnglishValue, phprop.VaporPressure.database.temperature)
                 phprop.VaporPressure.database.temperature = EnglishValue
                 
                 tempvptmpi = phprop.VaporPressure.input.temperature
                 Call TEMPCNV(EnglishValue, phprop.VaporPressure.input.temperature)
                 phprop.VaporPressure.input.temperature = EnglishValue
                 
                 tempactmp = phprop.ActivityCoefficient.UNIFAC.temperature
                 Call TEMPCNV(EnglishValue, phprop.ActivityCoefficient.UNIFAC.temperature)
                 phprop.ActivityCoefficient.UNIFAC.temperature = EnglishValue
                 
                 temphregtmp = phprop.HenrysConstant.regress.temperature
                 Call TEMPCNV(EnglishValue, phprop.HenrysConstant.regress.temperature)
                 phprop.HenrysConstant.regress.temperature = EnglishValue
                 
                 temphfittmp = phprop.HenrysConstant.fit.UNIFAC.temperature
                 Call TEMPCNV(EnglishValue, phprop.HenrysConstant.fit.UNIFAC.temperature)
                 phprop.HenrysConstant.fit.UNIFAC.temperature = EnglishValue
                 
                 temphopttmp = phprop.HenrysConstant.operatingT.UNIFAC.temperature
                 Call TEMPCNV(EnglishValue, phprop.HenrysConstant.operatingT.UNIFAC.temperature)
                 phprop.HenrysConstant.operatingT.UNIFAC.temperature = EnglishValue
       
                 For J = 1 To phprop.HenrysConstant.NumberOfDatabaseHenrysConstants
                    temphdbttmp(J) = phprop.HenrysConstant.database(J).temperature
                    temphunttmp(J) = phprop.HenrysConstant.UNIFAC(J).temperature
                    Call TEMPCNV(EnglishValue, phprop.HenrysConstant.database(J).temperature)
                    phprop.HenrysConstant.database(J).temperature = EnglishValue
                    Call TEMPCNV(EnglishValue, phprop.HenrysConstant.UNIFAC(J).temperature)
                    phprop.HenrysConstant.UNIFAC(J).temperature = EnglishValue
                 Next J
                 
                 temphtmpi = phprop.HenrysConstant.input.temperature
                 Call TEMPCNV(EnglishValue, phprop.HenrysConstant.input.temperature)
                 phprop.HenrysConstant.input.temperature = EnglishValue
                 
                 templdtmp = phprop.LiquidDensity.database.temperature
                 Call TEMPCNV(EnglishValue, phprop.LiquidDensity.database.temperature)
                 phprop.LiquidDensity.database.temperature = EnglishValue
                 
                 templdutmp = phprop.LiquidDensity.UNIFAC.temperature
                 Call TEMPCNV(EnglishValue, phprop.LiquidDensity.UNIFAC.temperature)
                 phprop.LiquidDensity.UNIFAC.temperature = EnglishValue
                 
                 templdtmpi = phprop.LiquidDensity.input.temperature
                 Call TEMPCNV(EnglishValue, phprop.LiquidDensity.input.temperature)
                 phprop.LiquidDensity.input.temperature = EnglishValue
                 
                 tempmvopttmp = phprop.MolarVolume.operatingT.database.temperature
                 Call TEMPCNV(EnglishValue, phprop.MolarVolume.operatingT.database.temperature)
                 phprop.MolarVolume.operatingT.database.temperature = EnglishValue
                 
                 tempmvoptutmp = phprop.MolarVolume.operatingT.UNIFAC.temperature
                 Call TEMPCNV(EnglishValue, phprop.MolarVolume.operatingT.UNIFAC.temperature)
                 phprop.MolarVolume.operatingT.UNIFAC.temperature = EnglishValue
                 
                 tempmvopttmpi = phprop.MolarVolume.operatingT.input.temperature
                 Call TEMPCNV(EnglishValue, phprop.MolarVolume.operatingT.input.temperature)
                 phprop.MolarVolume.operatingT.input.temperature = EnglishValue
                 
                 tempmvtmp = phprop.MolarVolume.BoilingPoint.UNIFAC.temperature
                 Call TEMPCNV(EnglishValue, phprop.MolarVolume.BoilingPoint.UNIFAC.temperature)
                 phprop.MolarVolume.BoilingPoint.UNIFAC.temperature = EnglishValue
                 
                 tempmvtmpi = phprop.MolarVolume.BoilingPoint.input.temperature
                 Call TEMPCNV(EnglishValue, phprop.MolarVolume.BoilingPoint.input.temperature)
                 phprop.MolarVolume.BoilingPoint.input.temperature = EnglishValue
                 
                 tempaqfittmp = phprop.AqueousSolubility.fit.UNIFAC.temperature
                 Call TEMPCNV(EnglishValue, phprop.AqueousSolubility.fit.UNIFAC.temperature)
                 phprop.AqueousSolubility.fit.UNIFAC.temperature = EnglishValue
                 
                 tempaqopttmp = phprop.AqueousSolubility.operatingT.UNIFAC.temperature
                 Call TEMPCNV(EnglishValue, phprop.AqueousSolubility.operatingT.UNIFAC.temperature)
                 phprop.AqueousSolubility.operatingT.UNIFAC.temperature = EnglishValue
       
                 tempaqdbtmp = phprop.AqueousSolubility.database.temperature
                 Call TEMPCNV(EnglishValue, phprop.AqueousSolubility.database.temperature)
                 phprop.AqueousSolubility.database.temperature = EnglishValue
                 
                 tempaquntmp = phprop.AqueousSolubility.UNIFAC.temperature
                 Call TEMPCNV(EnglishValue, phprop.AqueousSolubility.UNIFAC.temperature)
                 phprop.AqueousSolubility.UNIFAC.temperature = EnglishValue
                 
                 tempaqtmpi = phprop.AqueousSolubility.input.temperature
                 Call TEMPCNV(EnglishValue, phprop.AqueousSolubility.input.temperature)
                 phprop.AqueousSolubility.input.temperature = EnglishValue
                 
                 tempoctopttmp = phprop.OctWaterPartCoeff.operatingT.UNIFAC.temperature
                 Call TEMPCNV(EnglishValue, phprop.OctWaterPartCoeff.operatingT.UNIFAC.temperature)
                 phprop.OctWaterPartCoeff.operatingT.UNIFAC.temperature = EnglishValue
              
                 tempoctdbtmp = phprop.OctWaterPartCoeff.database.temperature
                 Call TEMPCNV(EnglishValue, phprop.OctWaterPartCoeff.database.temperature)
                 phprop.OctWaterPartCoeff.database.temperature = EnglishValue

                 tempoctuntmp = phprop.OctWaterPartCoeff.databaseT.UNIFAC.temperature
                 Call TEMPCNV(EnglishValue, phprop.OctWaterPartCoeff.databaseT.UNIFAC.temperature)
                 phprop.OctWaterPartCoeff.databaseT.UNIFAC.temperature = EnglishValue

                 tempocttmpi = phprop.OctWaterPartCoeff.input.temperature
                 Call TEMPCNV(EnglishValue, phprop.OctWaterPartCoeff.input.temperature)
                 phprop.OctWaterPartCoeff.input.temperature = EnglishValue

                 templdhltmp = phprop.LiquidDiffusivity.haydukLaudie.temperature
                 Call TEMPCNV(EnglishValue, phprop.LiquidDiffusivity.haydukLaudie.temperature)
                 phprop.LiquidDiffusivity.haydukLaudie.temperature = EnglishValue

                 templdptmp = phprop.LiquidDiffusivity.polson.temperature
                 Call TEMPCNV(EnglishValue, phprop.LiquidDiffusivity.polson.temperature)
                 phprop.LiquidDiffusivity.polson.temperature = EnglishValue
                 
                 templdwctmp = phprop.LiquidDiffusivity.wilkeChang.temperature
                 Call TEMPCNV(EnglishValue, phprop.LiquidDiffusivity.wilkeChang.temperature)
                 phprop.LiquidDiffusivity.wilkeChang.temperature = EnglishValue
                 
                 templdtmpi = phprop.LiquidDiffusivity.input.temperature
                 Call TEMPCNV(EnglishValue, phprop.LiquidDiffusivity.input.temperature)
                 phprop.LiquidDiffusivity.input.temperature = EnglishValue
                 
                 tempgdwltmp = phprop.GasDiffusivity.wilkeLee.temperature
                 Call TEMPCNV(EnglishValue, phprop.GasDiffusivity.wilkeLee.temperature)
                 phprop.GasDiffusivity.wilkeLee.temperature = EnglishValue
                 
                 tempgdtmpi = phprop.GasDiffusivity.input.temperature
                 Call TEMPCNV(EnglishValue, phprop.GasDiffusivity.input.temperature)
                 phprop.GasDiffusivity.input.temperature = EnglishValue

                 'Convert values
                 tempvp = phprop.VaporPressure.CurrentSelection.Value
                 tempvpi = phprop.VaporPressure.input.Value
                 Call VPCONV(EnglishValue, phprop.VaporPressure.CurrentSelection.Value)
                 phprop.VaporPressure.CurrentSelection.Value = EnglishValue
                 phprop.VaporPressure.database.Value = EnglishValue
                 Call VPCONV(EnglishValue, phprop.VaporPressure.input.Value)
                 phprop.VaporPressure.input.Value = EnglishValue
                 
                 tempmw = phprop.MolecularWeight.CurrentSelection.Value
                 tempmwdb = phprop.MolecularWeight.database.Value
                 tempmwun = phprop.MolecularWeight.UNIFAC.Value
                 tempmwi = phprop.MolecularWeight.input.Value
                 Call MWCONV(EnglishValue, phprop.MolecularWeight.CurrentSelection.Value)
                 phprop.MolecularWeight.CurrentSelection.Value = EnglishValue
                 Call MWCONV(EnglishValue, phprop.MolecularWeight.database.Value)
                 phprop.MolecularWeight.database.Value = EnglishValue
                 Call MWCONV(EnglishValue, phprop.MolecularWeight.UNIFAC.Value)
                 phprop.MolecularWeight.UNIFAC.Value = EnglishValue
                 Call MWCONV(EnglishValue, phprop.MolecularWeight.input.Value)
                 phprop.MolecularWeight.input.Value = EnglishValue
              
                 tempbp = phprop.BoilingPoint.CurrentSelection.Value
                 tempbpdb = phprop.BoilingPoint.database.Value
                 tempbpi = phprop.BoilingPoint.input.Value
                 Call NBPCONV(EnglishValue, phprop.BoilingPoint.CurrentSelection.Value)
                 phprop.BoilingPoint.CurrentSelection.Value = EnglishValue
                 Call NBPCONV(EnglishValue, phprop.BoilingPoint.database.Value)
                 phprop.BoilingPoint.database.Value = EnglishValue
                 Call NBPCONV(EnglishValue, phprop.BoilingPoint.input.Value)
                 phprop.BoilingPoint.input.Value = EnglishValue
                 
                 templd = phprop.LiquidDensity.CurrentSelection.Value
                 templddb = phprop.LiquidDensity.database.Value
                 templdun = phprop.LiquidDensity.UNIFAC.Value
                 templdi = phprop.LiquidDensity.input.Value
                 Call LDENSCNV(EnglishValue, phprop.LiquidDensity.CurrentSelection.Value)
                 phprop.LiquidDensity.CurrentSelection.Value = EnglishValue
                 Call LDENSCNV(EnglishValue, phprop.LiquidDensity.database.Value)
                 phprop.LiquidDensity.database.Value = EnglishValue
                 Call LDENSCNV(EnglishValue, phprop.LiquidDensity.UNIFAC.Value)
                 phprop.LiquidDensity.UNIFAC.Value = EnglishValue
                 Call LDENSCNV(EnglishValue, phprop.LiquidDensity.input.Value)
                 phprop.LiquidDensity.input.Value = EnglishValue
              
                 tempmvopt = phprop.MolarVolume.operatingT.CurrentSelection.Value
                 tempmvoptdb = phprop.MolarVolume.operatingT.database.Value
                 tempmvoptun = phprop.MolarVolume.operatingT.UNIFAC.Value
                 tempmvopti = phprop.MolarVolume.operatingT.input.Value
                 Call MVOPTCNV(EnglishValue, phprop.MolarVolume.operatingT.CurrentSelection.Value)
                 phprop.MolarVolume.operatingT.CurrentSelection.Value = EnglishValue
                 Call MVOPTCNV(EnglishValue, phprop.MolarVolume.operatingT.database.Value)
                 phprop.MolarVolume.operatingT.database.Value = EnglishValue
                 Call MVOPTCNV(EnglishValue, phprop.MolarVolume.operatingT.UNIFAC.Value)
                 phprop.MolarVolume.operatingT.UNIFAC.Value = EnglishValue
                 Call MVOPTCNV(EnglishValue, phprop.MolarVolume.operatingT.input.Value)
                 phprop.MolarVolume.operatingT.input.Value = EnglishValue
              
                 tempmv = phprop.MolarVolume.BoilingPoint.CurrentSelection.Value
                 tempmvsc = phprop.MolarVolume.BoilingPoint.UNIFAC.Value
                 tempmvi = phprop.MolarVolume.BoilingPoint.input.Value
                 Call MVNBPCNV(EnglishValue, phprop.MolarVolume.BoilingPoint.CurrentSelection.Value)
                 phprop.MolarVolume.BoilingPoint.CurrentSelection.Value = EnglishValue
                 Call MVNBPCNV(EnglishValue, phprop.MolarVolume.BoilingPoint.UNIFAC.Value)
                 phprop.MolarVolume.BoilingPoint.UNIFAC.Value = EnglishValue
                 Call MVNBPCNV(EnglishValue, phprop.MolarVolume.BoilingPoint.input.Value)
                 phprop.MolarVolume.BoilingPoint.input.Value = EnglishValue
              
                 templdiff = phprop.LiquidDiffusivity.CurrentSelection.Value
                 templhldiff = phprop.LiquidDiffusivity.haydukLaudie.Value
                 templpldiff = phprop.LiquidDiffusivity.polson.Value
                 templwcdiff = phprop.LiquidDiffusivity.wilkeChang.Value
                 templdiffi = phprop.LiquidDiffusivity.input.Value
                 Call LDIFFCNV(EnglishValue, phprop.LiquidDiffusivity.CurrentSelection.Value)
                 phprop.LiquidDiffusivity.CurrentSelection.Value = EnglishValue
                 Call LDIFFCNV(EnglishValue, phprop.LiquidDiffusivity.haydukLaudie.Value)
                 phprop.LiquidDiffusivity.haydukLaudie.Value = EnglishValue
                 Call LDIFFCNV(EnglishValue, phprop.LiquidDiffusivity.polson.Value)
                 phprop.LiquidDiffusivity.polson.Value = EnglishValue
                 Call LDIFFCNV(EnglishValue, phprop.LiquidDiffusivity.wilkeChang.Value)
                 phprop.LiquidDiffusivity.wilkeChang.Value = EnglishValue
                 Call LDIFFCNV(EnglishValue, phprop.LiquidDiffusivity.input.Value)
                 phprop.LiquidDiffusivity.input.Value = EnglishValue
              
                 tempgdiff = phprop.GasDiffusivity.CurrentSelection.Value
                 tempgdiffwl = phprop.GasDiffusivity.wilkeLee.Value
                 tempgdiffi = phprop.GasDiffusivity.input.Value
                 Call GDIFFCNV(EnglishValue, phprop.GasDiffusivity.CurrentSelection.Value)
                 phprop.GasDiffusivity.CurrentSelection.Value = EnglishValue
                 Call GDIFFCNV(EnglishValue, phprop.GasDiffusivity.wilkeLee.Value)
                 phprop.GasDiffusivity.wilkeLee.Value = EnglishValue
                 Call GDIFFCNV(EnglishValue, phprop.GasDiffusivity.input.Value)
                 phprop.GasDiffusivity.input.Value = EnglishValue
              
              End If
             
             Call PrintOneContaminantToFile
              
              'If English units were selected convert them back
              If cboUnits.ListIndex = 1 Then
                 
                 'Convert temperatures back
                 phprop.VaporPressure.database.temperature = tempvptmp
                 
                 phprop.VaporPressure.input.temperature = tempvptmpi
                 
                 phprop.ActivityCoefficient.UNIFAC.temperature = tempactmp
                 
                 phprop.HenrysConstant.regress.temperature = temphregtmp
                 
                 phprop.HenrysConstant.fit.UNIFAC.temperature = temphfittmp
                 
                 phprop.HenrysConstant.operatingT.UNIFAC.temperature = temphopttmp
       
                 For J = 1 To phprop.HenrysConstant.NumberOfDatabaseHenrysConstants
                    phprop.HenrysConstant.database(J).temperature = temphdbttmp(J)
                    phprop.HenrysConstant.UNIFAC(J).temperature = temphunttmp(J)
                 Next J
                 
                 phprop.HenrysConstant.input.temperature = temphtmpi
                 
                 phprop.LiquidDensity.database.temperature = templdtmp
                 
                 phprop.LiquidDensity.UNIFAC.temperature = templdutmp
                 
                 phprop.LiquidDensity.input.temperature = templdtmpi
                 
                 phprop.MolarVolume.operatingT.database.temperature = tempmvopttmp
                 
                 phprop.MolarVolume.operatingT.UNIFAC.temperature = tempmvoptutmp
                 
                 phprop.MolarVolume.operatingT.input.temperature = tempmvopttmpi
                 
                 phprop.MolarVolume.BoilingPoint.UNIFAC.temperature = tempmvtmp
                 
                 phprop.MolarVolume.BoilingPoint.input.temperature = tempmvtmpi
                 
                 phprop.AqueousSolubility.fit.UNIFAC.temperature = tempaqfittmp
                 
                 phprop.AqueousSolubility.operatingT.UNIFAC.temperature = tempaqopttmp
       
                 phprop.AqueousSolubility.database.temperature = tempaqdbtmp
                 
                 phprop.AqueousSolubility.UNIFAC.temperature = tempaquntmp
                 
                 phprop.AqueousSolubility.input.temperature = tempaqtmpi
                 
                 phprop.OctWaterPartCoeff.operatingT.UNIFAC.temperature = tempoctopttmp
              
                 phprop.OctWaterPartCoeff.database.temperature = tempoctdbtmp

                 phprop.OctWaterPartCoeff.databaseT.UNIFAC.temperature = tempoctuntmp

                 phprop.OctWaterPartCoeff.input.temperature = tempocttmpi

                 phprop.LiquidDiffusivity.haydukLaudie.temperature = templdhltmp

                 phprop.LiquidDiffusivity.polson.temperature = templdptmp
                 
                 phprop.LiquidDiffusivity.wilkeChang.temperature = templdwctmp
                 
                 phprop.LiquidDiffusivity.input.temperature = templdtmpi
                 
                 phprop.GasDiffusivity.wilkeLee.temperature = tempgdwltmp
                 
                 phprop.GasDiffusivity.input.temperature = tempgdtmpi
              
                 'Covert values back
                 phprop.OperatingPressure = tempress
              
                 phprop.OperatingTemperature = tempt
              
                 phprop.VaporPressure.CurrentSelection.Value = tempvp
                 phprop.VaporPressure.database.Value = tempvp
                 phprop.VaporPressure.input.Value = tempvpi
              
                 phprop.MolecularWeight.CurrentSelection.Value = tempmw
                 phprop.MolecularWeight.database.Value = tempmwdb
                 phprop.MolecularWeight.UNIFAC.Value = tempmwun
                 phprop.MolecularWeight.input.Value = tempmwi
                 
                 phprop.BoilingPoint.CurrentSelection.Value = tempbp
                 phprop.BoilingPoint.database.Value = tempbpdb
                 phprop.BoilingPoint.input.Value = tempbpi
                 
                 phprop.LiquidDensity.CurrentSelection.Value = templd
                 phprop.LiquidDensity.database.Value = templddb
                 phprop.LiquidDensity.UNIFAC.Value = templdun
                 phprop.LiquidDensity.input.Value = templdi
                 
                 phprop.MolarVolume.operatingT.CurrentSelection.Value = tempmvopt
                 phprop.MolarVolume.operatingT.database.Value = tempmvoptdb
                 phprop.MolarVolume.operatingT.UNIFAC.Value = tempmvoptun
                 phprop.MolarVolume.operatingT.input.Value = tempmvopti

                 phprop.MolarVolume.BoilingPoint.CurrentSelection.Value = tempmv
                 phprop.MolarVolume.BoilingPoint.UNIFAC.Value = tempmvsc
                 phprop.MolarVolume.BoilingPoint.input.Value = tempmvi
                 
                 phprop.LiquidDiffusivity.CurrentSelection.Value = templdiff
                 phprop.LiquidDiffusivity.haydukLaudie.Value = templhldiff
                 phprop.LiquidDiffusivity.polson.Value = templpldiff
                 phprop.LiquidDiffusivity.wilkeChang.Value = templwcdiff
                 phprop.LiquidDiffusivity.input.Value = templdiffi
              
                 phprop.GasDiffusivity.CurrentSelection.Value = tempgdiff
                 phprop.GasDiffusivity.wilkeLee.Value = tempgdiffwl
                 phprop.GasDiffusivity.input.Value = tempgdiffi
              
              End If
              
          End If

          If ChosenAtLeastOneAirWaterProperty Then
             Print #1,
             Print #1,
             Print #1,
             Print #1,
             Print #1,
             Print #1,
          End If

          If ChosenAtLeastOneAirWaterProperty Then
             
          'If English units are desired convert them here
          If cboUnits.ListIndex = 1 Then
              
             temppress = phprop.OperatingPressure
             Call PRESSCNV(EnglishValue, phprop.OperatingPressure)
             phprop.OperatingPressure = EnglishValue
              
             tempt = phprop.OperatingTemperature
             Call TEMPCNV(EnglishValue, phprop.OperatingTemperature)
             phprop.OperatingTemperature = EnglishValue
             
             'Convert Temperatures
             tempwdtmp = phprop.WaterDensity.correlation.temperature
             tempwdtmpi = phprop.WaterDensity.input.temperature
             Call TEMPCNV(EnglishValue, phprop.WaterDensity.correlation.temperature)
             phprop.WaterDensity.correlation.temperature = EnglishValue
             Call TEMPCNV(EnglishValue, phprop.WaterDensity.input.temperature)
             phprop.WaterDensity.input.temperature = EnglishValue
             
             tempwvtmp = phprop.WaterViscosity.correlation.temperature
             tempwvtmpi = phprop.WaterViscosity.input.temperature
             Call TEMPCNV(EnglishValue, phprop.WaterViscosity.correlation.temperature)
             phprop.WaterViscosity.correlation.temperature = EnglishValue
             Call TEMPCNV(EnglishValue, phprop.WaterViscosity.input.temperature)
             phprop.WaterViscosity.input.temperature = EnglishValue
             
             tempwsttmp = phprop.WaterSurfaceTension.correlation.temperature
             tempwsttmpi = phprop.WaterSurfaceTension.input.temperature
             Call TEMPCNV(EnglishValue, phprop.WaterSurfaceTension.correlation.temperature)
             phprop.WaterSurfaceTension.correlation.temperature = EnglishValue
             Call TEMPCNV(EnglishValue, phprop.WaterSurfaceTension.input.temperature)
             phprop.WaterSurfaceTension.input.temperature = EnglishValue
             
             tempadtmp = phprop.AirDensity.correlation.temperature
             tempadtmpi = phprop.AirDensity.input.temperature
             Call TEMPCNV(EnglishValue, phprop.AirDensity.correlation.temperature)
             phprop.AirDensity.correlation.temperature = EnglishValue
             Call TEMPCNV(EnglishValue, phprop.AirDensity.input.temperature)
             phprop.AirDensity.input.temperature = EnglishValue
             
             tempavtmp = phprop.AirViscosity.correlation.temperature
             tempavtmpi = phprop.AirViscosity.input.temperature
             Call TEMPCNV(EnglishValue, phprop.AirViscosity.correlation.temperature)
             phprop.AirViscosity.correlation.temperature = EnglishValue
             Call TEMPCNV(EnglishValue, phprop.AirViscosity.input.temperature)
             phprop.AirViscosity.input.temperature = EnglishValue
              
             'Convert Values
             tempwd = phprop.WaterDensity.CurrentSelection.Value
             tempwdcor = phprop.WaterDensity.correlation.Value
             tempwdi = phprop.WaterDensity.input.Value
             Call WDENSCNV(EnglishValue, phprop.WaterDensity.CurrentSelection.Value)
             phprop.WaterDensity.CurrentSelection.Value = EnglishValue
             Call WDENSCNV(EnglishValue, phprop.WaterDensity.correlation.Value)
             phprop.WaterDensity.correlation.Value = EnglishValue
             Call WDENSCNV(EnglishValue, phprop.WaterDensity.input.Value)
             phprop.WaterDensity.input.Value = EnglishValue
             
             tempwv = phprop.WaterViscosity.CurrentSelection.Value
             tempwvcor = phprop.WaterViscosity.correlation.Value
             tempwvi = phprop.WaterViscosity.input.Value
             Call WVISCCNV(EnglishValue, phprop.WaterViscosity.CurrentSelection.Value)
             phprop.WaterViscosity.CurrentSelection.Value = EnglishValue
             Call WVISCCNV(EnglishValue, phprop.WaterViscosity.correlation.Value)
             phprop.WaterViscosity.correlation.Value = EnglishValue
             Call WVISCCNV(EnglishValue, phprop.WaterViscosity.input.Value)
             phprop.WaterViscosity.input.Value = EnglishValue
             
             tempwst = phprop.WaterSurfaceTension.CurrentSelection.Value
             tempwstcor = phprop.WaterSurfaceTension.correlation.Value
             tempwsti = phprop.WaterSurfaceTension.input.Value
             Call H2OSTCNV(EnglishValue, phprop.WaterSurfaceTension.CurrentSelection.Value)
             phprop.WaterSurfaceTension.CurrentSelection.Value = EnglishValue
             Call H2OSTCNV(EnglishValue, phprop.WaterSurfaceTension.correlation.Value)
             phprop.WaterSurfaceTension.correlation.Value = EnglishValue
             Call H2OSTCNV(EnglishValue, phprop.WaterSurfaceTension.input.Value)
             phprop.WaterSurfaceTension.input.Value = EnglishValue
             
             tempad = phprop.AirDensity.CurrentSelection.Value
             tempadcor = phprop.AirDensity.correlation.Value
             tempadi = phprop.AirDensity.input.Value
             Call ADENSCNV(EnglishValue, phprop.AirDensity.CurrentSelection.Value)
             phprop.AirDensity.CurrentSelection.Value = EnglishValue
             Call ADENSCNV(EnglishValue, phprop.AirDensity.correlation.Value)
             phprop.AirDensity.correlation.Value = EnglishValue
             Call ADENSCNV(EnglishValue, phprop.AirDensity.input.Value)
             phprop.AirDensity.input.Value = EnglishValue
             
             tempav = phprop.AirViscosity.CurrentSelection.Value
             tempavcor = phprop.AirViscosity.correlation.Value
             tempavi = phprop.AirViscosity.input.Value
             Call AVISCCNV(EnglishValue, phprop.AirViscosity.CurrentSelection.Value)
             phprop.AirViscosity.CurrentSelection.Value = EnglishValue
             Call AVISCCNV(EnglishValue, phprop.AirViscosity.correlation.Value)
             phprop.AirViscosity.correlation.Value = EnglishValue
             Call AVISCCNV(EnglishValue, phprop.AirViscosity.input.Value)
             phprop.AirViscosity.input.Value = EnglishValue

          End If
             
             Call PrintAirWaterPropertiesToFile
          
          'If English units were selected convert them back
          If cboUnits.ListIndex = 1 Then
              
             phprop.OperatingPressure = temppress
              
             phprop.OperatingTemperature = tempt
             
             'Convert temperatures back
             
             phprop.WaterDensity.correlation.temperature = tempwdtmp
             phprop.WaterDensity.input.temperature = tempwdtmpi
             
             phprop.WaterViscosity.correlation.temperature = tempwvtmp
             phprop.WaterViscosity.input.temperature = tempwvtmpi
             
             phprop.WaterSurfaceTension.correlation.temperature = tempwsttmp
             phprop.WaterSurfaceTension.input.temperature = tempwsttmpi
             
             phprop.AirDensity.correlation.temperature = tempadtmp
             phprop.AirDensity.input.temperature = tempadtmpi
             
             phprop.AirViscosity.correlation.temperature = tempavtmp
             phprop.AirViscosity.input.temperature = tempavtmpi
              
             'Convert values back
             phprop.WaterDensity.CurrentSelection.Value = tempwd
             phprop.WaterDensity.correlation.Value = tempwdcor
             phprop.WaterDensity.input.Value = tempwdi
             
             phprop.WaterViscosity.CurrentSelection.Value = tempwv
             phprop.WaterViscosity.correlation.Value = tempwvcor
             phprop.WaterViscosity.input.Value = tempwvi
             
             phprop.WaterSurfaceTension.CurrentSelection.Value = tempwst
             phprop.WaterSurfaceTension.correlation.Value = tempwstcor
             phprop.WaterSurfaceTension.input.Value = tempwsti
             
             phprop.AirDensity.CurrentSelection.Value = tempad
             phprop.AirDensity.correlation.Value = tempadcor
             phprop.AirDensity.input.Value = tempadi
             
             phprop.AirViscosity.CurrentSelection.Value = tempav
             phprop.AirViscosity.correlation.Value = tempavcor
             phprop.AirViscosity.input.Value = tempavi

          End If
          
          End If

    End If

resume_exit36:
Exit Sub

error_printtofile:
MsgBox "Error printing to file", mb_exclamation, "StEPP"
Resume resume_exit36


End Sub

Private Sub PrintToPrinter()
    Dim CurrentlySelectedContaminant As Integer
    Dim i As Integer, J As Integer
    Dim PrintHeightOfEachContaminant As Integer   'The Printing Height of Each Contaminant
    Dim PrintSpaceBetweenContaminants As Integer
    Dim ChosenAtLeastOneContaminantProperty As Integer
    Dim ChosenAtLeastOneAirWaterProperty As Integer
    Dim EnglishValue As Double
    Static temphdbttmp(20) As Double
    Static temphunttmp(20) As Double

On Error GoTo error_printtoprinter

    ChosenAtLeastOneContaminantProperty = False
    ChosenAtLeastOneAirWaterProperty = False
    If optPrintProperties(0).Value Then
       ChosenAtLeastOneContaminantProperty = True
       ChosenAtLeastOneAirWaterProperty = True
    Else
       For i = 0 To 12
           If (frmPrint!chkProperties(i).Value = 1) Then
              ChosenAtLeastOneContaminantProperty = True
              Exit For
           End If
       Next i

       For i = 13 To 17
           If (frmPrint!chkProperties(i).Value = 1) Then
              ChosenAtLeastOneAirWaterProperty = True
              Exit For
           End If
       Next i
    End If


    If frmPrint!optPrintContaminants(0).Value Then   'Print All Contaminants
       If ChosenAtLeastOneContaminantProperty Then
          PrintHeightOfEachContaminant = 0
          CurrentlySelectedContaminant = contam_prop_form!cboSelectContaminant.ListIndex + 1
          For i = 1 To NumSelectedChemicals
              phprop = PropContaminant(i)
     
              'If English units are desired convert them here
              If cboUnits.ListIndex = 1 Then
              
                 temppress = phprop.OperatingPressure
                 Call PRESSCNV(EnglishValue, phprop.OperatingPressure)
                 phprop.OperatingPressure = EnglishValue
              
                 tempt = phprop.OperatingTemperature
                 Call TEMPCNV(EnglishValue, phprop.OperatingTemperature)
                 phprop.OperatingTemperature = EnglishValue
                 
                 'Change all temperatures
                 tempvptmp = phprop.VaporPressure.database.temperature
                 Call TEMPCNV(EnglishValue, phprop.VaporPressure.database.temperature)
                 phprop.VaporPressure.database.temperature = EnglishValue
                 
                 tempvptmpi = phprop.VaporPressure.input.temperature
                 Call TEMPCNV(EnglishValue, phprop.VaporPressure.input.temperature)
                 phprop.VaporPressure.input.temperature = EnglishValue
                 
                 tempactmp = phprop.ActivityCoefficient.UNIFAC.temperature
                 Call TEMPCNV(EnglishValue, phprop.ActivityCoefficient.UNIFAC.temperature)
                 phprop.ActivityCoefficient.UNIFAC.temperature = EnglishValue
                 
                 temphregtmp = phprop.HenrysConstant.regress.temperature
                 Call TEMPCNV(EnglishValue, phprop.HenrysConstant.regress.temperature)
                 phprop.HenrysConstant.regress.temperature = EnglishValue
                 
                 temphfittmp = phprop.HenrysConstant.fit.UNIFAC.temperature
                 Call TEMPCNV(EnglishValue, phprop.HenrysConstant.fit.UNIFAC.temperature)
                 phprop.HenrysConstant.fit.UNIFAC.temperature = EnglishValue
                 
                 temphopttmp = phprop.HenrysConstant.operatingT.UNIFAC.temperature
                 Call TEMPCNV(EnglishValue, phprop.HenrysConstant.operatingT.UNIFAC.temperature)
                 phprop.HenrysConstant.operatingT.UNIFAC.temperature = EnglishValue
       
                 For J = 1 To phprop.HenrysConstant.NumberOfDatabaseHenrysConstants
                    temphdbttmp(J) = phprop.HenrysConstant.database(J).temperature
                    temphunttmp(J) = phprop.HenrysConstant.UNIFAC(J).temperature
                    Call TEMPCNV(EnglishValue, phprop.HenrysConstant.database(J).temperature)
                    phprop.HenrysConstant.database(J).temperature = EnglishValue
                    Call TEMPCNV(EnglishValue, phprop.HenrysConstant.UNIFAC(J).temperature)
                    phprop.HenrysConstant.UNIFAC(J).temperature = EnglishValue
                 Next J
                 
                 temphtmpi = phprop.HenrysConstant.input.temperature
                 Call TEMPCNV(EnglishValue, phprop.HenrysConstant.input.temperature)
                 phprop.HenrysConstant.input.temperature = EnglishValue
                 
                 templdtmp = phprop.LiquidDensity.database.temperature
                 Call TEMPCNV(EnglishValue, phprop.LiquidDensity.database.temperature)
                 phprop.LiquidDensity.database.temperature = EnglishValue
                 
                 templdutmp = phprop.LiquidDensity.UNIFAC.temperature
                 Call TEMPCNV(EnglishValue, phprop.LiquidDensity.UNIFAC.temperature)
                 phprop.LiquidDensity.UNIFAC.temperature = EnglishValue
                 
                 templdtmpi = phprop.LiquidDensity.input.temperature
                 Call TEMPCNV(EnglishValue, phprop.LiquidDensity.input.temperature)
                 phprop.LiquidDensity.input.temperature = EnglishValue
                 
                 tempmvopttmp = phprop.MolarVolume.operatingT.database.temperature
                 Call TEMPCNV(EnglishValue, phprop.MolarVolume.operatingT.database.temperature)
                 phprop.MolarVolume.operatingT.database.temperature = EnglishValue
                 
                 tempmvoptutmp = phprop.MolarVolume.operatingT.UNIFAC.temperature
                 Call TEMPCNV(EnglishValue, phprop.MolarVolume.operatingT.UNIFAC.temperature)
                 phprop.MolarVolume.operatingT.UNIFAC.temperature = EnglishValue
                 
                 tempmvopttmpi = phprop.MolarVolume.operatingT.input.temperature
                 Call TEMPCNV(EnglishValue, phprop.MolarVolume.operatingT.input.temperature)
                 phprop.MolarVolume.operatingT.input.temperature = EnglishValue
                 
                 tempmvtmp = phprop.MolarVolume.BoilingPoint.UNIFAC.temperature
                 Call TEMPCNV(EnglishValue, phprop.MolarVolume.BoilingPoint.UNIFAC.temperature)
                 phprop.MolarVolume.BoilingPoint.UNIFAC.temperature = EnglishValue
                 
                 tempmvtmpi = phprop.MolarVolume.BoilingPoint.input.temperature
                 Call TEMPCNV(EnglishValue, phprop.MolarVolume.BoilingPoint.input.temperature)
                 phprop.MolarVolume.BoilingPoint.input.temperature = EnglishValue
                 
                 tempaqfittmp = phprop.AqueousSolubility.fit.UNIFAC.temperature
                 Call TEMPCNV(EnglishValue, phprop.AqueousSolubility.fit.UNIFAC.temperature)
                 phprop.AqueousSolubility.fit.UNIFAC.temperature = EnglishValue
                 
                 tempaqopttmp = phprop.AqueousSolubility.operatingT.UNIFAC.temperature
                 Call TEMPCNV(EnglishValue, phprop.AqueousSolubility.operatingT.UNIFAC.temperature)
                 phprop.AqueousSolubility.operatingT.UNIFAC.temperature = EnglishValue
       
                 tempaqdbtmp = phprop.AqueousSolubility.database.temperature
                 Call TEMPCNV(EnglishValue, phprop.AqueousSolubility.database.temperature)
                 phprop.AqueousSolubility.database.temperature = EnglishValue
                 
                 tempaquntmp = phprop.AqueousSolubility.UNIFAC.temperature
                 Call TEMPCNV(EnglishValue, phprop.AqueousSolubility.UNIFAC.temperature)
                 phprop.AqueousSolubility.UNIFAC.temperature = EnglishValue
                 
                 tempaqtmpi = phprop.AqueousSolubility.input.temperature
                 Call TEMPCNV(EnglishValue, phprop.AqueousSolubility.input.temperature)
                 phprop.AqueousSolubility.input.temperature = EnglishValue
                 
                 tempoctopttmp = phprop.OctWaterPartCoeff.operatingT.UNIFAC.temperature
                 Call TEMPCNV(EnglishValue, phprop.OctWaterPartCoeff.operatingT.UNIFAC.temperature)
                 phprop.OctWaterPartCoeff.operatingT.UNIFAC.temperature = EnglishValue
              
                 tempoctdbtmp = phprop.OctWaterPartCoeff.database.temperature
                 Call TEMPCNV(EnglishValue, phprop.OctWaterPartCoeff.database.temperature)
                 phprop.OctWaterPartCoeff.database.temperature = EnglishValue

                 tempoctuntmp = phprop.OctWaterPartCoeff.databaseT.UNIFAC.temperature
                 Call TEMPCNV(EnglishValue, phprop.OctWaterPartCoeff.databaseT.UNIFAC.temperature)
                 phprop.OctWaterPartCoeff.databaseT.UNIFAC.temperature = EnglishValue

                 tempocttmpi = phprop.OctWaterPartCoeff.input.temperature
                 Call TEMPCNV(EnglishValue, phprop.OctWaterPartCoeff.input.temperature)
                 phprop.OctWaterPartCoeff.input.temperature = EnglishValue

                 templdhltmp = phprop.LiquidDiffusivity.haydukLaudie.temperature
                 Call TEMPCNV(EnglishValue, phprop.LiquidDiffusivity.haydukLaudie.temperature)
                 phprop.LiquidDiffusivity.haydukLaudie.temperature = EnglishValue

                 templdptmp = phprop.LiquidDiffusivity.polson.temperature
                 Call TEMPCNV(EnglishValue, phprop.LiquidDiffusivity.polson.temperature)
                 phprop.LiquidDiffusivity.polson.temperature = EnglishValue
                 
                 templdwctmp = phprop.LiquidDiffusivity.wilkeChang.temperature
                 Call TEMPCNV(EnglishValue, phprop.LiquidDiffusivity.wilkeChang.temperature)
                 phprop.LiquidDiffusivity.wilkeChang.temperature = EnglishValue
                 
                 templdtmpi = phprop.LiquidDiffusivity.input.temperature
                 Call TEMPCNV(EnglishValue, phprop.LiquidDiffusivity.input.temperature)
                 phprop.LiquidDiffusivity.input.temperature = EnglishValue
                 
                 tempgdwltmp = phprop.GasDiffusivity.wilkeLee.temperature
                 Call TEMPCNV(EnglishValue, phprop.GasDiffusivity.wilkeLee.temperature)
                 phprop.GasDiffusivity.wilkeLee.temperature = EnglishValue
                 
                 tempgdtmpi = phprop.GasDiffusivity.input.temperature
                 Call TEMPCNV(EnglishValue, phprop.GasDiffusivity.input.temperature)
                 phprop.GasDiffusivity.input.temperature = EnglishValue

                 'Convert values
                 tempvp = phprop.VaporPressure.CurrentSelection.Value
                 tempvpi = phprop.VaporPressure.input.Value
                 Call VPCONV(EnglishValue, phprop.VaporPressure.CurrentSelection.Value)
                 phprop.VaporPressure.CurrentSelection.Value = EnglishValue
                 phprop.VaporPressure.database.Value = EnglishValue
                 Call VPCONV(EnglishValue, phprop.VaporPressure.input.Value)
                 phprop.VaporPressure.input.Value = EnglishValue
                 
                 tempmw = phprop.MolecularWeight.CurrentSelection.Value
                 tempmwdb = phprop.MolecularWeight.database.Value
                 tempmwun = phprop.MolecularWeight.UNIFAC.Value
                 tempmwi = phprop.MolecularWeight.input.Value
                 Call MWCONV(EnglishValue, phprop.MolecularWeight.CurrentSelection.Value)
                 phprop.MolecularWeight.CurrentSelection.Value = EnglishValue
                 Call MWCONV(EnglishValue, phprop.MolecularWeight.database.Value)
                 phprop.MolecularWeight.database.Value = EnglishValue
                 Call MWCONV(EnglishValue, phprop.MolecularWeight.UNIFAC.Value)
                 phprop.MolecularWeight.UNIFAC.Value = EnglishValue
                 Call MWCONV(EnglishValue, phprop.MolecularWeight.input.Value)
                 phprop.MolecularWeight.input.Value = EnglishValue
              
                 tempbp = phprop.BoilingPoint.CurrentSelection.Value
                 tempbpdb = phprop.BoilingPoint.database.Value
                 tempbpi = phprop.BoilingPoint.input.Value
                 Call NBPCONV(EnglishValue, phprop.BoilingPoint.CurrentSelection.Value)
                 phprop.BoilingPoint.CurrentSelection.Value = EnglishValue
                 Call NBPCONV(EnglishValue, phprop.BoilingPoint.database.Value)
                 phprop.BoilingPoint.database.Value = EnglishValue
                 Call NBPCONV(EnglishValue, phprop.BoilingPoint.input.Value)
                 phprop.BoilingPoint.input.Value = EnglishValue
                 
                 templd = phprop.LiquidDensity.CurrentSelection.Value
                 templddb = phprop.LiquidDensity.database.Value
                 templdun = phprop.LiquidDensity.UNIFAC.Value
                 templdi = phprop.LiquidDensity.input.Value
                 Call LDENSCNV(EnglishValue, phprop.LiquidDensity.CurrentSelection.Value)
                 phprop.LiquidDensity.CurrentSelection.Value = EnglishValue
                 Call LDENSCNV(EnglishValue, phprop.LiquidDensity.database.Value)
                 phprop.LiquidDensity.database.Value = EnglishValue
                 Call LDENSCNV(EnglishValue, phprop.LiquidDensity.UNIFAC.Value)
                 phprop.LiquidDensity.UNIFAC.Value = EnglishValue
                 Call LDENSCNV(EnglishValue, phprop.LiquidDensity.input.Value)
                 phprop.LiquidDensity.input.Value = EnglishValue
              
                 tempmvopt = phprop.MolarVolume.operatingT.CurrentSelection.Value
                 tempmvoptdb = phprop.MolarVolume.operatingT.database.Value
                 tempmvoptun = phprop.MolarVolume.operatingT.UNIFAC.Value
                 tempmvopti = phprop.MolarVolume.operatingT.input.Value
                 Call MVOPTCNV(EnglishValue, phprop.MolarVolume.operatingT.CurrentSelection.Value)
                 phprop.MolarVolume.operatingT.CurrentSelection.Value = EnglishValue
                 Call MVOPTCNV(EnglishValue, phprop.MolarVolume.operatingT.database.Value)
                 phprop.MolarVolume.operatingT.database.Value = EnglishValue
                 Call MVOPTCNV(EnglishValue, phprop.MolarVolume.operatingT.UNIFAC.Value)
                 phprop.MolarVolume.operatingT.UNIFAC.Value = EnglishValue
                 Call MVOPTCNV(EnglishValue, phprop.MolarVolume.operatingT.input.Value)
                 phprop.MolarVolume.operatingT.input.Value = EnglishValue
              
                 tempmv = phprop.MolarVolume.BoilingPoint.CurrentSelection.Value
                 tempmvsc = phprop.MolarVolume.BoilingPoint.UNIFAC.Value
                 tempmvi = phprop.MolarVolume.BoilingPoint.input.Value
                 Call MVNBPCNV(EnglishValue, phprop.MolarVolume.BoilingPoint.CurrentSelection.Value)
                 phprop.MolarVolume.BoilingPoint.CurrentSelection.Value = EnglishValue
                 Call MVNBPCNV(EnglishValue, phprop.MolarVolume.BoilingPoint.UNIFAC.Value)
                 phprop.MolarVolume.BoilingPoint.UNIFAC.Value = EnglishValue
                 Call MVNBPCNV(EnglishValue, phprop.MolarVolume.BoilingPoint.input.Value)
                 phprop.MolarVolume.BoilingPoint.input.Value = EnglishValue
              
                 templdiff = phprop.LiquidDiffusivity.CurrentSelection.Value
                 templhldiff = phprop.LiquidDiffusivity.haydukLaudie.Value
                 templpldiff = phprop.LiquidDiffusivity.polson.Value
                 templwcdiff = phprop.LiquidDiffusivity.wilkeChang.Value
                 templdiffi = phprop.LiquidDiffusivity.input.Value
                 Call LDIFFCNV(EnglishValue, phprop.LiquidDiffusivity.CurrentSelection.Value)
                 phprop.LiquidDiffusivity.CurrentSelection.Value = EnglishValue
                 Call LDIFFCNV(EnglishValue, phprop.LiquidDiffusivity.haydukLaudie.Value)
                 phprop.LiquidDiffusivity.haydukLaudie.Value = EnglishValue
                 Call LDIFFCNV(EnglishValue, phprop.LiquidDiffusivity.polson.Value)
                 phprop.LiquidDiffusivity.polson.Value = EnglishValue
                 Call LDIFFCNV(EnglishValue, phprop.LiquidDiffusivity.wilkeChang.Value)
                 phprop.LiquidDiffusivity.wilkeChang.Value = EnglishValue
                 Call LDIFFCNV(EnglishValue, phprop.LiquidDiffusivity.input.Value)
                 phprop.LiquidDiffusivity.input.Value = EnglishValue
              
                 tempgdiff = phprop.GasDiffusivity.CurrentSelection.Value
                 tempgdiffwl = phprop.GasDiffusivity.wilkeLee.Value
                 tempgdiffi = phprop.GasDiffusivity.input.Value
                 Call GDIFFCNV(EnglishValue, phprop.GasDiffusivity.CurrentSelection.Value)
                 phprop.GasDiffusivity.CurrentSelection.Value = EnglishValue
                 Call GDIFFCNV(EnglishValue, phprop.GasDiffusivity.wilkeLee.Value)
                 phprop.GasDiffusivity.wilkeLee.Value = EnglishValue
                 Call GDIFFCNV(EnglishValue, phprop.GasDiffusivity.input.Value)
                 phprop.GasDiffusivity.input.Value = EnglishValue
              
              End If

              For J = 1 To NUMBER_OF_PROPERTIES
                  HaveProperty(J) = phprop.HaveProperty(J)
              Next J
              For J = 1 To NUMBER_OF_PROPERTIES_AVAILABLE
                  PROPAVAILABLE(J) = phprop.PROPAVAILABLE(J)
              Next J

              If i = NumSelectedChemicals Then
                 If Printer.CurrentY + PrintHeightOfEachContaminant + BOTTOM_MARGIN_SAFETY_FACTOR > Printer.Height Then
                    Printer.NewPage
                 End If
              Else
                 If Printer.CurrentY + PrintHeightOfEachContaminant + PrintSpaceBetweenContaminants > Printer.Height Then
                    Printer.NewPage
                 End If
              End If

              Call PrintOneContaminant  'Prints the contaminant currently stored in structure phprop
              
              'If English units were selected convert them back
              If cboUnits.ListIndex = 1 Then
                 
                 'Convert temperatures back
                 phprop.VaporPressure.database.temperature = tempvptmp
                 
                 phprop.VaporPressure.input.temperature = tempvptmpi
                 
                 phprop.ActivityCoefficient.UNIFAC.temperature = tempactmp
                 
                 phprop.HenrysConstant.regress.temperature = temphregtmp
                 
                 phprop.HenrysConstant.fit.UNIFAC.temperature = temphfittmp
                 
                 phprop.HenrysConstant.operatingT.UNIFAC.temperature = temphopttmp
       
                 For J = 1 To phprop.HenrysConstant.NumberOfDatabaseHenrysConstants
                    phprop.HenrysConstant.database(J).temperature = temphdbttmp(J)
                    phprop.HenrysConstant.UNIFAC(J).temperature = temphunttmp(J)
                 Next J
                 
                 phprop.HenrysConstant.input.temperature = temphtmpi
                 
                 phprop.LiquidDensity.database.temperature = templdtmp
                 
                 phprop.LiquidDensity.UNIFAC.temperature = templdutmp
                 
                 phprop.LiquidDensity.input.temperature = templdtmpi
                 
                 phprop.MolarVolume.operatingT.database.temperature = tempmvopttmp
                 
                 phprop.MolarVolume.operatingT.UNIFAC.temperature = tempmvoptutmp
                 
                 phprop.MolarVolume.operatingT.input.temperature = tempmvopttmpi
                 
                 phprop.MolarVolume.BoilingPoint.UNIFAC.temperature = tempmvtmp
                 
                 phprop.MolarVolume.BoilingPoint.input.temperature = tempmvtmpi
                 
                 phprop.AqueousSolubility.fit.UNIFAC.temperature = tempaqfittmp
                 
                 phprop.AqueousSolubility.operatingT.UNIFAC.temperature = tempaqopttmp
       
                 phprop.AqueousSolubility.database.temperature = tempaqdbtmp
                 
                 phprop.AqueousSolubility.UNIFAC.temperature = tempaquntmp
                 
                 phprop.AqueousSolubility.input.temperature = tempaqtmpi
                 
                 phprop.OctWaterPartCoeff.operatingT.UNIFAC.temperature = tempoctopttmp
              
                 phprop.OctWaterPartCoeff.database.temperature = tempoctdbtmp

                 phprop.OctWaterPartCoeff.databaseT.UNIFAC.temperature = tempoctuntmp

                 phprop.OctWaterPartCoeff.input.temperature = tempocttmpi

                 phprop.LiquidDiffusivity.haydukLaudie.temperature = templdhltmp

                 phprop.LiquidDiffusivity.polson.temperature = templdptmp
                 
                 phprop.LiquidDiffusivity.wilkeChang.temperature = templdwctmp
                 
                 phprop.LiquidDiffusivity.input.temperature = templdtmpi
                 
                 phprop.GasDiffusivity.wilkeLee.temperature = tempgdwltmp
                 
                 phprop.GasDiffusivity.input.temperature = tempgdtmpi
              
                 'Covert values back
                 phprop.OperatingPressure = tempress
              
                 phprop.OperatingTemperature = tempt
              
                 phprop.VaporPressure.CurrentSelection.Value = tempvp
                 phprop.VaporPressure.database.Value = tempvp
                 phprop.VaporPressure.input.Value = tempvpi
              
                 phprop.MolecularWeight.CurrentSelection.Value = tempmw
                 phprop.MolecularWeight.database.Value = tempmwdb
                 phprop.MolecularWeight.UNIFAC.Value = tempmwun
                 phprop.MolecularWeight.input.Value = tempmwi
                 
                 phprop.BoilingPoint.CurrentSelection.Value = tempbp
                 phprop.BoilingPoint.database.Value = tempbpdb
                 phprop.BoilingPoint.input.Value = tempbpi
                 
                 phprop.LiquidDensity.CurrentSelection.Value = templd
                 phprop.LiquidDensity.database.Value = templddb
                 phprop.LiquidDensity.UNIFAC.Value = templdun
                 phprop.LiquidDensity.input.Value = templdi
                 
                 phprop.MolarVolume.operatingT.CurrentSelection.Value = tempmvopt
                 phprop.MolarVolume.operatingT.database.Value = tempmvoptdb
                 phprop.MolarVolume.operatingT.UNIFAC.Value = tempmvoptun
                 phprop.MolarVolume.operatingT.input.Value = tempmvopti

                 phprop.MolarVolume.BoilingPoint.CurrentSelection.Value = tempmv
                 phprop.MolarVolume.BoilingPoint.UNIFAC.Value = tempmvsc
                 phprop.MolarVolume.BoilingPoint.input.Value = tempmvi
                 
                 phprop.LiquidDiffusivity.CurrentSelection.Value = templdiff
                 phprop.LiquidDiffusivity.haydukLaudie.Value = templhldiff
                 phprop.LiquidDiffusivity.polson.Value = templpldiff
                 phprop.LiquidDiffusivity.wilkeChang.Value = templwcdiff
                 phprop.LiquidDiffusivity.input.Value = templdiffi
              
                 phprop.GasDiffusivity.CurrentSelection.Value = tempgdiff
                 phprop.GasDiffusivity.wilkeLee.Value = tempgdiffwl
                 phprop.GasDiffusivity.input.Value = tempgdiffi
              
              End If

              phprop.BoilingPoint.CurrentSelection.Value = tempbp
              
              If i = NumSelectedChemicals Then Exit For
              If i = 1 Then PrintHeightOfEachContaminant = Printer.CurrentY
              Printer.Print
              Printer.Print
              Printer.Print
              Printer.Print
              If i = 1 Then PrintSpaceBetweenContaminants = Printer.CurrentY - PrintHeightOfEachContaminant
           
          Next i

          phprop = PropContaminant(CurrentlySelectedContaminant)

          For i = 1 To NUMBER_OF_PROPERTIES
              HaveProperty(i) = phprop.HaveProperty(i)
          Next i
          For i = 1 To NUMBER_OF_PROPERTIES_AVAILABLE
              PROPAVAILABLE(i) = phprop.PROPAVAILABLE(i)
          Next i

       End If

       If ChosenAtLeastOneAirWaterProperty Then
          Printer.NewPage
       
          'If English units are desired convert them here
          If cboUnits.ListIndex = 1 Then
              
             temppress = phprop.OperatingPressure
             Call PRESSCNV(EnglishValue, phprop.OperatingPressure)
             phprop.OperatingPressure = EnglishValue
              
             tempt = phprop.OperatingTemperature
             Call TEMPCNV(EnglishValue, phprop.OperatingTemperature)
             phprop.OperatingTemperature = EnglishValue
             
             'Convert Temperatures
             tempwdtmp = phprop.WaterDensity.correlation.temperature
             tempwdtmpi = phprop.WaterDensity.input.temperature
             Call TEMPCNV(EnglishValue, phprop.WaterDensity.correlation.temperature)
             phprop.WaterDensity.correlation.temperature = EnglishValue
             Call TEMPCNV(EnglishValue, phprop.WaterDensity.input.temperature)
             phprop.WaterDensity.input.temperature = EnglishValue
             
             tempwvtmp = phprop.WaterViscosity.correlation.temperature
             tempwvtmpi = phprop.WaterViscosity.input.temperature
             Call TEMPCNV(EnglishValue, phprop.WaterViscosity.correlation.temperature)
             phprop.WaterViscosity.correlation.temperature = EnglishValue
             Call TEMPCNV(EnglishValue, phprop.WaterViscosity.input.temperature)
             phprop.WaterViscosity.input.temperature = EnglishValue
             
             tempwsttmp = phprop.WaterSurfaceTension.correlation.temperature
             tempwsttmpi = phprop.WaterSurfaceTension.input.temperature
             Call TEMPCNV(EnglishValue, phprop.WaterSurfaceTension.correlation.temperature)
             phprop.WaterSurfaceTension.correlation.temperature = EnglishValue
             Call TEMPCNV(EnglishValue, phprop.WaterSurfaceTension.input.temperature)
             phprop.WaterSurfaceTension.input.temperature = EnglishValue
             
             tempadtmp = phprop.AirDensity.correlation.temperature
             tempadtmpi = phprop.AirDensity.input.temperature
             Call TEMPCNV(EnglishValue, phprop.AirDensity.correlation.temperature)
             phprop.AirDensity.correlation.temperature = EnglishValue
             Call TEMPCNV(EnglishValue, phprop.AirDensity.input.temperature)
             phprop.AirDensity.input.temperature = EnglishValue
             
             tempavtmp = phprop.AirViscosity.correlation.temperature
             tempavtmpi = phprop.AirViscosity.input.temperature
             Call TEMPCNV(EnglishValue, phprop.AirViscosity.correlation.temperature)
             phprop.AirViscosity.correlation.temperature = EnglishValue
             Call TEMPCNV(EnglishValue, phprop.AirViscosity.input.temperature)
             phprop.AirViscosity.input.temperature = EnglishValue
              
             'Convert Values
             tempwd = phprop.WaterDensity.CurrentSelection.Value
             tempwdcor = phprop.WaterDensity.correlation.Value
             tempwdi = phprop.WaterDensity.input.Value
             Call WDENSCNV(EnglishValue, phprop.WaterDensity.CurrentSelection.Value)
             phprop.WaterDensity.CurrentSelection.Value = EnglishValue
             Call WDENSCNV(EnglishValue, phprop.WaterDensity.correlation.Value)
             phprop.WaterDensity.correlation.Value = EnglishValue
             Call WDENSCNV(EnglishValue, phprop.WaterDensity.input.Value)
             phprop.WaterDensity.input.Value = EnglishValue
             
             tempwv = phprop.WaterViscosity.CurrentSelection.Value
             tempwvcor = phprop.WaterViscosity.correlation.Value
             tempwvi = phprop.WaterViscosity.input.Value
             Call WVISCCNV(EnglishValue, phprop.WaterViscosity.CurrentSelection.Value)
             phprop.WaterViscosity.CurrentSelection.Value = EnglishValue
             Call WVISCCNV(EnglishValue, phprop.WaterViscosity.correlation.Value)
             phprop.WaterViscosity.correlation.Value = EnglishValue
             Call WVISCCNV(EnglishValue, phprop.WaterViscosity.input.Value)
             phprop.WaterViscosity.input.Value = EnglishValue
             
             tempwst = phprop.WaterSurfaceTension.CurrentSelection.Value
             tempwstcor = phprop.WaterSurfaceTension.correlation.Value
             tempwsti = phprop.WaterSurfaceTension.input.Value
             Call H2OSTCNV(EnglishValue, phprop.WaterSurfaceTension.CurrentSelection.Value)
             phprop.WaterSurfaceTension.CurrentSelection.Value = EnglishValue
             Call H2OSTCNV(EnglishValue, phprop.WaterSurfaceTension.correlation.Value)
             phprop.WaterSurfaceTension.correlation.Value = EnglishValue
             Call H2OSTCNV(EnglishValue, phprop.WaterSurfaceTension.input.Value)
             phprop.WaterSurfaceTension.input.Value = EnglishValue
             
             tempad = phprop.AirDensity.CurrentSelection.Value
             tempadcor = phprop.AirDensity.correlation.Value
             tempadi = phprop.AirDensity.input.Value
             Call ADENSCNV(EnglishValue, phprop.AirDensity.CurrentSelection.Value)
             phprop.AirDensity.CurrentSelection.Value = EnglishValue
             Call ADENSCNV(EnglishValue, phprop.AirDensity.correlation.Value)
             phprop.AirDensity.correlation.Value = EnglishValue
             Call ADENSCNV(EnglishValue, phprop.AirDensity.input.Value)
             phprop.AirDensity.input.Value = EnglishValue
             
             tempav = phprop.AirViscosity.CurrentSelection.Value
             tempavcor = phprop.AirViscosity.correlation.Value
             tempavi = phprop.AirViscosity.input.Value
             Call AVISCCNV(EnglishValue, phprop.AirViscosity.CurrentSelection.Value)
             phprop.AirViscosity.CurrentSelection.Value = EnglishValue
             Call AVISCCNV(EnglishValue, phprop.AirViscosity.correlation.Value)
             phprop.AirViscosity.correlation.Value = EnglishValue
             Call AVISCCNV(EnglishValue, phprop.AirViscosity.input.Value)
             phprop.AirViscosity.input.Value = EnglishValue

          End If
          
          Call PrintAirWaterProperties
       
          'If English units were selected convert them back
          If cboUnits.ListIndex = 1 Then
              
             phprop.OperatingPressure = temppress
              
             phprop.OperatingTemperature = tempt
             
             'Convert temperatures back
             
             phprop.WaterDensity.correlation.temperature = tempwdtmp
             phprop.WaterDensity.input.temperature = tempwdtmpi
             
             phprop.WaterViscosity.correlation.temperature = tempwvtmp
             phprop.WaterViscosity.input.temperature = tempwvtmpi
             
             phprop.WaterSurfaceTension.correlation.temperature = tempwsttmp
             phprop.WaterSurfaceTension.input.temperature = tempwsttmpi
             
             phprop.AirDensity.correlation.temperature = tempadtmp
             phprop.AirDensity.input.temperature = tempadtmpi
             
             phprop.AirViscosity.correlation.temperature = tempavtmp
             phprop.AirViscosity.input.temperature = tempavtmpi
              
             'Convert values back
             phprop.WaterDensity.CurrentSelection.Value = tempwd
             phprop.WaterDensity.correlation.Value = tempwdcor
             phprop.WaterDensity.input.Value = tempwdi
             
             phprop.WaterViscosity.CurrentSelection.Value = tempwv
             phprop.WaterViscosity.correlation.Value = tempwvcor
             phprop.WaterViscosity.input.Value = tempwvi
             
             phprop.WaterSurfaceTension.CurrentSelection.Value = tempwst
             phprop.WaterSurfaceTension.correlation.Value = tempwstcor
             phprop.WaterSurfaceTension.input.Value = tempwsti
             
             phprop.AirDensity.CurrentSelection.Value = tempad
             phprop.AirDensity.correlation.Value = tempadcor
             phprop.AirDensity.input.Value = tempadi
             
             phprop.AirViscosity.CurrentSelection.Value = tempav
             phprop.AirViscosity.correlation.Value = tempavcor
             phprop.AirViscosity.input.Value = tempavi

          End If
       
       End If
       
    ElseIf frmPrint!optPrintContaminants(1).Value Then   'Print Currently Selected Contaminant
       If ChosenAtLeastOneContaminantProperty Then
          
          'Convert******************************************************

              'If English units are desired convert them here
              If cboUnits.ListIndex = 1 Then
              
                 temppress = phprop.OperatingPressure
                 Call PRESSCNV(EnglishValue, phprop.OperatingPressure)
                 phprop.OperatingPressure = EnglishValue
              
                 tempt = phprop.OperatingTemperature
                 Call TEMPCNV(EnglishValue, phprop.OperatingTemperature)
                 phprop.OperatingTemperature = EnglishValue
                 
                 'Change all temperatures
                 tempvptmp = phprop.VaporPressure.database.temperature
                 Call TEMPCNV(EnglishValue, phprop.VaporPressure.database.temperature)
                 phprop.VaporPressure.database.temperature = EnglishValue
                 
                 tempvptmpi = phprop.VaporPressure.input.temperature
                 Call TEMPCNV(EnglishValue, phprop.VaporPressure.input.temperature)
                 phprop.VaporPressure.input.temperature = EnglishValue
                 
                 tempactmp = phprop.ActivityCoefficient.UNIFAC.temperature
                 Call TEMPCNV(EnglishValue, phprop.ActivityCoefficient.UNIFAC.temperature)
                 phprop.ActivityCoefficient.UNIFAC.temperature = EnglishValue
                 
                 temphregtmp = phprop.HenrysConstant.regress.temperature
                 Call TEMPCNV(EnglishValue, phprop.HenrysConstant.regress.temperature)
                 phprop.HenrysConstant.regress.temperature = EnglishValue
                 
                 temphfittmp = phprop.HenrysConstant.fit.UNIFAC.temperature
                 Call TEMPCNV(EnglishValue, phprop.HenrysConstant.fit.UNIFAC.temperature)
                 phprop.HenrysConstant.fit.UNIFAC.temperature = EnglishValue
                 
                 temphopttmp = phprop.HenrysConstant.operatingT.UNIFAC.temperature
                 Call TEMPCNV(EnglishValue, phprop.HenrysConstant.operatingT.UNIFAC.temperature)
                 phprop.HenrysConstant.operatingT.UNIFAC.temperature = EnglishValue
       
                 For J = 1 To phprop.HenrysConstant.NumberOfDatabaseHenrysConstants
                    temphdbttmp(J) = phprop.HenrysConstant.database(J).temperature
                    temphunttmp(J) = phprop.HenrysConstant.UNIFAC(J).temperature
                    Call TEMPCNV(EnglishValue, phprop.HenrysConstant.database(J).temperature)
                    phprop.HenrysConstant.database(J).temperature = EnglishValue
                    Call TEMPCNV(EnglishValue, phprop.HenrysConstant.UNIFAC(J).temperature)
                    phprop.HenrysConstant.UNIFAC(J).temperature = EnglishValue
                 Next J
                 
                 temphtmpi = phprop.HenrysConstant.input.temperature
                 Call TEMPCNV(EnglishValue, phprop.HenrysConstant.input.temperature)
                 phprop.HenrysConstant.input.temperature = EnglishValue
                 
                 templdtmp = phprop.LiquidDensity.database.temperature
                 Call TEMPCNV(EnglishValue, phprop.LiquidDensity.database.temperature)
                 phprop.LiquidDensity.database.temperature = EnglishValue
                 
                 templdutmp = phprop.LiquidDensity.UNIFAC.temperature
                 Call TEMPCNV(EnglishValue, phprop.LiquidDensity.UNIFAC.temperature)
                 phprop.LiquidDensity.UNIFAC.temperature = EnglishValue
                 
                 templdtmpi = phprop.LiquidDensity.input.temperature
                 Call TEMPCNV(EnglishValue, phprop.LiquidDensity.input.temperature)
                 phprop.LiquidDensity.input.temperature = EnglishValue
                 
                 tempmvopttmp = phprop.MolarVolume.operatingT.database.temperature
                 Call TEMPCNV(EnglishValue, phprop.MolarVolume.operatingT.database.temperature)
                 phprop.MolarVolume.operatingT.database.temperature = EnglishValue
                 
                 tempmvoptutmp = phprop.MolarVolume.operatingT.UNIFAC.temperature
                 Call TEMPCNV(EnglishValue, phprop.MolarVolume.operatingT.UNIFAC.temperature)
                 phprop.MolarVolume.operatingT.UNIFAC.temperature = EnglishValue
                 
                 tempmvopttmpi = phprop.MolarVolume.operatingT.input.temperature
                 Call TEMPCNV(EnglishValue, phprop.MolarVolume.operatingT.input.temperature)
                 phprop.MolarVolume.operatingT.input.temperature = EnglishValue
                 
                 tempmvtmp = phprop.MolarVolume.BoilingPoint.UNIFAC.temperature
                 Call TEMPCNV(EnglishValue, phprop.MolarVolume.BoilingPoint.UNIFAC.temperature)
                 phprop.MolarVolume.BoilingPoint.UNIFAC.temperature = EnglishValue
                 
                 tempmvtmpi = phprop.MolarVolume.BoilingPoint.input.temperature
                 Call TEMPCNV(EnglishValue, phprop.MolarVolume.BoilingPoint.input.temperature)
                 phprop.MolarVolume.BoilingPoint.input.temperature = EnglishValue
                 
                 tempaqfittmp = phprop.AqueousSolubility.fit.UNIFAC.temperature
                 Call TEMPCNV(EnglishValue, phprop.AqueousSolubility.fit.UNIFAC.temperature)
                 phprop.AqueousSolubility.fit.UNIFAC.temperature = EnglishValue
                 
                 tempaqopttmp = phprop.AqueousSolubility.operatingT.UNIFAC.temperature
                 Call TEMPCNV(EnglishValue, phprop.AqueousSolubility.operatingT.UNIFAC.temperature)
                 phprop.AqueousSolubility.operatingT.UNIFAC.temperature = EnglishValue
       
                 tempaqdbtmp = phprop.AqueousSolubility.database.temperature
                 Call TEMPCNV(EnglishValue, phprop.AqueousSolubility.database.temperature)
                 phprop.AqueousSolubility.database.temperature = EnglishValue
                 
                 tempaquntmp = phprop.AqueousSolubility.UNIFAC.temperature
                 Call TEMPCNV(EnglishValue, phprop.AqueousSolubility.UNIFAC.temperature)
                 phprop.AqueousSolubility.UNIFAC.temperature = EnglishValue
                 
                 tempaqtmpi = phprop.AqueousSolubility.input.temperature
                 Call TEMPCNV(EnglishValue, phprop.AqueousSolubility.input.temperature)
                 phprop.AqueousSolubility.input.temperature = EnglishValue
                 
                 tempoctopttmp = phprop.OctWaterPartCoeff.operatingT.UNIFAC.temperature
                 Call TEMPCNV(EnglishValue, phprop.OctWaterPartCoeff.operatingT.UNIFAC.temperature)
                 phprop.OctWaterPartCoeff.operatingT.UNIFAC.temperature = EnglishValue
              
                 tempoctdbtmp = phprop.OctWaterPartCoeff.database.temperature
                 Call TEMPCNV(EnglishValue, phprop.OctWaterPartCoeff.database.temperature)
                 phprop.OctWaterPartCoeff.database.temperature = EnglishValue

                 tempoctuntmp = phprop.OctWaterPartCoeff.databaseT.UNIFAC.temperature
                 Call TEMPCNV(EnglishValue, phprop.OctWaterPartCoeff.databaseT.UNIFAC.temperature)
                 phprop.OctWaterPartCoeff.databaseT.UNIFAC.temperature = EnglishValue

                 tempocttmpi = phprop.OctWaterPartCoeff.input.temperature
                 Call TEMPCNV(EnglishValue, phprop.OctWaterPartCoeff.input.temperature)
                 phprop.OctWaterPartCoeff.input.temperature = EnglishValue

                 templdhltmp = phprop.LiquidDiffusivity.haydukLaudie.temperature
                 Call TEMPCNV(EnglishValue, phprop.LiquidDiffusivity.haydukLaudie.temperature)
                 phprop.LiquidDiffusivity.haydukLaudie.temperature = EnglishValue

                 templdptmp = phprop.LiquidDiffusivity.polson.temperature
                 Call TEMPCNV(EnglishValue, phprop.LiquidDiffusivity.polson.temperature)
                 phprop.LiquidDiffusivity.polson.temperature = EnglishValue
                 
                 templdwctmp = phprop.LiquidDiffusivity.wilkeChang.temperature
                 Call TEMPCNV(EnglishValue, phprop.LiquidDiffusivity.wilkeChang.temperature)
                 phprop.LiquidDiffusivity.wilkeChang.temperature = EnglishValue
                 
                 templdtmpi = phprop.LiquidDiffusivity.input.temperature
                 Call TEMPCNV(EnglishValue, phprop.LiquidDiffusivity.input.temperature)
                 phprop.LiquidDiffusivity.input.temperature = EnglishValue
                 
                 tempgdwltmp = phprop.GasDiffusivity.wilkeLee.temperature
                 Call TEMPCNV(EnglishValue, phprop.GasDiffusivity.wilkeLee.temperature)
                 phprop.GasDiffusivity.wilkeLee.temperature = EnglishValue
                 
                 tempgdtmpi = phprop.GasDiffusivity.input.temperature
                 Call TEMPCNV(EnglishValue, phprop.GasDiffusivity.input.temperature)
                 phprop.GasDiffusivity.input.temperature = EnglishValue

                 'Convert values
                 tempvp = phprop.VaporPressure.CurrentSelection.Value
                 tempvpi = phprop.VaporPressure.input.Value
                 Call VPCONV(EnglishValue, phprop.VaporPressure.CurrentSelection.Value)
                 phprop.VaporPressure.CurrentSelection.Value = EnglishValue
                 phprop.VaporPressure.database.Value = EnglishValue
                 Call VPCONV(EnglishValue, phprop.VaporPressure.input.Value)
                 phprop.VaporPressure.input.Value = EnglishValue
              
                 tempmw = phprop.MolecularWeight.CurrentSelection.Value
                 tempmwdb = phprop.MolecularWeight.database.Value
                 tempmwun = phprop.MolecularWeight.UNIFAC.Value
                 tempmwi = phprop.MolecularWeight.input.Value
                 Call MWCONV(EnglishValue, phprop.MolecularWeight.CurrentSelection.Value)
                 phprop.MolecularWeight.CurrentSelection.Value = EnglishValue
                 Call MWCONV(EnglishValue, phprop.MolecularWeight.database.Value)
                 phprop.MolecularWeight.database.Value = EnglishValue
                 Call MWCONV(EnglishValue, phprop.MolecularWeight.UNIFAC.Value)
                 phprop.MolecularWeight.UNIFAC.Value = EnglishValue
                 Call MWCONV(EnglishValue, phprop.MolecularWeight.input.Value)
                 phprop.MolecularWeight.input.Value = EnglishValue
                 
                 tempbp = phprop.BoilingPoint.CurrentSelection.Value
                 tempbpdb = phprop.BoilingPoint.database.Value
                 tempbpi = phprop.BoilingPoint.input.Value
                 Call NBPCONV(EnglishValue, phprop.BoilingPoint.CurrentSelection.Value)
                 phprop.BoilingPoint.CurrentSelection.Value = EnglishValue
                 Call NBPCONV(EnglishValue, phprop.BoilingPoint.database.Value)
                 phprop.BoilingPoint.database.Value = EnglishValue
                 Call NBPCONV(EnglishValue, phprop.BoilingPoint.input.Value)
                 phprop.BoilingPoint.input.Value = EnglishValue
                 
                 templd = phprop.LiquidDensity.CurrentSelection.Value
                 templddb = phprop.LiquidDensity.database.Value
                 templdun = phprop.LiquidDensity.UNIFAC.Value
                 templdi = phprop.LiquidDensity.input.Value
                 Call LDENSCNV(EnglishValue, phprop.LiquidDensity.CurrentSelection.Value)
                 phprop.LiquidDensity.CurrentSelection.Value = EnglishValue
                 Call LDENSCNV(EnglishValue, phprop.LiquidDensity.database.Value)
                 phprop.LiquidDensity.database.Value = EnglishValue
                 Call LDENSCNV(EnglishValue, phprop.LiquidDensity.UNIFAC.Value)
                 phprop.LiquidDensity.UNIFAC.Value = EnglishValue
                 Call LDENSCNV(EnglishValue, phprop.LiquidDensity.input.Value)
                 phprop.LiquidDensity.input.Value = EnglishValue
                 
                 tempmvopt = phprop.MolarVolume.operatingT.CurrentSelection.Value
                 tempmvoptdb = phprop.MolarVolume.operatingT.database.Value
                 tempmvoptun = phprop.MolarVolume.operatingT.UNIFAC.Value
                 tempmvopti = phprop.MolarVolume.operatingT.input.Value
                 Call MVOPTCNV(EnglishValue, phprop.MolarVolume.operatingT.CurrentSelection.Value)
                 phprop.MolarVolume.operatingT.CurrentSelection.Value = EnglishValue
                 Call MVOPTCNV(EnglishValue, phprop.MolarVolume.operatingT.database.Value)
                 phprop.MolarVolume.operatingT.database.Value = EnglishValue
                 Call MVOPTCNV(EnglishValue, phprop.MolarVolume.operatingT.UNIFAC.Value)
                 phprop.MolarVolume.operatingT.UNIFAC.Value = EnglishValue
                 Call MVOPTCNV(EnglishValue, phprop.MolarVolume.operatingT.input.Value)
                 phprop.MolarVolume.operatingT.input.Value = EnglishValue
                 
                 tempmv = phprop.MolarVolume.BoilingPoint.CurrentSelection.Value
                 tempmvsc = phprop.MolarVolume.BoilingPoint.UNIFAC.Value
                 tempmvi = phprop.MolarVolume.BoilingPoint.input.Value
                 Call MVNBPCNV(EnglishValue, phprop.MolarVolume.BoilingPoint.CurrentSelection.Value)
                 phprop.MolarVolume.BoilingPoint.CurrentSelection.Value = EnglishValue
                 Call MVNBPCNV(EnglishValue, phprop.MolarVolume.BoilingPoint.UNIFAC.Value)
                 phprop.MolarVolume.BoilingPoint.UNIFAC.Value = EnglishValue
                 Call MVNBPCNV(EnglishValue, phprop.MolarVolume.BoilingPoint.input.Value)
                 phprop.MolarVolume.BoilingPoint.input.Value = EnglishValue
              
                 templdiff = phprop.LiquidDiffusivity.CurrentSelection.Value
                 templhldiff = phprop.LiquidDiffusivity.haydukLaudie.Value
                 templpldiff = phprop.LiquidDiffusivity.polson.Value
                 templwcdiff = phprop.LiquidDiffusivity.wilkeChang.Value
                 templdiffi = phprop.LiquidDiffusivity.input.Value
                 Call LDIFFCNV(EnglishValue, phprop.LiquidDiffusivity.CurrentSelection.Value)
                 phprop.LiquidDiffusivity.CurrentSelection.Value = EnglishValue
                 Call LDIFFCNV(EnglishValue, phprop.LiquidDiffusivity.haydukLaudie.Value)
                 phprop.LiquidDiffusivity.haydukLaudie.Value = EnglishValue
                 Call LDIFFCNV(EnglishValue, phprop.LiquidDiffusivity.polson.Value)
                 phprop.LiquidDiffusivity.polson.Value = EnglishValue
                 Call LDIFFCNV(EnglishValue, phprop.LiquidDiffusivity.wilkeChang.Value)
                 phprop.LiquidDiffusivity.wilkeChang.Value = EnglishValue
                 Call LDIFFCNV(EnglishValue, phprop.LiquidDiffusivity.input.Value)
                 phprop.LiquidDiffusivity.input.Value = EnglishValue
                 
                 tempgdiff = phprop.GasDiffusivity.CurrentSelection.Value
                 tempgdiffwl = phprop.GasDiffusivity.wilkeLee.Value
                 tempgdiffi = phprop.GasDiffusivity.input.Value
                 Call GDIFFCNV(EnglishValue, phprop.GasDiffusivity.CurrentSelection.Value)
                 phprop.GasDiffusivity.CurrentSelection.Value = EnglishValue
                 Call GDIFFCNV(EnglishValue, phprop.GasDiffusivity.wilkeLee.Value)
                 phprop.GasDiffusivity.wilkeLee.Value = EnglishValue
                 Call GDIFFCNV(EnglishValue, phprop.GasDiffusivity.input.Value)
                 phprop.GasDiffusivity.input.Value = EnglishValue
                 
              End If

          'EndConvert****************************************************

          Call PrintOneContaminant
              
          'ReConvert******************************************************

              'If English units were selected convert them back
              If cboUnits.ListIndex = 1 Then
                 
                 'Convert temperatures back
                 phprop.VaporPressure.database.temperature = tempvptmp
                 
                 phprop.VaporPressure.input.temperature = tempvptmpi
                 
                 phprop.ActivityCoefficient.UNIFAC.temperature = tempactmp
                 
                 phprop.HenrysConstant.regress.temperature = temphregtmp
                 
                 phprop.HenrysConstant.fit.UNIFAC.temperature = temphfittmp
                 
                 phprop.HenrysConstant.operatingT.UNIFAC.temperature = temphopttmp
       
                 For J = 1 To phprop.HenrysConstant.NumberOfDatabaseHenrysConstants
                    phprop.HenrysConstant.database(J).temperature = temphdbttmp(J)
                    phprop.HenrysConstant.UNIFAC(J).temperature = temphunttmp(J)
                 Next J
                 
                 phprop.HenrysConstant.input.temperature = temphtmpi
                 
                 phprop.LiquidDensity.database.temperature = templdtmp
                 
                 phprop.LiquidDensity.UNIFAC.temperature = templdutmp
                 
                 phprop.LiquidDensity.input.temperature = templdtmpi
                 
                 phprop.MolarVolume.operatingT.database.temperature = tempmvopttmp
                 
                 phprop.MolarVolume.operatingT.UNIFAC.temperature = tempmvoptutmp
                 
                 phprop.MolarVolume.operatingT.input.temperature = tempmvopttmpi
                 
                 phprop.MolarVolume.BoilingPoint.UNIFAC.temperature = tempmvtmp
                 
                 phprop.MolarVolume.BoilingPoint.input.temperature = tempmvtmpi
                 
                 phprop.AqueousSolubility.fit.UNIFAC.temperature = tempaqfittmp
                 
                 phprop.AqueousSolubility.operatingT.UNIFAC.temperature = tempaqopttmp
       
                 phprop.AqueousSolubility.database.temperature = tempaqdbtmp
                 
                 phprop.AqueousSolubility.UNIFAC.temperature = tempaquntmp
                 
                 phprop.AqueousSolubility.input.temperature = tempaqtmpi
                 
                 phprop.OctWaterPartCoeff.operatingT.UNIFAC.temperature = tempoctopttmp
              
                 phprop.OctWaterPartCoeff.database.temperature = tempoctdbtmp

                 phprop.OctWaterPartCoeff.databaseT.UNIFAC.temperature = tempoctuntmp

                 phprop.OctWaterPartCoeff.input.temperature = tempocttmpi

                 phprop.LiquidDiffusivity.haydukLaudie.temperature = templdhltmp

                 phprop.LiquidDiffusivity.polson.temperature = templdptmp
                 
                 phprop.LiquidDiffusivity.wilkeChang.temperature = templdwctmp
                 
                 phprop.LiquidDiffusivity.input.temperature = templdtmpi
                 
                 phprop.GasDiffusivity.wilkeLee.temperature = tempgdwltmp
                 
                 phprop.GasDiffusivity.input.temperature = tempgdtmpi
              
                 'Covert values back
                 phprop.OperatingPressure = temppress
              
                 phprop.OperatingTemperature = tempt
              
                 phprop.VaporPressure.CurrentSelection.Value = tempvp
                 phprop.VaporPressure.database.Value = tempvp
                 phprop.VaporPressure.input.Value = tempvpi
                 
                 phprop.MolecularWeight.CurrentSelection.Value = tempmw
                 phprop.MolecularWeight.database.Value = tempmwdb
                 phprop.MolecularWeight.UNIFAC.Value = tempmwun
                 phprop.MolecularWeight.input.Value = tempmwi
              
                 phprop.MolecularWeight.CurrentSelection.Value = tempmw
                 phprop.MolecularWeight.database.Value = tempmw
                 phprop.MolecularWeight.UNIFAC.Value = tempmw
                 phprop.MolecularWeight.input.Value = tempmwi
              
                 phprop.BoilingPoint.CurrentSelection.Value = tempbp
                 phprop.BoilingPoint.database.Value = tempbpdb
                 phprop.BoilingPoint.input.Value = tempbpi
                 
                 phprop.LiquidDensity.CurrentSelection.Value = templd
                 phprop.LiquidDensity.database.Value = templddb
                 phprop.LiquidDensity.UNIFAC.Value = templdun
                 phprop.LiquidDensity.input.Value = templdi
                 
                 phprop.MolarVolume.operatingT.CurrentSelection.Value = tempmvopt
                 phprop.MolarVolume.operatingT.database.Value = tempmvoptdb
                 phprop.MolarVolume.operatingT.UNIFAC.Value = tempmvoptun
                 phprop.MolarVolume.operatingT.input.Value = tempmvopti

                 phprop.MolarVolume.BoilingPoint.CurrentSelection.Value = tempmv
                 phprop.MolarVolume.BoilingPoint.UNIFAC.Value = tempmvsc
                 phprop.MolarVolume.BoilingPoint.input.Value = tempmvi
                 
                 phprop.LiquidDiffusivity.CurrentSelection.Value = templdiff
                 phprop.LiquidDiffusivity.haydukLaudie.Value = templhldiff
                 phprop.LiquidDiffusivity.polson.Value = templpldiff
                 phprop.LiquidDiffusivity.wilkeChang.Value = templwcdiff
                 phprop.LiquidDiffusivity.input.Value = templdiffi
              
                 phprop.GasDiffusivity.CurrentSelection.Value = tempgdiff
                 phprop.GasDiffusivity.wilkeLee.Value = tempgdiffwl
                 phprop.GasDiffusivity.input.Value = tempgdiffi
              
              End If

          'EndReConvert****************************************************

       End If

       If ChosenAtLeastOneAirWaterProperty Then
          Printer.NewPage
          
'ConvertAirWater******************************************************

          'If English units are desired convert them here
          If cboUnits.ListIndex = 1 Then
              
             temppress = phprop.OperatingPressure
             Call PRESSCNV(EnglishValue, phprop.OperatingPressure)
             phprop.OperatingPressure = EnglishValue
              
             tempt = phprop.OperatingTemperature
             Call TEMPCNV(EnglishValue, phprop.OperatingTemperature)
             phprop.OperatingTemperature = EnglishValue
             
             'Convert Temperatures
             tempwdtmp = phprop.WaterDensity.correlation.temperature
             tempwdtmpi = phprop.WaterDensity.input.temperature
             Call TEMPCNV(EnglishValue, phprop.WaterDensity.correlation.temperature)
             phprop.WaterDensity.correlation.temperature = EnglishValue
             Call TEMPCNV(EnglishValue, phprop.WaterDensity.input.temperature)
             phprop.WaterDensity.input.temperature = EnglishValue
             
             tempwvtmp = phprop.WaterViscosity.correlation.temperature
             tempwvtmpi = phprop.WaterViscosity.input.temperature
             Call TEMPCNV(EnglishValue, phprop.WaterViscosity.correlation.temperature)
             phprop.WaterViscosity.correlation.temperature = EnglishValue
             Call TEMPCNV(EnglishValue, phprop.WaterViscosity.input.temperature)
             phprop.WaterViscosity.input.temperature = EnglishValue
             
             tempwsttmp = phprop.WaterSurfaceTension.correlation.temperature
             tempwsttmpi = phprop.WaterSurfaceTension.input.temperature
             Call TEMPCNV(EnglishValue, phprop.WaterSurfaceTension.correlation.temperature)
             phprop.WaterSurfaceTension.correlation.temperature = EnglishValue
             Call TEMPCNV(EnglishValue, phprop.WaterSurfaceTension.input.temperature)
             phprop.WaterSurfaceTension.input.temperature = EnglishValue
             
             tempadtmp = phprop.AirDensity.correlation.temperature
             tempadtmpi = phprop.AirDensity.input.temperature
             Call TEMPCNV(EnglishValue, phprop.AirDensity.correlation.temperature)
             phprop.AirDensity.correlation.temperature = EnglishValue
             Call TEMPCNV(EnglishValue, phprop.AirDensity.input.temperature)
             phprop.AirDensity.input.temperature = EnglishValue
             
             tempavtmp = phprop.AirViscosity.correlation.temperature
             tempavtmpi = phprop.AirViscosity.input.temperature
             Call TEMPCNV(EnglishValue, phprop.AirViscosity.correlation.temperature)
             phprop.AirViscosity.correlation.temperature = EnglishValue
             Call TEMPCNV(EnglishValue, phprop.AirViscosity.input.temperature)
             phprop.AirViscosity.input.temperature = EnglishValue
              
             'Covert Values
             tempwd = phprop.WaterDensity.CurrentSelection.Value
             tempwdcor = phprop.WaterDensity.correlation.Value
             tempwdi = phprop.WaterDensity.input.Value
             Call WDENSCNV(EnglishValue, phprop.WaterDensity.CurrentSelection.Value)
             phprop.WaterDensity.CurrentSelection.Value = EnglishValue
             Call WDENSCNV(EnglishValue, phprop.WaterDensity.correlation.Value)
             phprop.WaterDensity.correlation.Value = EnglishValue
             Call WDENSCNV(EnglishValue, phprop.WaterDensity.input.Value)
             phprop.WaterDensity.input.Value = EnglishValue
             
             tempwv = phprop.WaterViscosity.CurrentSelection.Value
             tempwvcor = phprop.WaterViscosity.correlation.Value
             tempwvi = phprop.WaterViscosity.input.Value
             Call WVISCCNV(EnglishValue, phprop.WaterViscosity.CurrentSelection.Value)
             phprop.WaterViscosity.CurrentSelection.Value = EnglishValue
             Call WVISCCNV(EnglishValue, phprop.WaterViscosity.correlation.Value)
             phprop.WaterViscosity.correlation.Value = EnglishValue
             Call WVISCCNV(EnglishValue, phprop.WaterViscosity.input.Value)
             phprop.WaterViscosity.input.Value = EnglishValue
             
             tempwst = phprop.WaterSurfaceTension.CurrentSelection.Value
             tempwstcor = phprop.WaterSurfaceTension.correlation.Value
             tempwsti = phprop.WaterSurfaceTension.input.Value
             Call H2OSTCNV(EnglishValue, phprop.WaterSurfaceTension.CurrentSelection.Value)
             phprop.WaterSurfaceTension.CurrentSelection.Value = EnglishValue
             Call H2OSTCNV(EnglishValue, phprop.WaterSurfaceTension.correlation.Value)
             phprop.WaterSurfaceTension.correlation.Value = EnglishValue
             Call H2OSTCNV(EnglishValue, phprop.WaterSurfaceTension.input.Value)
             phprop.WaterSurfaceTension.input.Value = EnglishValue
             
             tempad = phprop.AirDensity.CurrentSelection.Value
             tempadcor = phprop.AirDensity.correlation.Value
             tempadi = phprop.AirDensity.input.Value
             Call ADENSCNV(EnglishValue, phprop.AirDensity.CurrentSelection.Value)
             phprop.AirDensity.CurrentSelection.Value = EnglishValue
             Call ADENSCNV(EnglishValue, phprop.AirDensity.correlation.Value)
             phprop.AirDensity.correlation.Value = EnglishValue
             Call ADENSCNV(EnglishValue, phprop.AirDensity.input.Value)
             phprop.AirDensity.input.Value = EnglishValue
             
             tempav = phprop.AirViscosity.CurrentSelection.Value
             tempavcor = phprop.AirViscosity.correlation.Value
             tempavi = phprop.AirViscosity.input.Value
             Call AVISCCNV(EnglishValue, phprop.AirViscosity.CurrentSelection.Value)
             phprop.AirViscosity.CurrentSelection.Value = EnglishValue
             Call AVISCCNV(EnglishValue, phprop.AirViscosity.correlation.Value)
             phprop.AirViscosity.correlation.Value = EnglishValue
             Call AVISCCNV(EnglishValue, phprop.AirViscosity.input.Value)
             phprop.AirViscosity.input.Value = EnglishValue

          End If

'EndConvertAirWater****************************************************
          
          Call PrintAirWaterProperties
              
'ReConvertAirWater******************************************************

          'If English units were selected convert them back
          If cboUnits.ListIndex = 1 Then
              
             phprop.OperatingPressure = temppress
              
             phprop.OperatingTemperature = tempt
             
             'Convert temperatures back
             phprop.WaterDensity.correlation.temperature = tempwdtmp
             phprop.WaterDensity.input.temperature = tempwdtmpi

             phprop.WaterViscosity.correlation.temperature = tempwvtmp
             phprop.WaterViscosity.input.temperature = tempwvtmpi
             
             phprop.WaterSurfaceTension.correlation.temperature = tempwsttmp
             phprop.WaterSurfaceTension.input.temperature = tempwsttmpi
             
             phprop.AirDensity.correlation.temperature = tempadtmp
             phprop.AirDensity.input.temperature = tempadtmpi
             
             phprop.AirViscosity.correlation.temperature = tempavtmp
             phprop.AirViscosity.input.temperature = tempavtmpi
              
             'Convert values back
             phprop.WaterDensity.CurrentSelection.Value = tempwd
             phprop.WaterDensity.correlation.Value = tempwdcor
             phprop.WaterDensity.input.Value = tempwdi
             
             phprop.WaterViscosity.CurrentSelection.Value = tempwv
             phprop.WaterViscosity.correlation.Value = tempwvcor
             phprop.WaterViscosity.input.Value = tempwvi
             
             phprop.WaterSurfaceTension.CurrentSelection.Value = tempwst
             phprop.WaterSurfaceTension.correlation.Value = tempwstcor
             phprop.WaterSurfaceTension.input.Value = tempwsti
             
             phprop.AirDensity.CurrentSelection.Value = tempad
             phprop.AirDensity.correlation.Value = tempadcor
             phprop.AirDensity.input.Value = tempadi
             
             phprop.AirViscosity.CurrentSelection.Value = tempav
             phprop.AirViscosity.correlation.Value = tempavcor
             phprop.AirViscosity.input.Value = tempavi

          End If

'EndReConvertAirWater****************************************************

       End If

    End If
    Printer.EndDoc

resume_exit37:
Exit Sub

error_printtoprinter:
MsgBox "Error printing to printer", mb_exclamation, "StEPP"
Resume resume_exit37


End Sub

Private Sub PrintVaporPressurePrinter()
   Dim ValueString As String

    Select Case frmPrint!cboPropertyDescription.ListIndex
       Case 0   'Print Selected Value Only
          If HaveProperty(VAPOR_PRESSURE) Then
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = Format$(phprop.VaporPressure.CurrentSelection.Value, GetTheFormat(phprop.VaporPressure.CurrentSelection.Value))
             Printer.Print "Vapor Pressure"; Tab(TABVALUE); ValueString; Tab(TABUNITS); Units(VAPOR_PRESSURE); Tab(TABSOURCE); GetSource(phprop.VaporPressure.CurrentSelection.Source)
          Else
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = "Not Available"
             Printer.Print "Vapor Pressure"; Tab(TABVALUE); ValueString
          End If

       Case 1   'Print Full Description of Vapor Pressure
          HeightVaporPressure = 0
          Printer.FontSize = 12
          PrintMsg = ""
          HeightVaporPressure = HeightVaporPressure + NUMLINES_PROPERTY_NAME * Printer.TextHeight(PrintMsg)
          Printer.FontSize = 10
          HeightVaporPressure = HeightVaporPressure + NUMLINES_VAPOR_PRESSURE * Printer.TextHeight(PrintMsg)
          TotalHeightThisPage = Printer.CurrentY
          If TotalHeightThisPage + HeightVaporPressure + BOTTOM_MARGIN_SAFETY_FACTOR > Printer.Height Then
             Printer.NewPage
             Call PrintTitleContinuation
          End If

          Printer.FontBold = True
          Printer.FontSize = 12
          Printer.Print "Property:  VAPOR PRESSURE"
          Printer.Print
          Printer.FontBold = False
          Printer.FontUnderline = True
          Printer.FontSize = 10
          Printer.Print Tab(TABFULLSOURCE); "Source:"; Tab(TABFULLVALUE); "Value:"; Tab(TABFULLUNITS); "Units:"; Tab(TABFULLTEMPERATURE); "Temp.:"; Tab(TABFULLCODE); "Code:"
          Printer.FontUnderline = False
          Printer.Print
          Call FullyPrintVaporPressureToPrinter
          Printer.Print
          Printer.Print

    End Select
End Sub

Private Sub PrintVaporPressureToFile()
   Dim ValueString As String

    Select Case frmPrint!cboPropertyDescription.ListIndex
       Case 0   'Print Selected Value Only
          If HaveProperty(VAPOR_PRESSURE) Then
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = Format$(phprop.VaporPressure.CurrentSelection.Value, GetTheFormat(phprop.VaporPressure.CurrentSelection.Value))
             Print #1, "Vapor Pressure"; Tab(TABVALUE); ValueString; Tab(TABUNITS); Units(VAPOR_PRESSURE); Tab(TABSOURCE); GetSource(phprop.VaporPressure.CurrentSelection.Source)
          Else
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = "Not Available"
             Print #1, "Vapor Pressure"; Tab(TABVALUE); ValueString
          End If

       Case 1   'Print Full Description of Vapor Pressure
          Print #1, "Property:  VAPOR PRESSURE"
          Print #1,
          Print #1, Tab(TABFULLSOURCE); "Source:"; Tab(TABFULLVALUE); "Value:"; Tab(TABFULLUNITS); "Units:"; Tab(TABFULLTEMPERATURE); "Temperature:"; Tab(TABFULLCODE); "Code:"
          Print #1,
          Call FullyPrintVaporPressureToFile
          Print #1,
          Print #1,

    End Select

End Sub

Private Sub PrintWaterDensityPrinter()
    Dim ValueString As String

    Select Case frmPrint!cboPropertyDescription.ListIndex
       Case 0   'Print Selected Value Only
          If HaveProperty(WATER_DENSITY) Then
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = Format$(phprop.WaterDensity.CurrentSelection.Value, WATER_DENSITY_FORMAT)
             Printer.Print "Water Density"; Tab(TABVALUE); ValueString; Tab(TABUNITS); Units(WATER_DENSITY); Tab(TABSOURCE); GetSource(phprop.WaterDensity.CurrentSelection.Source)
          Else
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = "Not Available"
             Printer.Print "Water Density"; Tab(TABVALUE); ValueString
          End If
       Case 1   'Print Full Description of Water Density
          HeightWaterDensity = 0
          Printer.FontSize = 12
          PrintMsg = ""
          HeightWaterDensity = HeightWaterDensity + NUMLINES_PROPERTY_NAME * Printer.TextHeight(PrintMsg)
          Printer.FontSize = 10
          HeightWaterDensity = HeightWaterDensity + NUMLINES_WATER_DENSITY * Printer.TextHeight(PrintMsg)
          TotalHeightThisPage = Printer.CurrentY
          If TotalHeightThisPage + HeightWaterDensity + BOTTOM_MARGIN_SAFETY_FACTOR > Printer.Height Then
             Printer.NewPage
             Call PrintAirWaterTitleContinuation
          End If

          Printer.FontBold = True
          Printer.FontSize = 12
          Printer.Print "Property:  WATER DENSITY"
          Printer.Print
          Printer.FontBold = False
          Printer.FontUnderline = True
          Printer.FontSize = 10
          Printer.Print Tab(TABFULLSOURCE); "Source:"; Tab(TABFULLVALUE); "Value:"; Tab(TABFULLUNITS); "Units:"; Tab(TABFULLTEMPERATURE); "Temp.:"; Tab(TABFULLCODE); "Code:"
          Printer.FontUnderline = False
          Printer.Print
          Call FullyPrintWaterDensityToPrinter
          Printer.Print
          Printer.Print

    End Select

End Sub

Private Sub PrintWaterDensityToFile()
    Dim ValueString As String

    Select Case frmPrint!cboPropertyDescription.ListIndex
       Case 0   'Print Selected Value Only
          If HaveProperty(WATER_DENSITY) Then
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = Format$(phprop.WaterDensity.CurrentSelection.Value, WATER_DENSITY_FORMAT)
             Print #1, "Water Density"; Tab(TABVALUE); ValueString; Tab(TABUNITS); Units(WATER_DENSITY); Tab(TABSOURCE); GetSource(phprop.WaterDensity.CurrentSelection.Source)
          Else
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = "Not Available"
             Print #1, "Water Density"; Tab(TABVALUE); ValueString
          End If
       Case 1   'Print Full Description of Water Density
          Print #1, "Property:  WATER DENSITY"
          Print #1,
          Print #1, Tab(TABFULLSOURCE); "Source:"; Tab(TABFULLVALUE); "Value:"; Tab(TABFULLUNITS); "Units:"; Tab(TABFULLTEMPERATURE); "Temp.:"; Tab(TABFULLCODE); "Code:"
          Print #1,
          Call FullyPrintWaterDensityToFile
          Print #1,
          Print #1,

    End Select

End Sub

Private Sub PrintWaterSurfaceTensionPrinter()
    Dim ValueString As String

    Select Case frmPrint!cboPropertyDescription.ListIndex
       Case 0   'Print Selected Value Only
          If HaveProperty(WATER_SURFACE_TENSION) Then
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = Format$(phprop.WaterSurfaceTension.CurrentSelection.Value, GetTheFormat(phprop.WaterSurfaceTension.CurrentSelection.Value))
             Printer.Print "Water Surface Tension"; Tab(TABVALUE); ValueString; Tab(TABUNITS); Units(WATER_SURFACE_TENSION); Tab(TABSOURCE); GetSource(phprop.WaterSurfaceTension.CurrentSelection.Source)
          Else
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = "Not Available"
             Printer.Print "Water Surface Tension"; Tab(TABVALUE); ValueString
          End If

       Case 1   'Print Full Description of Water Surface Tension
          HeightWaterSurfaceTension = 0
          Printer.FontSize = 12
          PrintMsg = ""
          HeightWaterSurfaceTension = HeightWaterSurfaceTension + NUMLINES_PROPERTY_NAME * Printer.TextHeight(PrintMsg)
          Printer.FontSize = 10
          HeightWaterSurfaceTension = HeightWaterSurfaceTension + NUMLINES_WATER_SURFACE_TENSION * Printer.TextHeight(PrintMsg)
          TotalHeightThisPage = Printer.CurrentY
          If TotalHeightThisPage + HeightWaterSurfaceTension + BOTTOM_MARGIN_SAFETY_FACTOR > Printer.Height Then
             Printer.NewPage
             Call PrintAirWaterTitleContinuation
          End If

          Printer.FontBold = True
          Printer.FontSize = 12
          Printer.Print "Property:  WATER SURFACE TENSION"
          Printer.Print
          Printer.FontBold = False
          Printer.FontUnderline = True
          Printer.FontSize = 10
          Printer.Print Tab(TABFULLSOURCE); "Source:"; Tab(TABFULLVALUE); "Value:"; Tab(TABFULLUNITS); "Units:"; Tab(TABFULLTEMPERATURE); "Temp.:"; Tab(TABFULLCODE); "Code:"
          Printer.FontUnderline = False
          Printer.Print
          Call FullyPrintWaterSurfaceTensionToPrinter
          Printer.Print
          Printer.Print

    End Select

End Sub

Private Sub PrintWaterSurfaceTensionToFile()
    Dim ValueString As String

    Select Case frmPrint!cboPropertyDescription.ListIndex
       Case 0   'Print Selected Value Only
          If HaveProperty(WATER_SURFACE_TENSION) Then
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = Format$(phprop.WaterSurfaceTension.CurrentSelection.Value, GetTheFormat(phprop.WaterSurfaceTension.CurrentSelection.Value))
             Print #1, "Water Surface Tension"; Tab(TABVALUE); ValueString; Tab(TABUNITS); Units(WATER_SURFACE_TENSION); Tab(TABSOURCE); GetSource(phprop.WaterSurfaceTension.CurrentSelection.Source)
          Else
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = "Not Available"
             Print #1, "Water Surface Tension"; Tab(TABVALUE); ValueString
          End If

       Case 1   'Print Full Description of Water Surface Tension
          Print #1, "Property:  WATER SURFACE TENSION"
          Print #1,
          Print #1, Tab(TABFULLSOURCE); "Source:"; Tab(TABFULLVALUE); "Value:"; Tab(TABFULLUNITS); "Units:"; Tab(TABFULLTEMPERATURE); "Temp.:"; Tab(TABFULLCODE); "Code:"
          Print #1,
          Call FullyPrintWaterSurfaceTensionToFile
          Print #1,
          Print #1,

    End Select

End Sub

Private Sub PrintWaterViscosityPrinter()
    Dim ValueString As String

    Select Case frmPrint!cboPropertyDescription.ListIndex
       Case 0   'Print Selected Value Only
          If HaveProperty(WATER_VISCOSITY) Then
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = Format$(phprop.WaterViscosity.CurrentSelection.Value, GetTheFormat(phprop.WaterViscosity.CurrentSelection.Value))
             Printer.Print "Water Viscosity"; Tab(TABVALUE); ValueString; Tab(TABUNITS); Units(WATER_VISCOSITY); Tab(TABSOURCE); GetSource(phprop.WaterViscosity.CurrentSelection.Source)
          Else
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = "Not Available"
             Printer.Print "Water Viscosity"; Tab(TABVALUE); ValueString
          End If
       Case 1   'Print Full Description of Water Viscosity
          HeightWaterViscosity = 0
          Printer.FontSize = 12
          PrintMsg = ""
          HeightWaterViscosity = HeightWaterViscosity + NUMLINES_PROPERTY_NAME * Printer.TextHeight(PrintMsg)
          Printer.FontSize = 10
          HeightWaterViscosity = HeightWaterViscosity + NUMLINES_WATER_VISCOSITY * Printer.TextHeight(PrintMsg)
          TotalHeightThisPage = Printer.CurrentY
          If TotalHeightThisPage + HeightWaterViscosity + BOTTOM_MARGIN_SAFETY_FACTOR > Printer.Height Then
             Printer.NewPage
             Call PrintAirWaterTitleContinuation
          End If

          Printer.FontBold = True
          Printer.FontSize = 12
          Printer.Print "Property:  WATER VISCOSITY"
          Printer.Print
          Printer.FontBold = False
          Printer.FontUnderline = True
          Printer.FontSize = 10
          Printer.Print Tab(TABFULLSOURCE); "Source:"; Tab(TABFULLVALUE); "Value:"; Tab(TABFULLUNITS); "Units:"; Tab(TABFULLTEMPERATURE); "Temp.:"; Tab(TABFULLCODE); "Code:"
          Printer.FontUnderline = False
          Printer.Print
          Call FullyPrintWaterViscosityToPrinter
          Printer.Print
          Printer.Print

    End Select

End Sub

Private Sub PrintWaterViscosityToFile()
    Dim ValueString As String

    Select Case frmPrint!cboPropertyDescription.ListIndex
       Case 0   'Print Selected Value Only
          If HaveProperty(WATER_VISCOSITY) Then
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = Format$(phprop.WaterViscosity.CurrentSelection.Value, GetTheFormat(phprop.WaterViscosity.CurrentSelection.Value))
             Print #1, "Water Viscosity"; Tab(TABVALUE); ValueString; Tab(TABUNITS); Units(WATER_VISCOSITY); Tab(TABSOURCE); GetSource(phprop.WaterViscosity.CurrentSelection.Source)
          Else
             ValueString = Space$(VALUELENGTH)
             RSet ValueString = "Not Available"
             Print #1, "Water Viscosity"; Tab(TABVALUE); ValueString
          End If
       Case 1   'Print Full Description of Water Viscosity
          Print #1, "Property:  WATER VISCOSITY"
          Print #1,
          Print #1, Tab(TABFULLSOURCE); "Source:"; Tab(TABFULLVALUE); "Value:"; Tab(TABFULLUNITS); "Units:"; Tab(TABFULLTEMPERATURE); "Temp.:"; Tab(TABFULLCODE); "Code:"
          Print #1,
          Call FullyPrintWaterViscosityToFile
          Print #1,
          Print #1,

    End Select

End Sub


