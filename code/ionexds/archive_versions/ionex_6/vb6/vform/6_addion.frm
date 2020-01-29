VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{B16553C3-06DB-101B-85B2-0000C009BE81}#1.0#0"; "Spin32.ocx"
Begin VB.Form frmAddComponent 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Component"
   ClientHeight    =   6480
   ClientLeft      =   1740
   ClientTop       =   390
   ClientWidth     =   4305
   ControlBox      =   0   'False
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6480
   ScaleWidth      =   4305
   Begin Threed.SSFrame fraIonProperties 
      Height          =   1455
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   4095
      _Version        =   65536
      _ExtentX        =   7223
      _ExtentY        =   2566
      _StockProps     =   14
      Caption         =   "Properties of the Ion:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ComboBox cboAddIonUnits 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   2820
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   720
         Width           =   1152
      End
      Begin VB.TextBox txtAddIon 
         Appearance      =   0  'Flat
         Height          =   288
         Index           =   2
         Left            =   1680
         TabIndex        =   9
         Top             =   720
         Width           =   1032
      End
      Begin VB.ComboBox cboAddIonUnits 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   2820
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   360
         Width           =   1152
      End
      Begin VB.TextBox txtAddIon 
         Appearance      =   0  'Flat
         Height          =   288
         Index           =   1
         Left            =   1680
         TabIndex        =   7
         Top             =   360
         Width           =   1032
      End
      Begin Spin.SpinButton SpinButton1 
         Height          =   255
         Left            =   2520
         TabIndex        =   15
         Top             =   1080
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   450
         _StockProps     =   73
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
      End
      Begin VB.Label lblValence 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1800
         TabIndex        =   37
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblValenceSign 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "+"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1680
         TabIndex        =   14
         Top             =   1080
         Width           =   135
      End
      Begin VB.Label lblAddIon 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Valence"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   0
         TabIndex        =   13
         Top             =   1140
         Width           =   1575
      End
      Begin VB.Label lblAddIon 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Initial Conc."
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   0
         TabIndex        =   12
         Top             =   780
         Width           =   1575
      End
      Begin VB.Label lblAddIon 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Molecular Weight"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   0
         TabIndex        =   11
         Top             =   420
         Width           =   1575
      End
   End
   Begin Threed.SSFrame fraSeparationFactor 
      Height          =   1455
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   4095
      _Version        =   65536
      _ExtentX        =   7223
      _ExtentY        =   2566
      _StockProps     =   14
      Caption         =   "Separation Factor:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtAlphaValue 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2460
         TabIndex        =   17
         Top             =   660
         Width           =   1515
      End
      Begin VB.CommandButton cmdViewSeparationFactors 
         Appearance      =   0  'Flat
         Caption         =   "View All Separation Factors"
         Height          =   312
         Left            =   360
         TabIndex        =   16
         Top             =   1080
         Width           =   3552
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Value:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2460
         TabIndex        =   21
         Top             =   420
         Width           =   1515
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ALPHA"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   300
         TabIndex        =   20
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblAlpha 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   19
         Top             =   420
         Width           =   1155
      End
      Begin VB.Label lblAlpha 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   18
         Top             =   720
         Width           =   1155
      End
      Begin VB.Shape Shape5 
         Height          =   675
         Left            =   240
         Top             =   360
         Width           =   2115
      End
   End
   Begin Threed.SSFrame fraNernstHaskell 
      Height          =   2175
      Left            =   120
      TabIndex        =   4
      Top             =   3720
      Width           =   4095
      _Version        =   65536
      _ExtentX        =   7223
      _ExtentY        =   3836
      _StockProps     =   14
      Caption         =   "Parameters for Nernst-Haskell Correlation:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmdAddIon 
         Appearance      =   0  'Flat
         Caption         =   "Add Cation"
         Height          =   252
         Index           =   1
         Left            =   2280
         TabIndex        =   25
         Top             =   1440
         Width           =   1452
      End
      Begin VB.CommandButton cmdAddIon 
         Appearance      =   0  'Flat
         Caption         =   "Add Anion"
         Height          =   252
         Index           =   0
         Left            =   360
         TabIndex        =   24
         Top             =   1440
         Width           =   1452
      End
      Begin VB.ComboBox cboCation 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2220
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   540
         Width           =   1692
      End
      Begin VB.ComboBox cboAnion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   300
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   540
         Width           =   1632
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "L.I.C. = Limiting Ionic Conductance"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   36
         Top             =   1870
         Width           =   3675
      End
      Begin VB.Label lblAddIon 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "L.I.C."
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   35
         Top             =   1140
         Width           =   795
      End
      Begin VB.Label lblAddIon 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Valence"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   34
         Top             =   900
         Width           =   795
      End
      Begin VB.Shape Shape1 
         Height          =   1575
         Left            =   120
         Top             =   240
         Width           =   1935
      End
      Begin VB.Shape Shape3 
         Height          =   315
         Left            =   120
         Top             =   1800
         Width           =   3855
      End
      Begin VB.Label lblAddIonValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   3060
         TabIndex        =   33
         Top             =   1140
         Width           =   795
      End
      Begin VB.Label lblAddIonValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   3060
         TabIndex        =   32
         Top             =   900
         Width           =   795
      End
      Begin VB.Label lblAddIon 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "L.I.C."
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   2160
         TabIndex        =   31
         Top             =   1140
         Width           =   795
      End
      Begin VB.Label lblAddIon 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Valence"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   2160
         TabIndex        =   30
         Top             =   900
         Width           =   795
      End
      Begin VB.Label lblAddIonValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   1080
         TabIndex        =   29
         Top             =   1140
         Width           =   795
      End
      Begin VB.Label lblAddIonValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   1080
         TabIndex        =   28
         Top             =   900
         Width           =   795
      End
      Begin VB.Shape Shape2 
         Height          =   1575
         Left            =   2040
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label lblKineticParameters 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cation:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   2100
         TabIndex        =   27
         Top             =   300
         Width           =   615
      End
      Begin VB.Label lblKineticParameters 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Anion:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   26
         Top             =   300
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      Caption         =   "OK"
      Height          =   372
      Left            =   1140
      TabIndex        =   1
      Top             =   6000
      Width           =   975
   End
   Begin VB.TextBox txtAddIon 
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   180
      Width           =   2832
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Caption         =   "Cancel"
      Height          =   372
      Left            =   2220
      TabIndex        =   2
      Top             =   6000
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Ion:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   480
      LinkItem        =   "Name"
      TabIndex        =   3
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "frmAddComponent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim Temp_Text As String
Dim IsError As Integer

Private Sub cboAddIonUnits_Click(Index As Integer)
    Dim ValueToDisplay As Double

    Select Case Index
       Case 1   'Initial Concentration
            Select Case cboAddIonUnits(1).ListIndex
               Case CONCENTRATION_MG_per_L    'mg/L
                    ValueToDisplay = ChangedIon.InitialConcentration
               Case CONCENTRATION_UG_per_L   'ug/L
                    ValueToDisplay = ChangedIon.InitialConcentration * ConcentrationConversionFactor(CONCENTRATION_UG_per_L, ChangedIon.Valence, ChangedIon.MolecularWeight)
               Case CONCENTRATION_G_per_L    'g/L
                    ValueToDisplay = ChangedIon.InitialConcentration * ConcentrationConversionFactor(CONCENTRATION_G_per_L, ChangedIon.Valence, ChangedIon.MolecularWeight)
               Case CONCENTRATION_MEQ_per_L   'meq/L
                    ValueToDisplay = ChangedIon.InitialConcentration * ConcentrationConversionFactor(CONCENTRATION_MEQ_per_L, ChangedIon.Valence, ChangedIon.MolecularWeight)
               Case CONCENTRATION_EQ_per_L   'eq/L
                    ValueToDisplay = ChangedIon.InitialConcentration * ConcentrationConversionFactor(CONCENTRATION_EQ_per_L, ChangedIon.Valence, ChangedIon.MolecularWeight)
               Case CONCENTRATION_MMOL_per_L   'mmol/L
                    ValueToDisplay = ChangedIon.InitialConcentration * ConcentrationConversionFactor(CONCENTRATION_MMOL_per_L, ChangedIon.Valence, ChangedIon.MolecularWeight)
               Case CONCENTRATION_UMOL_per_L   'umol/L
                    ValueToDisplay = ChangedIon.InitialConcentration * ConcentrationConversionFactor(CONCENTRATION_UMOL_per_L, ChangedIon.Valence, ChangedIon.MolecularWeight)
               Case CONCENTRATION_GMOL_per_L   'gmol/L
                    ValueToDisplay = ChangedIon.InitialConcentration * ConcentrationConversionFactor(CONCENTRATION_GMOL_per_L, ChangedIon.Valence, ChangedIon.MolecularWeight)

            End Select
            txtAddIon(2).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

    End Select
End Sub

Private Sub cboAnion_Click()

    'Anion
    ChangedIon.Kinetic.NernstHaskellAnion = NernstHaskell.Anion(cboAnion.ListIndex + 1)
    lblAddIonValue(0).Caption = "-" & Format$(ChangedIon.Kinetic.NernstHaskellAnion.Valence, "0")
    lblAddIonValue(1).Caption = Trim$(Str$(ChangedIon.Kinetic.NernstHaskellAnion.LimitingIonicConductance))

End Sub

Private Sub cboCation_Click()

    'Cation
    ChangedIon.Kinetic.NernstHaskellCation = NernstHaskell.Cation(cboCation.ListIndex + 1)
    lblAddIonValue(2).Caption = "+" & Format$(ChangedIon.Kinetic.NernstHaskellCation.Valence, "0")
    lblAddIonValue(3).Caption = Trim$(Str$(ChangedIon.Kinetic.NernstHaskellCation.LimitingIonicConductance))
    
End Sub

Private Sub cmdCancel_Click()

    frmAddComponent.Hide

End Sub

Private Sub cmdOK_Click()
    Dim msg As String
    Dim i As Integer, j As Integer
    Dim OldIonToEdit As Integer

    If frmAddComponent.Caption = "Add Cation" Then
       If Trim$(txtAddIon(0).Text) = "Cation" Then
          msg = "You have not selected a name for this new ion.  "
          msg = msg & "'Cation' is not a legal ion name.  Please "
          msg = msg & "specify a name for this new ion."
          MsgBox msg, MB_ICONSTOP, "Illegal Ion Name"
          Exit Sub
       End If

       NowProj.NumberOfCations = NowProj.NumberOfCations + 1
       NumberOfIons = NowProj.NumberOfCations
       NowProj.Cation(NowProj.NumberOfCations) = ChangedIon
'       Call CalculateSumEquivInitialConc
       frmIonExchangeMain!cboIons(0).AddItem NowProj.Cation(NowProj.NumberOfCations).Name
       frmIonExchangeMain!cboIons(2).AddItem NowProj.Cation(NowProj.NumberOfCations).Name
'       frmIonExchangeMain!cboKinDimComponent.AddItem nowproj.cation( nowproj.NumberOfCations).Name
       frmInputKineticParameters!cboIon.AddItem NowProj.Cation(NowProj.NumberOfCations).Name
       If NowProj.NumberOfCations > 1 Then
          frmIonExchangeMain!lstIons(0).AddItem NowProj.Cation(NowProj.NumberOfCations).Name
       Else
          frmIonExchangeMain!cboIons(0).ListIndex = 0
       End If

       Call CalculateKineticParameters

'       OldIonToEdit = NumberOfIonToEdit
'       For I = 1 To  nowproj.NumberOfCations
'           NumberOfIonToEdit = I
'           Call CalculateDimensionlessGroups
'       Next I
'       NumberOfIonToEdit = OldIonToEdit

       For i = 1 To NowProj.NumberOfCations
           OneDimSeparationFactors(i) = NowProj.Cation(i).SeparationFactor
       Next i
       Call CalculateSeparationFactors
       
       If frmIonExchangeMain!cboIons(2).ListCount = 1 Then
          frmIonExchangeMain!cboIons(2).ListIndex = 0
'          frmIonExchangeMain!cboKinDimComponent.ListIndex = 0
          frmInputKineticParameters!cboIon.ListIndex = 0
          frmIonExchangeMain!fraKineticDimensionless.Enabled = True
       End If
       
    ElseIf frmAddComponent.Caption = "Add Anion" Then
       If Trim$(txtAddIon(0).Text) = "Anion" Then
          msg = "You have not selected a name for this new ion.  "
          msg = msg & "'Anion' is not a legal ion name.  Please "
          msg = msg & "specify a name for this new ion."
          MsgBox msg, MB_ICONSTOP, "Illegal Ion Name"
          Exit Sub
       End If

       NowProj.NumberOfAnions = NowProj.NumberOfAnions + 1
       NumberOfIons = NowProj.NumberOfAnions
       NowProj.Anion(NowProj.NumberOfAnions) = ChangedIon
'       Call CalculateSumEquivInitialConc
       frmIonExchangeMain!cboIons(1).AddItem NowProj.Anion(NowProj.NumberOfAnions).Name
       frmIonExchangeMain!cboIons(2).AddItem NowProj.Anion(NowProj.NumberOfAnions).Name
'       frmIonExchangeMain!cboKinDimComponent.AddItem nowproj.anion(nowproj.NumberOfAnions).Name
       frmInputKineticParameters!cboIon.AddItem NowProj.Anion(NowProj.NumberOfAnions).Name

       If NowProj.NumberOfAnions > 1 Then
          frmIonExchangeMain!lstIons(1).AddItem NowProj.Anion(NowProj.NumberOfAnions).Name
       Else
          frmIonExchangeMain!cboIons(1).ListIndex = 0
       End If
       
       Call CalculateKineticParameters
       
'       OldIonToEdit = NumberOfIonToEdit
'       For I = 1 To nowproj.NumberOfAnions
'           NumberOfIonToEdit = I
'           Call CalculateDimensionlessGroups
'       Next I
'       NumberOfIonToEdit = OldIonToEdit

       For i = 1 To NowProj.NumberOfAnions
           OneDimSeparationFactors(i) = NowProj.Anion(i).SeparationFactor
       Next i
       Call CalculateSeparationFactors
  

       If frmIonExchangeMain!cboIons(2).ListCount = 1 Then
          frmIonExchangeMain!cboIons(2).ListIndex = 0
'          frmIonExchangeMain!cboKinDimComponent.ListIndex = 0
          frmInputKineticParameters!cboIon.ListIndex = 0
          frmIonExchangeMain!fraKineticDimensionless.Enabled = True
       End If

    ElseIf frmAddComponent.Caption = "Edit Cation" Then
       NowProj.Cation(NumberOfIonToEdit) = ChangedIon
       Call CalculateKineticParameters
       
       For j = 1 To NumSelectedCations
           If Cations_Selected(j) = NumberOfIonToEdit Then
              Call CalculateSumEquivInitialConc
      
              OldIonToEdit = NumberOfIonToEdit
              For i = 1 To NumSelectedCations
                  NumberOfIonToEdit = Cations_Selected(i)
                  Call CalculateDimensionlessGroups
              Next i
              NumberOfIonToEdit = OldIonToEdit
           End If
       Next j

       For i = 1 To NowProj.NumberOfCations
           OneDimSeparationFactors(i) = NowProj.Cation(i).SeparationFactor
       Next i
       Call CalculateSeparationFactors
     
    ElseIf frmAddComponent.Caption = "Edit Anion" Then
       NowProj.Anion(NumberOfIonToEdit) = ChangedIon
       Call CalculateKineticParameters

       For j = 1 To NumSelectedAnions
           If Anions_Selected(j) = NumberOfIonToEdit Then
              Call CalculateSumEquivInitialConc
      
              OldIonToEdit = NumberOfIonToEdit
              For i = 1 To NumSelectedAnions
                  NumberOfIonToEdit = Anions_Selected(i)
                  Call CalculateDimensionlessGroups
              Next i
              NumberOfIonToEdit = OldIonToEdit
           End If
       Next j

       For i = 1 To NowProj.NumberOfAnions
           OneDimSeparationFactors(i) = NowProj.Anion(i).SeparationFactor
       Next i
       Call CalculateSeparationFactors

    End If

    frmAddComponent.Hide

End Sub

Private Sub cmdViewSeparationFactors_Click()
    Dim i As Integer

    For i = 1 To MAX_CHEMICAL
        OldOneDimSeparationFactors(i) = OneDimSeparationFactors(i)
    Next i
    OldOptionButtonSeparationFactors = SeparationFactorInput.Value

    frmSeparationFactors.Show 1

End Sub

Private Sub Form_Load()
    Dim PositionLeft As Integer

    frmAddComponent.WindowState = 0
    frmAddComponent.width = 4400
    frmAddComponent.height = 6800

    'Position the form on the screen (Centered in Right Half of It)
    If WindowState = 0 Then
       'don't attempt if screen Minimized or Maximized
       PositionLeft = frmIonExchangeMain.left + frmIonExchangeMain.width - Screen.width / 2
       PositionLeft = (PositionLeft / 2 - frmAddComponent.width / 2)
       Move (Screen.width / 2 + PositionLeft), (Screen.height - frmAddComponent.height) / 2
    End If

    Call Populate_frmAddComponent_Units
     
End Sub

Private Sub spnValence_SpinDown()

    If CInt(lblValence.Caption) = 1 Then
       lblValence.Caption = 10
    Else
       lblValence.Caption = Str$(CInt(lblValence.Caption) - 1)
    End If
    ChangedIon.Valence = CDbl(lblValence.Caption)
    Call UpdateInitialConcentrationValence
    Call CalculateEquivalentInitialConc(ChangedIon.EquivalentInitialConcentration, ChangedIon.InitialConcentration, ChangedIon.Valence, ChangedIon.MolecularWeight)

End Sub

Private Sub spnValence_SpinUp()

    If CInt(lblValence.Caption) = 10 Then
       lblValence.Caption = 1
    Else
       lblValence.Caption = Str$(CInt(lblValence.Caption) + 1)
    End If
    ChangedIon.Valence = CDbl(lblValence.Caption)
    Call UpdateInitialConcentrationValence
    Call CalculateEquivalentInitialConc(ChangedIon.EquivalentInitialConcentration, ChangedIon.InitialConcentration, ChangedIon.Valence, ChangedIon.MolecularWeight)

End Sub



Private Sub txtAddIon_GotFocus(Index As Integer)
    Call TextGetFocus(txtAddIon(Index), Temp_Text)
End Sub

Private Sub txtAddIon_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
       KeyAscii = 0
       SendKeys "{Tab}"
       Exit Sub
    End If

    If Index <> 0 Then
       Call NumberCheck(KeyAscii)
    End If

End Sub

Private Sub txtAddIon_LostFocus(Index As Integer)
    Dim ValueChanged As Integer, NewValue As Double
    Dim msg As String, CurrentUnits As Integer
    Dim OldValue As Double, ValueToDisplay As Double
    Dim i As Integer, InitialConcentrationUnits As Integer
    Dim Input_Ion_Name As String

    If Index = 0 Then   'Inputting Name of Ion
       Input_Ion_Name = Trim$(txtAddIon(0).Text)
       If Input_Ion_Name = Trim$(Temp_Text) Then Exit Sub
       If Trim$(frmAddComponent.Caption) = "Add Cation" Then
          For i = 1 To NowProj.NumberOfCations
              If Trim$(NowProj.Cation(i).Name) = Input_Ion_Name Then
                 msg = "The name that has been input duplicates a "
                 msg = msg & "cation for which properties are already "
                 msg = msg & "specified.  If you wish to edit properties "
                 msg = msg & "of an existing ion, you must choose the "
                 msg = msg & "'Edit Properties' option rather than the "
                 msg = msg & "'Add Cation' option."
                 MsgBox msg, MB_ICONSTOP, "Duplicate Cation Name"
                 txtAddIon(0).Text = Temp_Text
                 txtAddIon(0).SetFocus
                 Exit Sub
              End If
          Next i
          For i = 1 To NowProj.NumberOfAnions
              If Trim$(NowProj.Anion(i).Name) = Input_Ion_Name Then
                 msg = "The name that has been input duplicates an "
                 msg = msg & "anion for which properties are already "
                 msg = msg & "specified.  It is not allowable to have duplicate names "
                 msg = msg & "for anions and cations.  If you wish to edit properties "
                 msg = msg & "of an existing ion, you must choose the "
                 msg = msg & "'Edit Properties' option rather than the "
                 msg = msg & "'Add Cation' option."
                 MsgBox msg, MB_ICONSTOP, "Duplicate Name"
                 txtAddIon(0).Text = Temp_Text
                 txtAddIon(0).SetFocus
                 Exit Sub
              End If
          Next i
          If SeparationFactorInput.Row = True Then
             lblAlpha(1).Caption = Trim$(txtAddIon(0).Text)
             If NowProj.NumberOfCations = 0 Then lblAlpha(2).Caption = lblAlpha(1).Caption
          Else
             lblAlpha(2).Caption = Trim$(txtAddIon(0).Text)
             If NowProj.NumberOfCations = 0 Then lblAlpha(1).Caption = lblAlpha(2).Caption
          End If

       End If

       If Trim$(frmAddComponent.Caption) = "Add Anion" Then
          For i = 1 To NowProj.NumberOfAnions
              If Trim$(NowProj.Anion(i).Name) = Input_Ion_Name Then
                 msg = "The name that has been input duplicates an "
                 msg = msg & "anion for which properties are already "
                 msg = msg & "specified.  If you wish to edit properties "
                 msg = msg & "of an existing ion, you must choose the "
                 msg = msg & "'Edit Properties' option rather than the "
                 msg = msg & "'Add Anion' option."
                 MsgBox msg, MB_ICONSTOP, "Duplicate Anion Name"
                 txtAddIon(0).Text = Temp_Text
                 txtAddIon(0).SetFocus
                 Exit Sub
              End If
          Next i
          For i = 1 To NowProj.NumberOfCations
              If Trim$(NowProj.Cation(i).Name) = Input_Ion_Name Then
                 msg = "The name that has been input duplicates a "
                 msg = msg & "cation for which properties are already "
                 msg = msg & "specified.  It is not allowable to have duplicate names "
                 msg = msg & "for anions and cations.  If you wish to edit properties "
                 msg = msg & "of an existing ion, you must choose the "
                 msg = msg & "'Edit Properties' option rather than the "
                 msg = msg & "'Add Anion' option."
                 MsgBox msg, MB_ICONSTOP, "Duplicate Name"
                 txtAddIon(0).Text = Temp_Text
                 txtAddIon(0).SetFocus
                 Exit Sub
              End If
          Next i
          If SeparationFactorInput.Row = True Then
             lblAlpha(1).Caption = Trim$(txtAddIon(0).Text)
             If NowProj.NumberOfAnions = 0 Then lblAlpha(2).Caption = lblAlpha(1).Caption
          Else
             lblAlpha(2).Caption = Trim$(txtAddIon(0).Text)
             If NowProj.NumberOfAnions = 0 Then lblAlpha(1).Caption = lblAlpha(2).Caption
          End If
       End If

       ChangedIon.Name = Trim$(txtAddIon(0).Text)

       If AddingCation Then
          NowProj.Cation(NumberOfIons).Name = ChangedIon.Name
       ElseIf AddingAnion Then
          NowProj.Anion(NumberOfIons).Name = ChangedIon.Name
       End If

       Exit Sub
    End If

    Call TextHandleError(IsError, txtAddIon(Index), Temp_Text)

    If Not IsError Then
       NewValue = CDbl(txtAddIon(Index).Text)
       'Convert NewValue to Standard Units if Necessary
       Select Case Index
          Case 1   'Molecular Weight
               OldValue = ChangedIon.MolecularWeight
          Case 2   'Initial Concentration
               OldValue = ChangedIon.InitialConcentration
               CurrentUnits = cboAddIonUnits(1).ListIndex
               If CurrentUnits <> 0 Then
                  NewValue = NewValue / ConcentrationConversionFactor(CurrentUnits, ChangedIon.Valence, ChangedIon.MolecularWeight)
               End If
       End Select

       Select Case Index

          Case 1    'Molecular Weight
             Call NumberChanged(ValueChanged, OldValue, NewValue)
             If ValueChanged Then
                If HaveValue(NewValue) Then
                   ChangedIon.MolecularWeight = NewValue
                   Call UpdateInitialConcentrationMolWt
                   Call CalculateEquivalentInitialConc(ChangedIon.EquivalentInitialConcentration, ChangedIon.InitialConcentration, ChangedIon.Valence, ChangedIon.MolecularWeight)
                Else
                   txtAddIon(1).Text = Temp_Text
                   txtAddIon(1).SetFocus
                   Exit Sub
                End If
             End If

          Case 2    'Initial Concentration
             Call NumberChanged(ValueChanged, OldValue, NewValue)
             If ValueChanged Then
                If NewValue >= 0 Then
                   ChangedIon.InitialConcentration = NewValue
                   Call CalculateEquivalentInitialConc(ChangedIon.EquivalentInitialConcentration, ChangedIon.InitialConcentration, ChangedIon.Valence, ChangedIon.MolecularWeight)
                   
                Else
                   txtAddIon(2).Text = Temp_Text
                   txtAddIon(2).SetFocus
                   Exit Sub
                End If
             End If
             
       End Select

    End If

End Sub

Private Sub txtAlphaValue_GotFocus()
    Call TextGetFocus(txtAlphaValue, Temp_Text)
End Sub

Private Sub txtAlphaValue_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
       KeyAscii = 0
       SendKeys "{Tab}"
       Exit Sub
    End If

    Call NumberCheck(KeyAscii)

End Sub

Private Sub txtAlphaValue_LostFocus()
    Dim ValueChanged As Integer, NewValue As Double
    Dim msg As String, CurrentUnits As Integer
    Dim OldValue As Double, ValueToDisplay As Double

    Call TextHandleError(IsError, txtAlphaValue, Temp_Text)

    If Not IsError Then
       NewValue = CDbl(txtAlphaValue.Text)
       OldValue = ChangedIon.SeparationFactor

       Call NumberChanged(ValueChanged, OldValue, NewValue)
       If ValueChanged Then
          If HaveValue(NewValue) Then
             ChangedIon.SeparationFactor = NewValue
             OneDimSeparationFactors(NumberOfIonToEdit) = NewValue
          Else
             txtAlphaValue.Text = Temp_Text
             txtAlphaValue.SetFocus
             Exit Sub
          End If
       End If
       
    End If
End Sub

Private Sub UpdateInitialConcentrationMolWt()
    Dim ValueToDisplay As Double
    Dim InitialConcentrationUnits As Integer

    InitialConcentrationUnits = cboAddIonUnits(1).ListIndex
    Select Case InitialConcentrationUnits
       Case CONCENTRATION_MEQ_per_L, CONCENTRATION_EQ_per_L, CONCENTRATION_MMOL_per_L, CONCENTRATION_UMOL_per_L, CONCENTRATION_GMOL_per_L
          'Recalculate Concentration displayed as it depends on molecular weight
          ValueToDisplay = ChangedIon.InitialConcentration * ConcentrationConversionFactor(InitialConcentrationUnits, ChangedIon.Valence, ChangedIon.MolecularWeight)
          txtAddIon(2).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
    End Select

End Sub

Private Sub UpdateInitialConcentrationValence()
    Dim ValueToDisplay As Double
    Dim InitialConcentrationUnits As Integer

    InitialConcentrationUnits = cboAddIonUnits(1).ListIndex
    Select Case InitialConcentrationUnits
       Case CONCENTRATION_MEQ_per_L, CONCENTRATION_EQ_per_L
          'Recalculate Concentration displayed as it depends on valence
          ValueToDisplay = ChangedIon.InitialConcentration * ConcentrationConversionFactor(InitialConcentrationUnits, ChangedIon.Valence, ChangedIon.MolecularWeight)
          txtAddIon(2).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
    End Select

End Sub

