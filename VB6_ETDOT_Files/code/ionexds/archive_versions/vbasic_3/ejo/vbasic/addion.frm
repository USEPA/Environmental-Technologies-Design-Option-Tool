VERSION 2.00
Begin Form frmAddComponent 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Component"
   ClientHeight    =   6480
   ClientLeft      =   1740
   ClientTop       =   390
   ClientWidth     =   4305
   ControlBox      =   0   'False
   Height          =   6885
   Left            =   1680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   4305
   Top             =   45
   Width           =   4425
   Begin SSFrame fraSeparationFactor 
      Caption         =   "Separation Factor"
      ForeColor       =   &H00000000&
      Height          =   1392
      Left            =   120
      ShadowColor     =   1  'Black
      TabIndex        =   25
      Top             =   2100
      Width           =   4092
      Begin CommandButton cmdViewSeparationFactors 
         Caption         =   "View All Separation Factors"
         Height          =   312
         Left            =   240
         TabIndex        =   36
         Top             =   1020
         Width           =   3552
      End
      Begin TextBox txtAlphaValue 
         Height          =   285
         Left            =   2340
         TabIndex        =   6
         Top             =   600
         Width           =   1515
      End
      Begin Shape Shape6 
         Height          =   672
         Left            =   2220
         Top             =   300
         Width           =   1692
      End
      Begin Shape Shape5 
         Height          =   672
         Left            =   120
         Top             =   300
         Width           =   2112
      End
      Begin Label lblAlpha 
         BorderStyle     =   1  'Fixed Single
         Height          =   252
         Index           =   2
         Left            =   960
         TabIndex        =   35
         Top             =   660
         Width           =   1152
      End
      Begin Label lblAlpha 
         BorderStyle     =   1  'Fixed Single
         Height          =   252
         Index           =   1
         Left            =   960
         TabIndex        =   34
         Top             =   360
         Width           =   1152
      End
      Begin Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "ALPHA"
         Height          =   192
         Left            =   180
         TabIndex        =   33
         Top             =   540
         Width           =   732
      End
      Begin Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Value:"
         Height          =   192
         Left            =   2340
         TabIndex        =   32
         Top             =   360
         Width           =   1512
      End
   End
   Begin SSFrame fraIonProperties 
      Caption         =   "Properties of the Ion:"
      ForeColor       =   &H00000000&
      Height          =   1452
      Left            =   120
      ShadowColor     =   1  'Black
      TabIndex        =   31
      Top             =   600
      Width           =   4092
      Begin TextBox txtAddIon 
         Height          =   288
         Index           =   1
         Left            =   1740
         TabIndex        =   1
         Top             =   360
         Width           =   1032
      End
      Begin ComboBox cboAddIonUnits 
         Height          =   288
         Index           =   0
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   1152
      End
      Begin TextBox txtAddIon 
         Height          =   288
         Index           =   2
         Left            =   1740
         TabIndex        =   3
         Top             =   720
         Width           =   1032
      End
      Begin ComboBox cboAddIonUnits 
         Height          =   288
         Index           =   1
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   720
         Width           =   1152
      End
      Begin SpinButton spnValence 
         Height          =   252
         Left            =   2520
         Top             =   1080
         Width           =   252
      End
      Begin Label lblAddIon 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Molecular Weight"
         Height          =   192
         Index           =   4
         Left            =   60
         TabIndex        =   26
         Top             =   420
         Width           =   1572
      End
      Begin Label lblAddIon 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Initial Conc."
         Height          =   192
         Index           =   5
         Left            =   60
         TabIndex        =   27
         Top             =   780
         Width           =   1572
      End
      Begin Label lblAddIon 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Valence"
         Height          =   192
         Index           =   6
         Left            =   60
         TabIndex        =   28
         Top             =   1140
         Width           =   1572
      End
      Begin Label lblValenceSign 
         Alignment       =   2  'Center
         Caption         =   "+"
         Height          =   252
         Left            =   1740
         TabIndex        =   29
         Top             =   1080
         Width           =   132
      End
      Begin Label lblValence 
         Height          =   252
         Left            =   1860
         TabIndex        =   5
         Top             =   1080
         Width           =   672
      End
      Begin Shape Shape4 
         Height          =   252
         Left            =   1740
         Top             =   1080
         Width           =   792
      End
   End
   Begin CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   372
      Left            =   1380
      TabIndex        =   11
      Top             =   6000
      Width           =   732
   End
   Begin TextBox txtAddIon 
      Height          =   288
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Top             =   180
      Width           =   2832
   End
   Begin SSFrame fraNernstHaskell 
      Caption         =   "Parameters for Nernst-Haskell Correlation:"
      ForeColor       =   &H00000000&
      Height          =   2352
      Left            =   120
      TabIndex        =   13
      Top             =   3540
      Width           =   4092
      Begin ComboBox cboAnion 
         Height          =   288
         Left            =   300
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   660
         Width           =   1632
      End
      Begin ComboBox cboCation 
         Height          =   288
         Left            =   2220
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   660
         Width           =   1692
      End
      Begin CommandButton cmdAddIon 
         Caption         =   "Add Anion"
         Height          =   252
         Index           =   0
         Left            =   360
         TabIndex        =   8
         Top             =   1560
         Width           =   1452
      End
      Begin CommandButton cmdAddIon 
         Caption         =   "Add Cation"
         Height          =   252
         Index           =   1
         Left            =   2280
         TabIndex        =   10
         Top             =   1560
         Width           =   1452
      End
      Begin Label lblKineticParameters 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Anion:"
         Height          =   192
         Index           =   0
         Left            =   180
         TabIndex        =   14
         Top             =   420
         Width           =   612
      End
      Begin Label lblKineticParameters 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cation:"
         Height          =   192
         Index           =   1
         Left            =   2100
         TabIndex        =   15
         Top             =   420
         Width           =   612
      End
      Begin Shape Shape2 
         Height          =   1572
         Left            =   2040
         Top             =   360
         Width           =   1932
      End
      Begin Label lblAddIonValue 
         BorderStyle     =   1  'Fixed Single
         Height          =   192
         Index           =   0
         Left            =   1080
         TabIndex        =   16
         Top             =   1020
         Width           =   792
      End
      Begin Label lblAddIonValue 
         BorderStyle     =   1  'Fixed Single
         Height          =   192
         Index           =   1
         Left            =   1080
         TabIndex        =   17
         Top             =   1260
         Width           =   792
      End
      Begin Label lblAddIon 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Valence"
         Height          =   192
         Index           =   2
         Left            =   2160
         TabIndex        =   18
         Top             =   1020
         Width           =   792
      End
      Begin Label lblAddIon 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "L.I.C."
         Height          =   192
         Index           =   3
         Left            =   2160
         TabIndex        =   19
         Top             =   1260
         Width           =   792
      End
      Begin Label lblAddIonValue 
         BorderStyle     =   1  'Fixed Single
         Height          =   192
         Index           =   2
         Left            =   3060
         TabIndex        =   20
         Top             =   1020
         Width           =   792
      End
      Begin Label lblAddIonValue 
         BorderStyle     =   1  'Fixed Single
         Height          =   192
         Index           =   3
         Left            =   3060
         TabIndex        =   21
         Top             =   1260
         Width           =   792
      End
      Begin Shape Shape3 
         Height          =   312
         Left            =   120
         Top             =   1920
         Width           =   3852
      End
      Begin Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "L.I.C. = Limiting Ionic Conductance"
         Height          =   192
         Left            =   240
         TabIndex        =   22
         Top             =   1980
         Width           =   3672
      End
      Begin Shape Shape1 
         Height          =   1572
         Left            =   120
         Top             =   360
         Width           =   1932
      End
      Begin Label lblAddIon 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Valence"
         Height          =   192
         Index           =   0
         Left            =   180
         TabIndex        =   23
         Top             =   1020
         Width           =   792
      End
      Begin Label lblAddIon 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "L.I.C."
         Height          =   192
         Index           =   1
         Left            =   180
         TabIndex        =   24
         Top             =   1260
         Width           =   792
      End
   End
   Begin CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   372
      Left            =   2220
      TabIndex        =   12
      Top             =   6000
      Width           =   732
   End
   Begin Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ion:"
      Height          =   192
      Left            =   300
      LinkItem        =   "Name"
      TabIndex        =   30
      Top             =   240
      Width           =   852
   End
End
Option Explicit

Dim Temp_Text As String
Dim IsError As Integer

Sub cboAddIonUnits_Click (Index As Integer)
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

Sub cboAnion_Click ()

    'Anion
    ChangedIon.Kinetic.NernstHaskellAnion = NernstHaskell.Anion(cboAnion.ListIndex + 1)
    lblAddIonValue(0).Caption = "-" & Format$(ChangedIon.Kinetic.NernstHaskellAnion.Valence, "0")
    lblAddIonValue(1).Caption = Trim$(Str$(ChangedIon.Kinetic.NernstHaskellAnion.LimitingIonicConductance))

End Sub

Sub cboCation_Click ()

    'Cation
    ChangedIon.Kinetic.NernstHaskellCation = NernstHaskell.Cation(cboCation.ListIndex + 1)
    lblAddIonValue(2).Caption = "+" & Format$(ChangedIon.Kinetic.NernstHaskellCation.Valence, "0")
    lblAddIonValue(3).Caption = Trim$(Str$(ChangedIon.Kinetic.NernstHaskellCation.LimitingIonicConductance))
    
End Sub

Sub cmdCancel_Click ()

    frmAddComponent.Hide

End Sub

Sub cmdOK_Click ()
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

       NumberOfCations = NumberOfCations + 1
       NumberOfIons = NumberOfCations
       Cation(NumberOfCations) = ChangedIon
'       Call CalculateSumEquivInitialConc
       frmIonExchangeMain!cboIons(0).AddItem Cation(NumberOfCations).Name
       frmIonExchangeMain!cboIons(2).AddItem Cation(NumberOfCations).Name
'       frmIonExchangeMain!cboKinDimComponent.AddItem Cation(NumberOfCations).Name
       frmInputKineticParameters!cboIon.AddItem Cation(NumberOfCations).Name
       If NumberOfCations > 1 Then
          frmIonExchangeMain!lstIons(0).AddItem Cation(NumberOfCations).Name
       Else
          frmIonExchangeMain!cboIons(0).ListIndex = 0
       End If

       Call CalculateKineticParameters

'       OldIonToEdit = NumberOfIonToEdit
'       For I = 1 To NumberOfCations
'           NumberOfIonToEdit = I
'           Call CalculateDimensionlessGroups
'       Next I
'       NumberOfIonToEdit = OldIonToEdit

       For i = 1 To NumberOfCations
           OneDimSeparationFactors(i) = Cation(i).SeparationFactor
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

       NumberOfAnions = NumberOfAnions + 1
       NumberOfIons = NumberOfAnions
       Anion(NumberOfAnions) = ChangedIon
'       Call CalculateSumEquivInitialConc
       frmIonExchangeMain!cboIons(1).AddItem Anion(NumberOfAnions).Name
       frmIonExchangeMain!cboIons(2).AddItem Anion(NumberOfAnions).Name
'       frmIonExchangeMain!cboKinDimComponent.AddItem Anion(NumberOfAnions).Name
       frmInputKineticParameters!cboIon.AddItem Anion(NumberOfAnions).Name

       If NumberOfAnions > 1 Then
          frmIonExchangeMain!lstIons(1).AddItem Anion(NumberOfAnions).Name
       Else
          frmIonExchangeMain!cboIons(1).ListIndex = 0
       End If
       
       Call CalculateKineticParameters
       
'       OldIonToEdit = NumberOfIonToEdit
'       For I = 1 To NumberOfAnions
'           NumberOfIonToEdit = I
'           Call CalculateDimensionlessGroups
'       Next I
'       NumberOfIonToEdit = OldIonToEdit

       For i = 1 To NumberOfAnions
           OneDimSeparationFactors(i) = Anion(i).SeparationFactor
       Next i
       Call CalculateSeparationFactors
  

       If frmIonExchangeMain!cboIons(2).ListCount = 1 Then
          frmIonExchangeMain!cboIons(2).ListIndex = 0
'          frmIonExchangeMain!cboKinDimComponent.ListIndex = 0
          frmInputKineticParameters!cboIon.ListIndex = 0
          frmIonExchangeMain!fraKineticDimensionless.Enabled = True
       End If

    ElseIf frmAddComponent.Caption = "Edit Cation" Then
       Cation(NumberOfIonToEdit) = ChangedIon
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

       For i = 1 To NumberOfCations
           OneDimSeparationFactors(i) = Cation(i).SeparationFactor
       Next i
       Call CalculateSeparationFactors
     
    ElseIf frmAddComponent.Caption = "Edit Anion" Then
       Anion(NumberOfIonToEdit) = ChangedIon
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

       For i = 1 To NumberOfAnions
           OneDimSeparationFactors(i) = Anion(i).SeparationFactor
       Next i
       Call CalculateSeparationFactors

    End If

    frmAddComponent.Hide

End Sub

Sub cmdViewSeparationFactors_Click ()
    Dim i As Integer

    For i = 1 To MAX_CHEMICAL
        OldOneDimSeparationFactors(i) = OneDimSeparationFactors(i)
    Next i
    OldOptionButtonSeparationFactors = SeparationFactorInput.Value

    frmSeparationFactors.Show 1

End Sub

Sub Form_Load ()
    Dim PositionLeft As Integer

    frmAddComponent.WindowState = 0
    frmAddComponent.Width = 4400
    frmAddComponent.Height = 6800

    'Position the form on the screen (Centered in Right Half of It)
    If WindowState = 0 Then
       'don't attempt if screen Minimized or Maximized
       PositionLeft = frmIonExchangeMain.Left + frmIonExchangeMain.Width - Screen.Width / 2
       PositionLeft = (PositionLeft / 2 - frmAddComponent.Width / 2)
       Move (Screen.Width / 2 + PositionLeft), (Screen.Height - frmAddComponent.Height) / 2
    End If

End Sub

Sub spnValence_SpinDown ()

    If CInt(lblValence.Caption) = 1 Then
       lblValence.Caption = 10
    Else
       lblValence.Caption = Str$(CInt(lblValence.Caption) - 1)
    End If
    ChangedIon.Valence = CDbl(lblValence.Caption)
    Call UpdateInitialConcentrationValence
    Call CalculateEquivalentInitialConc(ChangedIon.EquivalentInitialConcentration, ChangedIon.InitialConcentration, ChangedIon.Valence, ChangedIon.MolecularWeight)

End Sub

Sub spnValence_SpinUp ()

    If CInt(lblValence.Caption) = 10 Then
       lblValence.Caption = 1
    Else
       lblValence.Caption = Str$(CInt(lblValence.Caption) + 1)
    End If
    ChangedIon.Valence = CDbl(lblValence.Caption)
    Call UpdateInitialConcentrationValence
    Call CalculateEquivalentInitialConc(ChangedIon.EquivalentInitialConcentration, ChangedIon.InitialConcentration, ChangedIon.Valence, ChangedIon.MolecularWeight)

End Sub

Sub txtAddIon_GotFocus (Index As Integer)
    Call TextGetFocus(txtAddIon(Index), Temp_Text)
End Sub

Sub txtAddIon_KeyPress (Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
       KeyAscii = 0
       SendKeys "{Tab}"
       Exit Sub
    End If

    If Index <> 0 Then
       Call NumberCheck(KeyAscii)
    End If

End Sub

Sub txtAddIon_LostFocus (Index As Integer)
    Dim ValueChanged As Integer, NewValue As Double
    Dim msg As String, CurrentUnits As Integer
    Dim OldValue As Double, ValueToDisplay As Double
    Dim i As Integer, InitialConcentrationUnits As Integer
    Dim Input_Ion_Name As String

    If Index = 0 Then   'Inputting Name of Ion
       Input_Ion_Name = Trim$(txtAddIon(0).Text)
       If Input_Ion_Name = Trim$(Temp_Text) Then Exit Sub
       If Trim$(frmAddComponent.Caption) = "Add Cation" Then
          For i = 1 To NumberOfCations
              If Trim$(Cation(i).Name) = Input_Ion_Name Then
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
          For i = 1 To NumberOfAnions
              If Trim$(Anion(i).Name) = Input_Ion_Name Then
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
             If NumberOfCations = 0 Then lblAlpha(2).Caption = lblAlpha(1).Caption
          Else
             lblAlpha(2).Caption = Trim$(txtAddIon(0).Text)
             If NumberOfCations = 0 Then lblAlpha(1).Caption = lblAlpha(2).Caption
          End If

       End If

       If Trim$(frmAddComponent.Caption) = "Add Anion" Then
          For i = 1 To NumberOfAnions
              If Trim$(Anion(i).Name) = Input_Ion_Name Then
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
          For i = 1 To NumberOfCations
              If Trim$(Cation(i).Name) = Input_Ion_Name Then
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
             If NumberOfAnions = 0 Then lblAlpha(2).Caption = lblAlpha(1).Caption
          Else
             lblAlpha(2).Caption = Trim$(txtAddIon(0).Text)
             If NumberOfAnions = 0 Then lblAlpha(1).Caption = lblAlpha(2).Caption
          End If
       End If

       ChangedIon.Name = Trim$(txtAddIon(0).Text)

       If AddingCation Then
          Cation(NumberOfIons).Name = ChangedIon.Name
       ElseIf AddingAnion Then
          Anion(NumberOfIons).Name = ChangedIon.Name
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

Sub txtAlphaValue_GotFocus ()
    Call TextGetFocus(txtAlphaValue, Temp_Text)
End Sub

Sub txtAlphaValue_KeyPress (KeyAscii As Integer)

    If KeyAscii = 13 Then
       KeyAscii = 0
       SendKeys "{Tab}"
       Exit Sub
    End If

    Call NumberCheck(KeyAscii)

End Sub

Sub txtAlphaValue_LostFocus ()
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

Sub UpdateInitialConcentrationMolWt ()
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

Sub UpdateInitialConcentrationValence ()
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

