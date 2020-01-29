Attribute VB_Name = "PrintModule"
Option Explicit

'Global frmPrint_DO_INPUTS As Boolean
'Global frmPrint_DO_OUTPUTS As Boolean
'Global frmPrint_DO_PLOTS As Boolean

Global Const USE_FONTNAME = "arial"
Global Const USE_FONTSIZE = 8
Global Const USE_FORMAT_CURRENCYSTANDARD = "$#,##0_);[Red]($#,##0)"
Global Const USE_FORMAT_CURRENCYDIGITSPAST2 = "$#,##0.00_);[Red]($#,##0.00)"

Type TankConcLabels_Type
  Label1 As String        'LINE 1 OF OUTPUT.
  Label2 As String        'LINE 2 OF OUTPUT (UNITS).
End Type
Global TankConcLabels() As TankConcLabels_Type
Global TankConcs() As Double       'GMOL/L OR MG/L
    'TankConcs(i,j,k): i=CHEMICAL #, j=TANK #, k=ROW #.
Global Tank_Times() As Double      'MINUTES





Const PrintModule_declarations_end = 0


Sub PrintCell( _
    f1 As Control, _
    r As Integer, _
    c As Integer, _
    v As Variant, _
    do_italics As Boolean, _
    do_bold As Boolean, _
    do_rightjustify)
Dim use_HAlign As Integer
  f1.EntryRC(r, c) = v
  f1.SetSelection r, c, r, c
  f1.SetFont _
      USE_FONTNAME, _
      USE_FONTSIZE, _
      do_bold, _
      do_italics, _
      False, _
      False, _
      QBColor(0), _
      False, _
      False
  use_HAlign = F1HAlignLeft
  If (do_rightjustify) Then use_HAlign = F1HAlignRight
  f1.SetAlignment _
      use_HAlign, _
      False, _
      F1VAlignBottom, _
      0
    'F1Book1.SetFont pFont, nSize, bBold, bItalic,
    '  bUnderline, bStrikeout, crColor, bOutline,
    '  bShadow
    'F1Book1.SetAlignment HAlign, bWordWrap, _
    '  VAlign, nOrientation
End Sub
Sub PrintCell_CurrencyStandard( _
    f1 As Control, _
    r As Integer, _
    c As Integer, _
    v As Variant)
  Call PrintCell(f1, r, c, v, False, False, True)
  f1.NumberFormat = USE_FORMAT_CURRENCYSTANDARD
End Sub
Sub PrintCell_CurrencyDigitsPast2( _
    f1 As Control, _
    r As Integer, _
    c As Integer, _
    v As Variant)
  Call PrintCell(f1, r, c, v, False, False, True)
  f1.NumberFormat = USE_FORMAT_CURRENCYDIGITSPAST2
End Sub
Sub PrintCell_QuantityStandard( _
    f1 As Control, _
    r As Integer, _
    c As Integer, _
    v As Variant)
Dim AbsValue As Double
Dim GetDoubleFormat As String
AbsValue = Abs(val(v))
  Select Case AbsValue
    Case 0#
      GetDoubleFormat = "0"
    'Case Is < 0.001
    '  GetDoubleFormat = "0.00E+00"
    Case Is < 0.01
      GetDoubleFormat = "0.00E+00"
    Case Is < 0.1
      GetDoubleFormat = "0.0000"
    Case Is < 1
      GetDoubleFormat = "0.000"
    Case Is < 10
      GetDoubleFormat = "0.00"
    Case Is < 100
      GetDoubleFormat = "0.0"
    Case Is < 1000
      GetDoubleFormat = "0"
    Case Is < 1000# * 1000# * 1000#
      GetDoubleFormat = "0"
    Case Else
      GetDoubleFormat = "0.00E+00"
  End Select
  Call PrintCell(f1, r, c, v, False, False, True)
  f1.NumberFormat = GetDoubleFormat
End Sub
Sub Print_Border0(f1 As Control, r1 As Integer, c1 As Integer, r2 As Integer, c2 As Integer, num_top_rows As Integer, num_left_cols As Integer)
Dim cc As Variant
  cc = QBColor(0)
  f1.SetSelection r1, c1, r2, c2
  f1.SetBorder 5, 0, 0, 0, 0, 1, cc, cc, cc, cc, cc
  If (r1 <> r2) Then
    f1.SetSelection r1 + num_top_rows, c1, r2, c2
    f1.SetBorder 5, -1, -1, -1, -1, 1, cc, cc, cc, cc, cc
  End If
  f1.SetSelection r1, c1 + num_left_cols, r2, c2
  f1.SetBorder 5, -1, -1, -1, -1, 1, cc, cc, cc, cc, cc
End Sub
Sub Print_Border(f1 As Control, r1 As Integer, c1 As Integer, r2 As Integer, c2 As Integer)
  Call Print_Border0(f1, r1, c1, r2, c2, 1, 2)
End Sub


Sub Print_Inputs(f1 As Control, proj As Project_Type, SheetIdx As Integer)
Dim i As Integer
Dim j As Integer
Dim r As Integer
Dim r0 As Integer
Dim USER_OPTION_DRAWBORDERS As Integer
Dim MAX_COLUMN As Long
Dim MAX_COLUMNWIDTH As Long
Dim temp As String
Dim temp1 As String
Dim temp2 As String
Dim temp3 As String
Dim sname As String
Dim tc As TargetCompound_Type

Dim is_nom As Boolean
Dim rthis As Integer
Dim wl As Wavelength_Type

Dim is_h2o2 As Boolean
Dim this_extcoef As Double
Dim this_quatyd As Double
Dim this_name As String

  f1.Sheet = SheetIdx
    
  USER_OPTION_DRAWBORDERS = True
  
  'SECTION: "TOP HEADER".
  r = 2
  Call PrintCell(f1, r + 0, 3, "Filename:", True, False, True)
  Call PrintCell(f1, r + 1, 3, "Unused:", True, False, True)
  Call PrintCell(f1, r + 2, 3, "Printed:", True, False, True)
  Call PrintCell(f1, r + 0, 4, Current_Filename, False, False, False)
  Call PrintCell(f1, r + 1, 4, "Unused", False, False, False)
  Call PrintCell(f1, r + 2, 4, Now, False, False, False)
  r = r + 3     'move to immediately after this section.
  
  'SECTION: "Reactor Properties."
  r = r + 3     'this many lines of space between sections.
  sname = "Reactor Properties."
  Call PrintCell(f1, r + 0, 1, sname, False, True, False)
  Call PrintCell(f1, r + 2, 3, "Reactor Type:", True, False, True)
  Call PrintCell(f1, r + 3, 3, "Volume:", True, False, True)
  Call PrintCell(f1, r + 4, 3, "Retention Time:", True, False, True)
  temp = IIf(proj.idreact = IDREACT_CMBR, "Initial H2O2:", "Influent H2O2:")
  Call PrintCell(f1, r + 5, 3, temp, True, False, True)
  Call PrintCell(f1, r + 6, 3, "Number of Tanks:", True, False, True)
  'Call PrintCell(f1, r + 2, 5, "-", True, False, False)
  Call PrintCell(f1, r + 3, 5, "liters", True, False, False)
  Call PrintCell(f1, r + 4, 5, "minutes", True, False, False)
  Call PrintCell(f1, r + 5, 5, "gmol/L", True, False, False)
  Call PrintCell(f1, r + 6, 5, "tanks", True, False, False)
  temp = IIf(proj.idreact = IDREACT_CMBR, "CMBR", "CMFR")
  Call PrintCell(f1, r + 2, 4, temp, False, False, True)
  Call PrintCell_QuantityStandard(f1, r + 3, 4, proj.volume)
  Select Case proj.idreact
    Case IDREACT_CMBR:
      Call PrintCell(f1, r + 4, 4, "n/a", True, False, True)
    Case IDREACT_CMFR:
      Call PrintCell_QuantityStandard(f1, r + 4, 4, proj.tau)
  End Select
  Call PrintCell_QuantityStandard(f1, r + 5, 4, proj.inf_h2o2)
  Select Case proj.idreact
    Case IDREACT_CMBR:
      Call PrintCell(f1, r + 6, 4, "n/a", True, False, True)
    Case IDREACT_CMFR:
      Call PrintCell_QuantityStandard(f1, r + 6, 4, proj.num_tanks)
  End Select
  r = r + 7         'move to immediately after this section.
  
  'SECTION: "Numerical Simulation Parameters."
  r = r + 2     'this many lines of space between sections.
  sname = "Numerical Simulation Parameters."
  Call PrintCell(f1, r + 0, 1, sname, False, True, False)
  Call PrintCell(f1, r + 2, 3, "Time Step:", True, False, True)
  Call PrintCell(f1, r + 3, 3, "Final Time:", True, False, True)
  Call PrintCell(f1, r + 4, 3, "# Retention Times to Simulate:", True, False, True)
  Call PrintCell(f1, r + 2, 5, "minutes", True, False, False)
  Call PrintCell(f1, r + 3, 5, "minutes", True, False, False)
  Call PrintCell(f1, r + 4, 5, "-", True, False, False)
  Call PrintCell_QuantityStandard(f1, r + 2, 4, proj.ssize / 60#)
  Select Case proj.idreact
    Case IDREACT_CMBR:
      Call PrintCell_QuantityStandard(f1, r + 3, 4, proj.ttotal)
    Case IDREACT_CMFR:
      Call PrintCell(f1, r + 3, 4, "n/a", True, False, True)
  End Select
  Select Case proj.idreact
    Case IDREACT_CMBR:
      Call PrintCell(f1, r + 4, 4, "n/a", True, False, True)
    Case IDREACT_CMFR:
      Call PrintCell_QuantityStandard(f1, r + 4, 4, proj.xntimes)
  End Select
  r = r + 5         'move to immediately after this section.
  
  'SECTION: "Water Quality Properties."
  r = r + 2     'this many lines of space between sections.
  sname = "Water Quality Properties."
  Call PrintCell(f1, r + 0, 1, sname, False, True, False)
  temp = IIf(proj.idreact = IDREACT_CMBR, "Initial pH:", "Influent pH:")
  Call PrintCell(f1, r + 2, 3, temp, True, False, True)
  temp = IIf(proj.idreact = IDREACT_CMBR, "Initial Phosphate Conc.:", "Influent Phosphate Conc.:")
  Call PrintCell(f1, r + 3, 3, temp, True, False, True)
  Call PrintCell(f1, r + 4, 3, "TIC Input As:", True, False, True)
  temp2 = IIf(proj.idcarbn = IDCARBN_TIC, "TIC", "Alkalinity")
  temp = IIf(proj.idreact = IDREACT_CMBR, "Initial", "Influent")
  temp = temp & " " & temp2 & " Concentration:"
  Call PrintCell(f1, r + 5, 3, temp, True, False, True)
  Call PrintCell(f1, r + 2, 5, "-", True, False, False)
  Call PrintCell(f1, r + 3, 5, "gmol/L", True, False, False)
  'Call PrintCell(f1, r + 4, 5, "-", True, False, False)
  temp = IIf(proj.idcarbn = IDCARBN_TIC, "gmol/L", "mg/L as CaCO3")
  Call PrintCell(f1, r + 5, 5, temp, True, False, False)
  Call PrintCell_QuantityStandard(f1, r + 2, 4, proj.ph0)
  Call PrintCell_QuantityStandard(f1, r + 3, 4, proj.phosph)
  Select Case proj.idcarbn
    Case IDCARBN_TIC:
      Call PrintCell(f1, r + 4, 4, "TIC", False, False, True)
      Call PrintCell_QuantityStandard(f1, r + 5, 4, proj.ticarbn)
    Case IDCARBN_ALKALINITY:
      Call PrintCell(f1, r + 4, 4, "Alkalinity", False, False, True)
      Call PrintCell_QuantityStandard(f1, r + 5, 4, proj.alk)
  End Select
  r = r + 6         'move to immediately after this section.
  
  'SECTION: "Target Compounds : Properties of Protonated Form."
  r = r + 2     'this many lines of space between sections.
  sname = "Target Compounds : Properties of Protonated Form."
  Call PrintCell(f1, r + 0, 1, sname, False, True, False)
  Call PrintCell(f1, r + 2, 2, "", True, False, False)
  temp = IIf(proj.idreact = IDREACT_CMBR, "Initial", "Influent")
  Call PrintCell(f1, r + 2, 3, temp, True, False, True)
  Call PrintCell(f1, r + 2, 4, "", True, False, True)
  Call PrintCell(f1, r + 2, 5, "", True, False, True)
  Call PrintCell(f1, r + 2, 6, "# C atoms", True, False, True)
  Call PrintCell(f1, r + 2, 7, "# hal. atoms", True, False, True)
  Call PrintCell(f1, r + 2, 8, "2nd order", True, False, True)
  Call PrintCell(f1, r + 3, 2, "Name", True, False, False)
  Call PrintCell(f1, r + 3, 3, "Conc.", True, False, True)
  Call PrintCell(f1, r + 3, 4, "Valence", True, False, True)
  Call PrintCell(f1, r + 3, 5, "Molec. Wt.", True, False, True)
  Call PrintCell(f1, r + 3, 6, "per molec.", True, False, True)
  Call PrintCell(f1, r + 3, 7, "per molec.", True, False, True)
  Call PrintCell(f1, r + 3, 8, "rate const.", True, False, True)
  Call PrintCell(f1, r + 4, 2, "-", True, False, False)
  Call PrintCell(f1, r + 4, 3, "gmol/L", True, False, True)
  Call PrintCell(f1, r + 4, 4, "-", True, False, True)
  Call PrintCell(f1, r + 4, 5, "g/gmol", True, False, True)
  Call PrintCell(f1, r + 4, 6, "-", True, False, True)
  Call PrintCell(f1, r + 4, 7, "-", True, False, True)
  Call PrintCell(f1, r + 4, 8, "(*)", True, False, True)
  For i = 1 To proj.TargetCompounds_Count
    is_nom = False
    tc = proj.TargetCompounds(i)
    If (Trim$(UCase$(tc.comname)) = Trim$(UCase$("NOM"))) Then
      is_nom = True
    End If
    rthis = r + 4 + i
    Call PrintCell(f1, rthis, 2, tc.comname, False, False, False)
    Call PrintCell_QuantityStandard(f1, rthis, 3, tc.concini)
    Call PrintCell_QuantityStandard(f1, rthis, 5, tc.mw)
    Call PrintCell_QuantityStandard(f1, rthis, 8, tc.xk)
    If (is_nom) Then
      Call PrintCell(f1, rthis, 4, "n/a", True, False, True)
      Call PrintCell(f1, rthis, 6, "n/a", True, False, True)
      Call PrintCell(f1, rthis, 7, "n/a", True, False, True)
    Else
      Call PrintCell_QuantityStandard(f1, rthis, 4, tc.val)
      Call PrintCell_QuantityStandard(f1, rthis, 6, tc.ncarbn)
      Call PrintCell_QuantityStandard(f1, rthis, 7, tc.nsubstt)
    End If
  Next i
  Call PrintCell(f1, r + 4 + 1 + proj.TargetCompounds_Count, 2, _
      "(*): For NOM, units are 1/(mg-L)-s; for all other compounds, units are L/gmol-s.", False, False, False)
  If (USER_OPTION_DRAWBORDERS) Then
    Call Print_Border0(f1, r + 2, 2, r + 4 + proj.TargetCompounds_Count, 8, 3, 1)
  End If
  r = r + 4 + i + 1 'move to immediately after this section.
   
  'SECTION: "Target Compounds : Equilibrium Reaction, Properties of De-protonated Form."
  r = r + 2     'this many lines of space between sections.
  sname = "Target Compounds : Equilibrium Reaction, Properties of De-protonated Form."
  Call PrintCell(f1, r + 0, 1, sname, False, True, False)
  Call PrintCell(f1, r + 2, 2, "", True, False, False)
  Call PrintCell(f1, r + 2, 3, "Equilibrium", True, False, True)
  Call PrintCell(f1, r + 2, 4, "", True, False, True)
  Call PrintCell(f1, r + 2, 5, "", True, False, True)
  Call PrintCell(f1, r + 2, 6, "2nd order", True, False, True)
  Call PrintCell(f1, r + 3, 2, "Name", True, False, False)
  Call PrintCell(f1, r + 3, 3, "Constant", True, False, True)
  Call PrintCell(f1, r + 3, 4, "Valence", True, False, True)
  Call PrintCell(f1, r + 3, 5, "Molec. Wt.", True, False, True)
  Call PrintCell(f1, r + 3, 6, "rate const.", True, False, True)
  Call PrintCell(f1, r + 4, 2, "-", True, False, False)
  Call PrintCell(f1, r + 4, 3, "gmol/L", True, False, True)
  Call PrintCell(f1, r + 4, 4, "-", True, False, True)
  Call PrintCell(f1, r + 4, 5, "g/gmol", True, False, True)
  Call PrintCell(f1, r + 4, 6, "L/gmol-s", True, False, True)
  For i = 1 To proj.TargetCompounds_Count
    is_nom = False
    tc = proj.TargetCompounds(i)
    If (Trim$(UCase$(tc.comname)) = Trim$(UCase$("NOM"))) Then
      is_nom = True
    End If
    rthis = r + 4 + i
    Call PrintCell(f1, rthis, 2, tc.comname, False, False, False)
    If (is_nom) Then
      Call PrintCell(f1, rthis, 3, "n/a", True, False, True)
      Call PrintCell(f1, rthis, 4, "n/a", True, False, True)
      Call PrintCell(f1, rthis, 5, "n/a", True, False, True)
      Call PrintCell(f1, rthis, 6, "n/a", True, False, True)
    Else
      Call PrintCell_QuantityStandard(f1, rthis, 3, tc.dep_xke)
      Call PrintCell_QuantityStandard(f1, rthis, 4, tc.dep_val)
      Call PrintCell_QuantityStandard(f1, rthis, 5, tc.dep_mw)
      Call PrintCell_QuantityStandard(f1, rthis, 6, tc.dep_xk)
    End If
  Next i
  If (USER_OPTION_DRAWBORDERS) Then
    Call Print_Border0(f1, r + 2, 2, r + 4 + proj.TargetCompounds_Count, 6, 3, 1)
  End If
  r = r + 4 + i     'move to immediately after this section.
   
  'SECTION: "Target Compounds : Other Reactions."
  r = r + 2     'this many lines of space between sections.
  sname = "Target Compounds : Other Reactions."
  Call PrintCell(f1, r + 0, 1, sname, False, True, False)
  Call PrintCell(f1, r + 2, 2, "", True, False, False)
  Call PrintCell(f1, r + 2, 3, "Rxn. With", True, False, True)
  Call PrintCell(f1, r + 2, 4, "Rxn. With", True, False, True)
  Call PrintCell(f1, r + 2, 5, "Rxn. With", True, False, True)
  Call PrintCell(f1, r + 2, 6, "Rxn. With", True, False, True)
  Call PrintCell(f1, r + 3, 2, "Name", True, False, False)
  Call PrintCell(f1, r + 3, 3, "CO3*-", True, False, True)
  Call PrintCell(f1, r + 3, 4, "HPO4*-", True, False, True)
  Call PrintCell(f1, r + 3, 5, "O2*-", True, False, True)
  Call PrintCell(f1, r + 3, 6, "HO2*", True, False, True)
  Call PrintCell(f1, r + 4, 2, "-", True, False, False)
  Call PrintCell(f1, r + 4, 3, "L/gmol-s", True, False, True)
  Call PrintCell(f1, r + 4, 4, "L/gmol-s", True, False, True)
  Call PrintCell(f1, r + 4, 5, "L/gmol-s", True, False, True)
  Call PrintCell(f1, r + 4, 6, "L/gmol-s", True, False, True)
  For i = 1 To proj.TargetCompounds_Count
    is_nom = False
    tc = proj.TargetCompounds(i)
    If (Trim$(UCase$(tc.comname)) = Trim$(UCase$("NOM"))) Then
      is_nom = True
    End If
    rthis = r + 4 + i
    Call PrintCell(f1, rthis, 2, tc.comname, False, False, False)
    If (is_nom) Then
      Call PrintCell(f1, rthis, 3, "n/a", True, False, True)
      Call PrintCell(f1, rthis, 4, "n/a", True, False, True)
      Call PrintCell(f1, rthis, 5, "n/a", True, False, True)
      Call PrintCell(f1, rthis, 6, "n/a", True, False, True)
    Else
      Call PrintCell_QuantityStandard(f1, rthis, 3, tc.xk_co3XM)
      Call PrintCell_QuantityStandard(f1, rthis, 4, tc.xk_hpo4XM)
      Call PrintCell_QuantityStandard(f1, rthis, 5, tc.xk_o2XM)
      Call PrintCell_QuantityStandard(f1, rthis, 6, tc.xk_ho2X)
    End If
  Next i
  If (USER_OPTION_DRAWBORDERS) Then
    Call Print_Border0(f1, r + 2, 2, r + 4 + proj.TargetCompounds_Count, 6, 3, 1)
  End If
  r = r + 4 + i      'move to immediately after this section.
   
  'SECTION: "Photochemical Parameters."
  r = r + 3     'this many lines of space between sections.
  sname = "Photochemical Parameters."
  Call PrintCell(f1, r + 0, 1, sname, False, True, False)
  Call PrintCell(f1, r + 2, 3, "Lamp Power:", True, False, True)
  Call PrintCell(f1, r + 3, 3, "Lamp Name:", True, False, True)
  Call PrintCell(f1, r + 4, 3, "UV Path Length:", True, False, True)
  Call PrintCell(f1, r + 5, 3, "Light Specification Method:", True, False, True)
  Call PrintCell(f1, r + 2, 5, "watts", True, False, False)
  'Call PrintCell(f1, r + 3, 5, "", True, False, False)
  Call PrintCell(f1, r + 4, 5, "cm", True, False, False)
  'Call PrintCell(f1, r + 5, 5, "", True, False, False)
  Call PrintCell_QuantityStandard(f1, r + 2, 4, proj.lamp_power)
  Call PrintCell(f1, r + 3, 4, proj.lamp_name, False, False, False)
  Call PrintCell_QuantityStandard(f1, r + 4, 4, proj.uvpathl)
  Select Case proj.iduvi
    Case IDUVI_EINSTEINS_L_S: temp = "Intensity, in Einsteins/L-s"
    Case IDUVI_WATTS: temp = "Intensity, in Watts"
    Case IDUVI_EFFICIENCY: temp = "Efficiency, dimensionless (range: 0-1)"
  End Select
  Call PrintCell(f1, r + 5, 4, temp, False, False, False)
  r = r + 6     'move to immediately after this section.
  
  'SECTION: "Photochemical Parameters : Light Intensities."
  r = r + 2     'this many lines of space between sections.
  sname = "Photochemical Parameters : Light Intensities."
  Call PrintCell(f1, r + 0, 1, sname, False, True, False)
  Select Case proj.iduvi
    Case IDUVI_EINSTEINS_L_S:
      temp1 = "UV Light"
      temp2 = "Intensity"
      temp3 = "Ein./L-s"
    Case IDUVI_WATTS:
      temp1 = "UV Light"
      temp2 = "Intensity"
      temp3 = "watts"
    Case IDUVI_EFFICIENCY:
      temp1 = ""
      temp2 = "Efficiency"
      temp3 = "dim'less"
  End Select
  Call PrintCell(f1, r + 2, 2, "", True, False, True)
  Call PrintCell(f1, r + 2, 3, temp1, True, False, True)
  Call PrintCell(f1, r + 3, 2, "Wavelength", True, False, True)
  Call PrintCell(f1, r + 3, 3, temp2, True, False, True)
  Call PrintCell(f1, r + 4, 2, "nm", True, False, True)
  Call PrintCell(f1, r + 4, 3, temp3, True, False, True)
  For i = 1 To proj.Wavelength_Count
    wl = proj.Wavelengths(i)
    rthis = r + 4 + i
    Call PrintCell_QuantityStandard(f1, rthis, 2, wl.lwave)
    Call PrintCell_QuantityStandard(f1, rthis, 3, wl.uvi)
  Next i
  If (USER_OPTION_DRAWBORDERS) Then
    Call Print_Border0(f1, r + 2, 2, r + 4 + proj.Wavelength_Count, 3, 3, 1)
  End If
  r = r + 4 + i      'move to immediately after this section.

  'SECTION: "Photochemical Parameters : Values for Each Compound."
  r = r + 2     'this many lines of space between sections.
  sname = "Photochemical Parameters : Values for Each Compound."
  Call PrintCell(f1, r + 0, 1, sname, False, True, False)
  Call PrintCell(f1, r + 2, 2, "", True, False, False)
  Call PrintCell(f1, r + 2, 3, "", True, False, True)
  Call PrintCell(f1, r + 2, 4, "Extinction", True, False, True)
  Call PrintCell(f1, r + 2, 5, "Quantum", True, False, True)
  Call PrintCell(f1, r + 3, 2, "Name", True, False, False)
  Call PrintCell(f1, r + 3, 3, "Wavelength", True, False, True)
  Call PrintCell(f1, r + 3, 4, "Coefficient", True, False, True)
  Call PrintCell(f1, r + 3, 5, "Yield", True, False, True)
  Call PrintCell(f1, r + 4, 2, "-", True, False, False)
  Call PrintCell(f1, r + 4, 3, "nm", True, False, True)
  Call PrintCell(f1, r + 4, 4, "-", True, False, True)
  Call PrintCell(f1, r + 4, 5, "-", True, False, True)
  For i = 1 To proj.TargetCompounds_Count + 1
    is_nom = False
    is_h2o2 = False
    If (i <= proj.TargetCompounds_Count) Then
      tc = proj.TargetCompounds(i)
      If (Trim$(UCase$(tc.comname)) = Trim$(UCase$("NOM"))) Then is_nom = True
      this_name = tc.comname
    Else
      is_h2o2 = True
      this_name = "H2O2"
    End If
    For j = 1 To proj.Wavelength_Count
      wl = proj.Wavelengths(j)
      rthis = r + 4 + j + (i - 1) * proj.Wavelength_Count
      If (is_h2o2) Then
        this_extcoef = proj.extcoef_h2o2(j)
        this_quatyd = proj.quatyd_h2o2(j)
      Else
        this_extcoef = proj.extcoef(i, j)
        this_quatyd = proj.quatyd(i, j)
      End If
      Call PrintCell(f1, rthis, 2, this_name, False, False, False)
      Call PrintCell_QuantityStandard(f1, rthis, 3, wl.lwave)
      Call PrintCell_QuantityStandard(f1, rthis, 4, this_extcoef)
      Call PrintCell_QuantityStandard(f1, rthis, 5, this_quatyd)
    Next j
  Next i
  If (USER_OPTION_DRAWBORDERS) Then
    Call Print_Border0(f1, r + 2, 2, r + 4 + proj.Wavelength_Count * (proj.TargetCompounds_Count + 1), 5, 3, 1)
  End If
  r = r + 4 + i      'move to immediately after this section.
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  'RESIZE THE COLUMNS.
  MAX_COLUMN = 20
  MAX_COLUMNWIDTH = 3000
  For i = 1 To MAX_COLUMN
    f1.ColWidth(i) = MAX_COLUMNWIDTH
  Next i

  'LAST STEP: RETURN CURSOR TO POSITION 1,1.
  f1.SetSelection 1, 1, 1, 1

End Sub


Sub Print_Outputs(f1 As Control, proj As Project_Type, SheetIdx As Integer, tank_number As Integer)
Dim r As Integer
Dim USER_OPTION_DRAWBORDERS As Boolean
Dim sname As String
Dim MAX_COLUMN As Long
Dim MAX_COLUMNWIDTH As Long

Dim num_chemicals As Integer
Dim num_tanks As Integer
Dim num_rows As Integer
Dim rthis As Integer
Dim cthis As Integer
Dim i As Integer
Dim j As Integer
Dim this_tank As Integer

  f1.Sheet = SheetIdx
    
  USER_OPTION_DRAWBORDERS = True
  
  'SECTION: "TOP HEADER".
  r = 2
  Call PrintCell(f1, r + 0, 3, "Filename:", True, False, True)
  Call PrintCell(f1, r + 1, 3, "Tank Number:", True, False, True)
  Call PrintCell(f1, r + 2, 3, "Printed:", True, False, True)
  Call PrintCell(f1, r + 0, 4, Current_Filename, False, False, False)
  Call PrintCell(f1, r + 1, 4, Trim$(Str$(tank_number)), False, False, True)
  Call PrintCell(f1, r + 2, 4, Now, False, False, False)
  r = r + 3     'move to immediately after this section.
  
  'SECTION: "Current Tank : Concentrations of All Species."
  r = r + 3     'this many lines of space between sections.
  sname = "Current Tank : Concentrations of All Species."
  Call PrintCell(f1, r + 0, 1, sname, False, True, False)
  num_chemicals = UBound(TankConcs, 1)
  num_tanks = UBound(TankConcs, 2)
  num_rows = UBound(TankConcs, 3)
  this_tank = tank_number
  For i = 1 To num_chemicals
    Call PrintCell(f1, r + 2, 1, "Tank #", True, False, True)
    Call PrintCell(f1, r + 3, 1, "-", True, False, True)
    Call PrintCell(f1, r + 2, 2, "Time", True, False, True)
    Call PrintCell(f1, r + 3, 2, "sec", True, False, True)
    Call PrintCell(f1, r + 2, 3, "Time", True, False, True)
    Call PrintCell(f1, r + 3, 3, "min", True, False, True)
    cthis = i + 3
    Call PrintCell(f1, r + 2, cthis, TankConcLabels(i).Label1, True, False, True)
    Call PrintCell(f1, r + 3, cthis, TankConcLabels(i).Label2, True, False, True)
    For j = 1 To num_rows
      rthis = r + 3 + j
      Call PrintCell_QuantityStandard(f1, rthis, cthis, TankConcs(i, this_tank, j))
      If (i = 1) Then
        Call PrintCell(f1, rthis, 1, Trim$(Str$(this_tank)), False, False, True)
        Call PrintCell_QuantityStandard(f1, rthis, 2, Tank_Times(j) * 60#)
        Call PrintCell_QuantityStandard(f1, rthis, 3, Tank_Times(j))
      End If
    Next j
  Next i
  
  
  

  
  
  
  
  
'Type TankConcLabels_Type
'  label1 As String        'LINE 1 OF OUTPUT.
'  label2 As String        'LINE 2 OF OUTPUT (UNITS).
'End Type
'Dim TankConcLabels() As TankConcLabels_Type
'Dim TankConcs() As Double
'    'TankConcs(i,j,k): i=CHEMICAL #, j=TANK #, k=ROW #.
'Dim Tank_Times() As Double



  
  
  'RESIZE THE COLUMNS.
  MAX_COLUMN = 20
  MAX_COLUMNWIDTH = 3000
  For i = 1 To MAX_COLUMN
    f1.ColWidth(i) = MAX_COLUMNWIDTH
  Next i

  'LAST STEP: RETURN CURSOR TO POSITION 1,1.
  f1.SetSelection 1, 1, 1, 1

End Sub


Sub PrintPrepare_OutputConcs(proj As Project_Type, Any_Error As Boolean)
Dim num_chemicals As Integer
Dim num_tanks As Integer
Dim num_rows As Integer
Dim fn_spec As String
Dim fn_this As String

Dim This_Conc As Double
Dim This_Time As Double
Dim Have_Loaded_Times_Already As Boolean
Dim hit_END_OF_TANK As Boolean
Dim hit_END_OF_FILE As Boolean
Dim i As Integer
Dim j As Integer
Dim f As Integer
Dim line_in As String
Dim row_count As Integer
Dim msg As String
Dim temp As String
Dim args_num As Integer

  Any_Error = False
  
  num_chemicals = -1      'DETERMINED LATER IN THIS SUBROUTINE.
  Select Case proj.idreact
    Case IDREACT_CMBR: num_tanks = 1
    Case IDREACT_CMFR: num_tanks = proj.num_tanks
  End Select
  num_tanks = proj.num_tanks
  num_rows = -1           'DETERMINED LATER IN THIS SUBROUTINE.

  'DETERMINE THE APPROPRIATE VALUE OF num_chemicals.
  'NOTE THAT num_chemicals ALSO INCLUDES SPACE FOR THE pH.
  num_chemicals = 0
  fn_spec = App.Path & "\exes\comp*.txt"
  fn_this = Trim$(Dir(fn_spec))
  Do While (1 = 1)
    If (fn_this = "") Then Exit Do
    num_chemicals = num_chemicals + 1
    fn_this = Dir
  Loop
  num_chemicals = num_chemicals + 1         'ALLOCATE SPACE FOR pH.

  'READ IN THE CONCENTRATION DATA.
  ReDim TankConcLabels(1 To num_chemicals)
  ReDim TankConcs(1 To num_chemicals, 1 To num_tanks, 1 To 10)
  ReDim Tank_Times(1 To 10)
  Have_Loaded_Times_Already = False
  For i = 1 To num_chemicals
    If (i = num_chemicals) Then
      fn_this = App.Path & "\exes\cmp_ph.txt"
    Else
      fn_this = App.Path & "\exes\comp" & Trim$(Str$(i)) & ".txt"
    End If
    If (Not FileExists(fn_this)) Then
      'ERROR: LESS THAN EXPECTED NUMBER OF COMPONENT CONCENTRATION FILES.
      msg = "There were less than the expected number of component concentration files."
      GoTo err_Corrupt_Files
    End If
    f = FreeFile
    Open fn_this For Input As #f
    Line Input #f, line_in      'COMPONENT NAME.
    TankConcLabels(i).Label1 = Trim$(line_in)
    Line Input #f, line_in      'COMPONENT UNITS.
    TankConcLabels(i).Label2 = Trim$(line_in)
    Line Input #f, line_in      'NOT USED.
    For j = 1 To num_tanks
      row_count = 0
      hit_END_OF_TANK = False
      hit_END_OF_FILE = False
      Do While (1 = 1)
        Line Input #f, line_in
        line_in = Trim$(UCase$(line_in))
        If (line_in = Trim$(UCase$("END_OF_TANK"))) Then
          hit_END_OF_TANK = True
          Exit Do
        End If
        If (line_in = Trim$(UCase$("END_OF_FILE"))) Then
          hit_END_OF_FILE = False
          Exit Do
        End If
        row_count = row_count + 1
        'PARSE THE INPUT LINE TO DETERMINE TIME AND CONCENTRATION.
        line_in = Parser_RemoveDuplicateSeparators(" ", line_in)
        args_num = Parser_GetNumArgs(" ", line_in)
        If (args_num <> 4) Then
          'ERROR: INVALID LINE IN FILE.
          Close #f
          msg = "There was an invalid line in " & fn_this & "."
          GoTo err_Corrupt_Files
        End If
        Call Parser_GetArg(" ", line_in, 3, temp)
        This_Time = CDbl(val(temp))
        Call Parser_GetArg(" ", line_in, 4, temp)
        This_Conc = CDbl(val(temp))
        'TRANSFER CONCENTRATION AND TIME INTO STORAGE.
        If (row_count > UBound(TankConcs, 3)) Then
          ReDim Preserve TankConcs(1 To num_chemicals, 1 To num_tanks, 1 To row_count)
          ReDim Preserve Tank_Times(1 To row_count)
        End If
        TankConcs(i, j, row_count) = This_Conc
        If (Not Have_Loaded_Times_Already) Then
          Tank_Times(row_count) = This_Time
        End If
      Loop
      If (Not Have_Loaded_Times_Already) Then
        Have_Loaded_Times_Already = True
      End If
      If (hit_END_OF_FILE) Then
        If (j < num_tanks) Then
          'ERROR: PREMATURE END OF FILE.
          Close #f
          msg = "There was a premature end of file in " & fn_this & "."
          GoTo err_Corrupt_Files
        End If
      End If
    Next j
    Close #f
  Next i
  
  Exit Sub
  
err_Corrupt_Files:
  Any_Error = True
  Call Show_Error(msg & "  Component output files appear to be corrupted.  Cancelling print.")
  Exit Sub
End Sub


Sub PrintTo_f1book(f1 As Control, proj As Project_Type)
Dim i As Integer
Dim j As Integer
Dim SortIndex_CaseNames() As Integer
Dim Case_Index As Integer
Dim Sheet_Index As Integer

Dim SheetIdx_Inputs As Integer
Dim SheetIdx_Outputs As Integer
Dim NumSheets_Outputs As Integer
Dim SheetIdx_ThisOutput As Integer
Dim NumSheets_Total As Integer
Dim name_out As String

Dim Any_Error As Boolean

  'PREPARE OUTPUT CONCENTRATIONS.
  Call PrintPrepare_OutputConcs(proj, Any_Error)
  If (Any_Error) Then Exit Sub

  'SET NUMBER OF SHEETS.
  SheetIdx_Inputs = 0
  If (frmPrint_DO_INPUTS) Then SheetIdx_Inputs = 1
  SheetIdx_Outputs = 0
  NumSheets_Outputs = 0
  If (frmPrint_DO_OUTPUTS) Then
    Select Case proj.idreact
      Case IDREACT_CMBR: NumSheets_Outputs = 1
      Case IDREACT_CMFR: NumSheets_Outputs = proj.num_tanks
    End Select
    SheetIdx_Outputs = SheetIdx_Inputs + 1
  End If
  If (frmPrint_DO_PLOTS) Then
    'DO NOTHING.
  End If
  NumSheets_Total = SheetIdx_Inputs + NumSheets_Outputs
  If (NumSheets_Total < 1) Then Exit Sub
  f1.NumSheets = NumSheets_Total
  
  'SET DEFAULT FONT.
  f1.SetDefaultFont USE_FONTNAME, USE_FONTSIZE
  
  ''SORT THE CASE NAMES.
  'frmMain.lsttempsorter.Clear
  'For j = 1 To proj.Cases_Count
  '  frmMain.lsttempsorter.AddItem proj.Cases(j).name
  '  frmMain.lsttempsorter.ItemData(frmMain.lsttempsorter.NewIndex) = j
  'Next j
  'If (proj.Cases_Count > 0) Then
  '  ReDim SortIndex_CaseNames(1 To proj.Cases_Count)
  '  For i = 1 To proj.Cases_Count
  '    SortIndex_CaseNames(i) = frmMain.lsttempsorter.ItemData(i - 1)
  '  Next i
  'End If
  
  If (frmPrint_DO_INPUTS) Then
    'PRINT THE INPUTS.
    f1.SheetName(SheetIdx_Inputs) = "Inputs"
    Call Print_Inputs(f1, proj, SheetIdx_Inputs)
  End If
  If (frmPrint_DO_OUTPUTS) Then
    For i = 1 To NumSheets_Outputs
      name_out = "error"
      Select Case proj.idreact
        Case IDREACT_CMBR: name_out = "Batch"
        Case IDREACT_CMFR: name_out = "Tank " & Trim$(Str$(i))
      End Select
      SheetIdx_ThisOutput = SheetIdx_Outputs + i - 1
      f1.SheetName(SheetIdx_ThisOutput) = name_out
      Call Print_Outputs(f1, proj, SheetIdx_ThisOutput, i)
    Next i
  End If
  If (frmPrint_DO_PLOTS) Then
    'DO NOTHING.
  End If
    
  'For i = 1 To proj.Cases_Count
  '  Case_Index = SortIndex_CaseNames(i)
  '  Sheet_Index = i
  '  f1.SheetName(Sheet_Index) = proj.Cases(Case_Index).name
  '  Call Print_Case(f1, proj, Case_Index, Sheet_Index)
  'Next i

  'RETURN TO SHEET #1.
  f1.Sheet = 1

End Sub

