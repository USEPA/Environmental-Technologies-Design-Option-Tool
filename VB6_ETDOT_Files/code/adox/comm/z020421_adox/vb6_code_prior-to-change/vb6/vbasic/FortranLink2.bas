Attribute VB_Name = "FortranLink2"
Option Explicit

Const DEPROTONATED_POSTFIX = "(-)"






Const FortranLink2_declarations_end = 0


Function qstr(v As Variant) As String
  qstr = NumberToMFBString(v)
End Function



Function FortranLink_LookupCompound(n As String) As Integer
Dim i As Integer
Dim found As Integer
  found = False
  For i = 1 To NowProj.ncomp
    If (Trim$(UCase$(n)) = Trim$(UCase$(NowProj.Fortran_Comp(i).comname))) Then
      found = True
      Exit For
    End If
  Next i
  If (found) Then
    FortranLink_LookupCompound = i
  Else
    FortranLink_LookupCompound = 0
  End If
End Function


Sub FortranLink_DoCompound(ByRef fcnt As Fortran_CompName_Type, n As String)
  fcnt.idx = FortranLink_LookupCompound(n)
  fcnt.name = n
End Sub


Sub FortranLink_DoTarget(ByRef fcnt As Fortran_CompName_Type, tnum As Integer, is_deprot As Integer)
Dim n As String
  n = NowProj.TargetCompounds(tnum + 1).comname
  If (is_deprot) Then
    n = n & DEPROTONATED_POSTFIX
  End If
  Call FortranLink_DoCompound(fcnt, n)
End Sub


Sub FortranLink_WriteInputFile()
Dim f As Integer
Dim fn_InputFile As String
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim idx As Integer
Dim fc As Fortran_Comp_Type
Dim fi As Fortran_IrrRxn_Type
Dim fr As Fortran_RevRxn_Type
Dim fp As Fortran_PhotRxn_Type
Dim fc2 As Fortran_Comp2_Type

Dim units As String
Dim qq As String

  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  'CALCULATE SOME SIZES.
  NowProj.ntarget = NowProj.TargetCompounds_Count - 1 'do not include NOM
  NowProj.nmultiacid = 2   'NowProj.ntarget + 6
  NowProj.NUM_REV = NowProj.ntarget + 3 * NowProj.nmultiacid  'NowProj.ntarget + 6
  NowProj.nphot = NowProj.ntarget + 2
  NowProj.nirrev = (NowProj.ntarget * 6) + 20
  NowProj.ncomp = (NowProj.ntarget * 2) + 14
  NowProj.nwvlen = NowProj.Wavelength_Count
  
  
  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  'CALCULATE COMPOUND ARRAY -- NowProj.Fortran_Comp().
  ReDim NowProj.Fortran_Comp(1 To NowProj.ncomp)
  'FIRST COMPOUND.
  fc.comname = "H2O2"
  fc.concini = NowProj.inf_h2o2
  fc.val = 0#
  fc.mw = 34#
  NowProj.Fortran_Comp(1) = fc
  idx = 2
  'TARGET COMPOUNDS.
  For i = 1 To NowProj.ntarget
    j = i + 1     'skip NOM (index 1) which is defined below.
    fc.comname = NowProj.TargetCompounds(j).comname
    fc.concini = NowProj.TargetCompounds(j).concini
    fc.val = NowProj.TargetCompounds(j).val
    fc.mw = NowProj.TargetCompounds(j).mw
    NowProj.Fortran_Comp(idx + i - 1) = fc
  Next i
  idx = 1 + NowProj.ntarget + 1
  'ADDITIONAL COMPOUNDS.
  fc.comname = "HO2*"
  fc.concini = 0#
  fc.val = 0#
  fc.mw = 33#
  NowProj.Fortran_Comp(idx + 0) = fc
  fc.comname = "HCO3-"
  fc.concini = 0#      'actual value set in FORTRAN.
  fc.val = -1#
  fc.mw = 61#
  NowProj.Fortran_Comp(idx + 1) = fc
  fc.comname = "H2PO4-"
  fc.concini = 0#      'actual value set in FORTRAN.
  fc.val = -1#
  fc.mw = 61#
  NowProj.Fortran_Comp(idx + 2) = fc
  fc.comname = "NOM"
  fc.concini = NowProj.TargetCompounds(1).concini 'note mg/L units.
  fc.val = 0#
  fc.mw = NowProj.TargetCompounds(1).mw
  NowProj.Fortran_Comp(idx + 3) = fc
  fc.comname = "CO3*-"
  fc.concini = 0#
  fc.val = -1#
  fc.mw = 62#
  NowProj.Fortran_Comp(idx + 4) = fc
  fc.comname = "HPO4*-"
  fc.concini = 0#
  fc.val = -1#
  fc.mw = 61#
  NowProj.Fortran_Comp(idx + 5) = fc
  fc.comname = "HO*"
  fc.concini = 0#
  fc.val = 0#
  fc.mw = 17#
  NowProj.Fortran_Comp(idx + 6) = fc
  fc.comname = "HO2-"
  fc.concini = 0#
  fc.val = -1#
  fc.mw = 33#
  NowProj.Fortran_Comp(idx + 7) = fc
  idx = idx + 8
  'DEPROTONATED TARGET COMPOUNDS.
  For i = 1 To NowProj.ntarget
    j = i + 1     'skip NOM (index 1); no deprotonated form.
    fc.comname = NowProj.TargetCompounds(j).comname & DEPROTONATED_POSTFIX
    fc.concini = 0#
    fc.val = -1#
    fc.mw = NowProj.TargetCompounds(j).mw - 1#
    NowProj.Fortran_Comp(idx + i - 1) = fc
  Next i
  idx = idx + NowProj.ntarget
  'THE REST OF THE COMPOUNDS.
  fc.comname = "O2*-"
  fc.concini = 0#
  fc.val = -1#
  fc.mw = 32#
  NowProj.Fortran_Comp(idx + 0) = fc
  fc.comname = "CO3--"
  fc.concini = 0#
  fc.val = -2#
  fc.mw = 60#
  NowProj.Fortran_Comp(idx + 1) = fc
  fc.comname = "HPO4--"
  fc.concini = 0#
  fc.val = -2#
  fc.mw = 60#
  NowProj.Fortran_Comp(idx + 2) = fc
  fc.comname = "H2CO3"
  fc.concini = 0#
  fc.val = -2#
  fc.mw = 62#
  NowProj.Fortran_Comp(idx + 3) = fc
  fc.comname = "H3PO4"
  fc.concini = 0#
  fc.val = -2#
  fc.mw = 60#
  NowProj.Fortran_Comp(idx + 4) = fc


  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  'CALCULATE IRREVERSIBLE REACTION ARRAY -- NowProj.Fortran_IrrRxn().
  ReDim NowProj.Fortran_IrrRxn(1 To NowProj.nirrev)
  'FIRST SET OF IRREVERSIBLE REACTIONS.
  idx = 1
  Call FortranLink_DoCompound(fi.compa, "H2O2")
  Call FortranLink_DoCompound(fi.compb, "HO*")
  Call FortranLink_DoCompound(fi.compc, "HO2*")
  Call FortranLink_DoCompound(fi.compd, "H2O (not considered in this model)")
  fi.xk = 2.7 * (10# ^ 7#)
  NowProj.Fortran_IrrRxn(idx + 0) = fi
  Call FortranLink_DoCompound(fi.compa, "HO*")
  Call FortranLink_DoCompound(fi.compb, "HO2-")
  Call FortranLink_DoCompound(fi.compc, "HO2*")
  Call FortranLink_DoCompound(fi.compd, "OH- (not considered in this model)")
  fi.xk = 7.5 * (10# ^ 9#)
  NowProj.Fortran_IrrRxn(idx + 1) = fi
  Call FortranLink_DoCompound(fi.compa, "H2O2")
  Call FortranLink_DoCompound(fi.compb, "HO2*")
  Call FortranLink_DoCompound(fi.compc, "HO*")
  Call FortranLink_DoCompound(fi.compd, "H2O + O2 (not considered in this model)")
  fi.xk = 3# * (10# ^ 0#)
  NowProj.Fortran_IrrRxn(idx + 2) = fi
  Call FortranLink_DoCompound(fi.compa, "H2O2")
  Call FortranLink_DoCompound(fi.compb, "O2*-")
  Call FortranLink_DoCompound(fi.compc, "HO*")
  Call FortranLink_DoCompound(fi.compd, "OH- + O2 (not considered in this model)")
  fi.xk = 0.13 * (10# ^ 0#)
  NowProj.Fortran_IrrRxn(idx + 3) = fi
  Call FortranLink_DoCompound(fi.compa, "HO*")
  Call FortranLink_DoCompound(fi.compb, "CO3--")
  Call FortranLink_DoCompound(fi.compc, "CO3*-")
  Call FortranLink_DoCompound(fi.compd, "OH- (not considered in this model)")
  fi.xk = 3.9 * (10# ^ 8#)
  NowProj.Fortran_IrrRxn(idx + 4) = fi
  Call FortranLink_DoCompound(fi.compa, "HO*")
  Call FortranLink_DoCompound(fi.compb, "HCO3-")
  Call FortranLink_DoCompound(fi.compc, "CO3*-")
  Call FortranLink_DoCompound(fi.compd, "H2O (not considered in this model)")
  fi.xk = 8.5 * (10# ^ 6#)
  NowProj.Fortran_IrrRxn(idx + 5) = fi
  Call FortranLink_DoCompound(fi.compa, "HO*")
  Call FortranLink_DoCompound(fi.compb, "HPO4--")
  Call FortranLink_DoCompound(fi.compc, "HPO4*-")
  Call FortranLink_DoCompound(fi.compd, "OH- (not considered in this model)")
  fi.xk = 1.5 * (10# ^ 5#)
  NowProj.Fortran_IrrRxn(idx + 6) = fi
  Call FortranLink_DoCompound(fi.compa, "HO*")
  Call FortranLink_DoCompound(fi.compb, "H2PO4-")
  Call FortranLink_DoCompound(fi.compc, "HPO4*-")
  Call FortranLink_DoCompound(fi.compd, "H2O (not considered in this model)")
  fi.xk = 2# * (10# ^ 4#)
  NowProj.Fortran_IrrRxn(idx + 7) = fi
  Call FortranLink_DoCompound(fi.compa, "H2O2")
  Call FortranLink_DoCompound(fi.compb, "CO3*-")
  Call FortranLink_DoCompound(fi.compc, "HCO3-")
  Call FortranLink_DoCompound(fi.compd, "HO2*")
  fi.xk = 4.3 * (10# ^ 5#)
  NowProj.Fortran_IrrRxn(idx + 8) = fi
  Call FortranLink_DoCompound(fi.compa, "HO2-")
  Call FortranLink_DoCompound(fi.compb, "CO3*-")
  Call FortranLink_DoCompound(fi.compc, "CO3--")
  Call FortranLink_DoCompound(fi.compd, "HO2*")
  fi.xk = 3# * (10# ^ 7#)
  NowProj.Fortran_IrrRxn(idx + 9) = fi
  Call FortranLink_DoCompound(fi.compa, "H2O2")
  Call FortranLink_DoCompound(fi.compb, "HPO4*-")
  Call FortranLink_DoCompound(fi.compc, "H2PO4-")
  Call FortranLink_DoCompound(fi.compd, "HO2*")
  fi.xk = 2.7 * (10# ^ 7#)
  NowProj.Fortran_IrrRxn(idx + 10) = fi
  Call FortranLink_DoCompound(fi.compa, "HO*")
  Call FortranLink_DoCompound(fi.compb, "HO*")
  Call FortranLink_DoCompound(fi.compc, "H2O2")
  Call FortranLink_DoCompound(fi.compd, "(no second product produced for this reaction)")
  fi.xk = 5.5 * (10# ^ 9#)
  NowProj.Fortran_IrrRxn(idx + 11) = fi
  Call FortranLink_DoCompound(fi.compa, "HO*")
  Call FortranLink_DoCompound(fi.compb, "HO2*")
  Call FortranLink_DoCompound(fi.compc, "H2O (not considered in this model)")
  Call FortranLink_DoCompound(fi.compd, "O2 (not considered in this model)")
  fi.xk = 6.6 * (10# ^ 9#)
  NowProj.Fortran_IrrRxn(idx + 12) = fi
  Call FortranLink_DoCompound(fi.compa, "HO2*")
  Call FortranLink_DoCompound(fi.compb, "HO2*")
  Call FortranLink_DoCompound(fi.compc, "H2O2")
  Call FortranLink_DoCompound(fi.compd, "O2 (not considered in this model)")
  fi.xk = 8.3 * (10# ^ 5#)
  NowProj.Fortran_IrrRxn(idx + 13) = fi
  Call FortranLink_DoCompound(fi.compa, "HO2*")
  Call FortranLink_DoCompound(fi.compb, "O2*-")
  Call FortranLink_DoCompound(fi.compc, "HO2-")
  Call FortranLink_DoCompound(fi.compd, "O2 (not considered in this model)")
  fi.xk = 9.7 * (10# ^ 7#)
  NowProj.Fortran_IrrRxn(idx + 14) = fi
  Call FortranLink_DoCompound(fi.compa, "HO*")
  Call FortranLink_DoCompound(fi.compb, "O2*-")
  Call FortranLink_DoCompound(fi.compc, "O2 (not considered in this model)")
  Call FortranLink_DoCompound(fi.compd, "OH- (not considered in this model)")
  fi.xk = 7# * (10# ^ 9#)
  NowProj.Fortran_IrrRxn(idx + 15) = fi
  Call FortranLink_DoCompound(fi.compa, "HO*")
  Call FortranLink_DoCompound(fi.compb, "CO3*-")
  Call FortranLink_DoCompound(fi.compc, "unknown (not considered in this model)")
  Call FortranLink_DoCompound(fi.compd, "unknown (not considered in this model)")
  fi.xk = 3# * (10# ^ 9#)
  NowProj.Fortran_IrrRxn(idx + 16) = fi
  Call FortranLink_DoCompound(fi.compa, "CO3*-")
  Call FortranLink_DoCompound(fi.compb, "O2*-")
  Call FortranLink_DoCompound(fi.compc, "CO3--")
  Call FortranLink_DoCompound(fi.compd, "O2 (not considered in this model)")
  fi.xk = 6# * (10# ^ 8#)
  NowProj.Fortran_IrrRxn(idx + 17) = fi
  Call FortranLink_DoCompound(fi.compa, "CO3*-")
  Call FortranLink_DoCompound(fi.compb, "CO3*-")
  Call FortranLink_DoCompound(fi.compc, "unknown (not considered in this model)")
  Call FortranLink_DoCompound(fi.compd, "unknown (not considered in this model)")
  fi.xk = 3# * (10# ^ 7#)
  NowProj.Fortran_IrrRxn(idx + 18) = fi
  idx = idx + 19
  'REACTIONS OF TARGET COMPOUNDS WITH HO*.
  For i = 1 To NowProj.ntarget
    Call FortranLink_DoCompound(fi.compa, "HO*")
    Call FortranLink_DoTarget(fi.compb, i, False)   'PROTONATED FORM.
    Call FortranLink_DoCompound(fi.compc, "unknown (not considered in this model)")
    Call FortranLink_DoCompound(fi.compd, "unknown (not considered in this model)")
    j = i + 1     'skip NOM (index 1).
    fi.xk = NowProj.TargetCompounds(j).xk
    NowProj.Fortran_IrrRxn(idx + i - 1) = fi
  Next i
  idx = idx + NowProj.ntarget
  'REACTIONS OF DEPROTONATED TARGET COMPOUNDS WITH HO*.
  For i = 1 To NowProj.ntarget
    Call FortranLink_DoCompound(fi.compa, "HO*")
    Call FortranLink_DoTarget(fi.compb, i, True)   'DEPROTONATED FORM.
    Call FortranLink_DoCompound(fi.compc, "unknown (not considered in this model)")
    Call FortranLink_DoCompound(fi.compd, "unknown (not considered in this model)")
    j = i + 1     'skip NOM (index 1).
    fi.xk = NowProj.TargetCompounds(j).dep_xk
    NowProj.Fortran_IrrRxn(idx + i - 1) = fi
  Next i
  idx = idx + NowProj.ntarget
  'REACTION OF NOM WITH HO*.
  Call FortranLink_DoCompound(fi.compa, "HO*")
  Call FortranLink_DoCompound(fi.compb, "NOM")
  Call FortranLink_DoCompound(fi.compc, "unknown (not considered in this model)")
  Call FortranLink_DoCompound(fi.compd, "unknown (not considered in this model)")
  fi.xk = NowProj.TargetCompounds(1).xk
  NowProj.Fortran_IrrRxn(idx) = fi
  idx = idx + 1
  'REACTIONS OF TARGET COMPOUNDS WITH CO3*-.
  For i = 1 To NowProj.ntarget
    Call FortranLink_DoCompound(fi.compa, "CO3*-")
    Call FortranLink_DoTarget(fi.compb, i, False)   'PROTONATED FORM.
    Call FortranLink_DoCompound(fi.compc, "unknown (not considered in this model)")
    Call FortranLink_DoCompound(fi.compd, "unknown (not considered in this model)")
    j = i + 1     'skip NOM (index 1).
    fi.xk = NowProj.TargetCompounds(j).xk_co3XM
    NowProj.Fortran_IrrRxn(idx + i - 1) = fi
  Next i
  idx = idx + NowProj.ntarget
  'REACTIONS OF TARGET COMPOUNDS WITH HPO4*-.
  For i = 1 To NowProj.ntarget
    Call FortranLink_DoCompound(fi.compa, "HPO4*-")
    Call FortranLink_DoTarget(fi.compb, i, False)   'PROTONATED FORM.
    Call FortranLink_DoCompound(fi.compc, "unknown (not considered in this model)")
    Call FortranLink_DoCompound(fi.compd, "unknown (not considered in this model)")
    j = i + 1     'skip NOM (index 1).
    fi.xk = NowProj.TargetCompounds(j).xk_hpo4XM
    NowProj.Fortran_IrrRxn(idx + i - 1) = fi
  Next i
  idx = idx + NowProj.ntarget
  'REACTIONS OF TARGET COMPOUNDS WITH O2*-.
  For i = 1 To NowProj.ntarget
    Call FortranLink_DoCompound(fi.compa, "O2*-")
    Call FortranLink_DoTarget(fi.compb, i, False)   'PROTONATED FORM.
    Call FortranLink_DoCompound(fi.compc, "unknown (not considered in this model)")
    Call FortranLink_DoCompound(fi.compd, "unknown (not considered in this model)")
    j = i + 1     'skip NOM (index 1).
    fi.xk = NowProj.TargetCompounds(j).xk_o2XM
    NowProj.Fortran_IrrRxn(idx + i - 1) = fi
  Next i
  idx = idx + NowProj.ntarget
  'REACTIONS OF TARGET COMPOUNDS WITH HO2*.
  For i = 1 To NowProj.ntarget
    Call FortranLink_DoCompound(fi.compa, "HO2*")
    Call FortranLink_DoTarget(fi.compb, i, False)   'PROTONATED FORM.
    Call FortranLink_DoCompound(fi.compc, "unknown (not considered in this model)")
    Call FortranLink_DoCompound(fi.compd, "unknown (not considered in this model)")
    j = i + 1     'skip NOM (index 1).
    fi.xk = NowProj.TargetCompounds(j).xk_ho2X
    NowProj.Fortran_IrrRxn(idx + i - 1) = fi
  Next i
  idx = idx + NowProj.ntarget
  
  
  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  'CALCULATE REVERSIBLE REACTION ARRAY -- NowProj.Fortran_RevRxn().
  ReDim NowProj.Fortran_RevRxn(1 To NowProj.NUM_REV)
  'FIRST REVERSIBLE REACTION.
  idx = 1
  Call FortranLink_DoCompound(fr.compe, "H2O2")
  Call FortranLink_DoCompound(fr.compf, "HO2-")
  fr.xke = 11.6
  NowProj.Fortran_RevRxn(idx + 0) = fr
  idx = 2
  'REVERSIBLE DEPROTONATION REACTIONS OF TARGET COMPOUNDS.
  For i = 1 To NowProj.ntarget
    Call FortranLink_DoTarget(fr.compe, i, False)   'PROTONATED FORM.
    Call FortranLink_DoTarget(fr.compf, i, True)    'DEPROTONATED FORM.
    j = i + 1     'skip NOM (index 1).
    fr.xke = NowProj.TargetCompounds(j).dep_xke
    NowProj.Fortran_RevRxn(idx + i - 1) = fr
  Next i
  idx = idx + NowProj.ntarget
  'THE REST OF THE REVERSIBLE REACTIONS.
  Call FortranLink_DoCompound(fr.compe, "HO2*")
  Call FortranLink_DoCompound(fr.compf, "O2*-")
  fr.xke = 4.8
  NowProj.Fortran_RevRxn(idx + 0) = fr
  Call FortranLink_DoCompound(fr.compe, "HCO3-")
  Call FortranLink_DoCompound(fr.compf, "CO3--")
  fr.xke = 10.3
  NowProj.Fortran_RevRxn(idx + 1) = fr
  Call FortranLink_DoCompound(fr.compe, "H2PO4-")
  Call FortranLink_DoCompound(fr.compf, "HPO4--")
  fr.xke = 7.2
  NowProj.Fortran_RevRxn(idx + 2) = fr
  Call FortranLink_DoCompound(fr.compe, "H2CO3")
  Call FortranLink_DoCompound(fr.compf, "HCO3-")
  fr.xke = 6.3
  NowProj.Fortran_RevRxn(idx + 3) = fr
  Call FortranLink_DoCompound(fr.compe, "H3PO4")
  Call FortranLink_DoCompound(fr.compf, "H2PO4-")
  fr.xke = 2.1
  NowProj.Fortran_RevRxn(idx + 4) = fr
  
  
  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  'CALCULATE PHOTOLYSIS REACTION ARRAY -- NowProj.Fortran_PhotRxn().
  ReDim NowProj.Fortran_PhotRxn(1 To NowProj.nphot)
  'FIRST PHOTOLYSIS REACTION.
  idx = 1
  Call FortranLink_DoCompound(fp.compg, "H2O2")
  Call FortranLink_DoCompound(fp.comph, "HO*")
  fp.stocphot = 2#
  ReDim fp.extcoef(1 To NowProj.nwvlen)
  For i = 1 To NowProj.nwvlen
    fp.extcoef(i) = EXTCOEF_DEFAULT_VALUE   '' 19#  Changed by KAM per Luke
  Next i
  ReDim fp.quatyd(1 To NowProj.nwvlen)
  For i = 1 To NowProj.nwvlen
    fp.quatyd(i) = QUATYD_DEFAULT_VALUE     '' 0.5  Changed by KAM per Luke
  Next i
  NowProj.Fortran_PhotRxn(idx + 0) = fp
  idx = 2
  'PHOTOLYSIS REACTIONS FOR TARGET COMPOUNDS.
  For i = 1 To NowProj.ntarget
    Call FortranLink_DoTarget(fp.compg, i, False)   'PROTONATED FORM.
    Call FortranLink_DoCompound(fp.comph, "unknown (not considered in this model)")
    j = i + 1     'skip NOM (index 1).
    fp.stocphot = 1#
    ReDim fp.extcoef(1 To NowProj.nwvlen)
    For k = 1 To NowProj.nwvlen
      fp.extcoef(k) = NowProj.extcoef(j, k)
    Next k
    ReDim fp.quatyd(1 To NowProj.nwvlen)
    For k = 1 To NowProj.nwvlen
      fp.quatyd(k) = NowProj.quatyd(j, k)
    Next k
    NowProj.Fortran_PhotRxn(idx + i - 1) = fp
  Next i
  idx = idx + NowProj.ntarget
  'THE FINAL PHOTOLYSIS REACTION.
  Call FortranLink_DoCompound(fp.compg, "NOM")
  Call FortranLink_DoCompound(fp.comph, "unknown (not considered in this model)")
  fp.stocphot = 1#
  ReDim fp.extcoef(1 To NowProj.nwvlen)
  For i = 1 To NowProj.nwvlen
    fp.extcoef(i) = EXTCOEF_DEFAULT_VALUE   ''0.0867  Changed by KAM per Luke
  Next i
  ReDim fp.quatyd(1 To NowProj.nwvlen)
  For i = 1 To NowProj.nwvlen
    fp.quatyd(i) = QUATYD_DEFAULT_VALUE     ''0# changed by KAM per Luke
  Next i
  NowProj.Fortran_PhotRxn(idx + 0) = fp
  idx = idx + 1
  
  
  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  'CALCULATE COMPOUND2 REACTION ARRAY -- NowProj.Fortran_Comp2().
  ReDim NowProj.Fortran_Comp2(1 To NowProj.ntarget)
  idx = 1
  For i = 1 To NowProj.ntarget
    j = i + 1     'skip NOM (index 1).
    fc2.ncarbn = NowProj.TargetCompounds(j).ncarbn
    fc2.nsubstt = NowProj.TargetCompounds(j).nsubstt
    NowProj.Fortran_Comp2(idx + i - 1) = fc2
  Next i
  
  
  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  'WRITE THE FORTRAN INPUT FILE.
  fn_InputFile = App.Path & "\exes\adox_in.txt"
  f = FreeFile
  Open fn_InputFile For Output As #f
  '---- MAIN INPUTS --------------------------------------------------------------------------------------------
  Call WriteFortranInput(f, NowProj.ntarget, "NTARGET, number of target organic compounds")
  Call WriteFortranInput(f, 0.0000000001, "EPS, DGEAR error criteria")
  Call WriteFortranInput(f, NowProj.idreact, "IDREACT, 0=CMBR or 1=CMFR")
  'Call WriteFortranInput(f, "1", "NTANK, tank number (not implemented yet)")
  Call WriteFortranInput(f, NowProj.num_tanks, "NTANK, number of tanks")
  Call WriteFortranInput(f, NowProj.volume, "VOLUME, tank volume, liters")
  Call WriteFortranInput(f, NowProj.ssize, "SSIZE, simulation time step, sec")
  Call WriteFortranInput(f, NowProj.ttotal, "TTOTAL, total simulation time (used for CMBR), min")
  Call WriteFortranInput(f, NowProj.xntimes, "XNTIMES, multiples of total hydraulic retention time (used for CMBR), dim'less")
  Call WriteFortranInput(f, NowProj.tau, "TAU, hydraulic retention time of a tank, min")
  Call WriteFortranInput(f, NowProj.opsize, "OPSIZE, time interval for output data, min")
  Call WriteFortranInput(f, NowProj.idcarbn, "IDCARBN, TIC input method, 0=[CO3], 1=alkalinity")
  Call WriteFortranInput(f, NowProj.alk, "ALK, alkalinity (if IDCARBN=1), mg/L as CaCO3")
  Call WriteFortranInput(f, NowProj.ticarbn, "TICARBN, total inorganic carbonate concentration [CO3], gmol/L")
  Call WriteFortranInput(f, NowProj.ph0, "PH(0), initial/influent pH value")
  Call WriteFortranInput(f, NowProj.phosph, "PHOSPH, total inorganic phosphate ion concentration, gmol/L")
  Call WriteFortranInput(f, NowProj.uvpathl, "UVPATHL, optical path length of UV light, cm")
  Call WriteFortranInput(f, NowProj.nwvlen, "NWVLEN, number of wavelength ranges specified by user")
  For i = 1 To NowProj.nwvlen
    Call WriteFortranInput(f, NowProj.Wavelengths(i).lwave, "LWAVE(" & Trim$(Str$(i)) & "), wavelength, nm")
  Next i
  Call WriteFortranInput(f, NowProj.lamp_power, "ELECTR_POWER, the lamp electrical power")
  Call WriteFortranInput(f, NowProj.iduvi, "IDUVI, how UV intensity (UVI) is input: 0=Eins./L-s, 1=Watts, 2=as efficiency")
  For i = 1 To NowProj.nwvlen
    Call WriteFortranInput(f, NowProj.Wavelengths(i).uvi, "UVI(" & Trim$(Str$(i)) & "), UV light parameter (refer to IDUVI for units)")
  Next i
  
  '---- COMPOUNDS --------------------------------------------------------------------------------------------
  Call WriteFortranInput(f, NowProj.ncomp, "NCOMP, number of compounds")
  qq = Chr$(34)
  For i = 1 To NowProj.ncomp
    units = "gmol/L"
    If (Trim$(UCase$(NowProj.Fortran_Comp(i).comname)) = Trim$(UCase$("NOM"))) Then
      units = "mg/L (note different units of mg/L)"
    End If
    Call WriteFortranInput(f, qq & NowProj.Fortran_Comp(i).comname & qq, "::COMNAME(" & qstr(i) & "), compound name")
    Call WriteFortranInput(f, NowProj.Fortran_Comp(i).concini, "  CONCINI(" & qstr(i) & "), initial/influent concentration, " & units)
    Call WriteFortranInput(f, NowProj.Fortran_Comp(i).val, "  VAL(" & qstr(i) & "), valence")
    Call WriteFortranInput(f, NowProj.Fortran_Comp(i).mw, "  MW(" & qstr(i) & "), molecular weight, g/gmol")
  Next i
  '---- IRREVERSIBLE REACTIONS --------------------------------------------------------------------------------------------
  Call WriteFortranInput(f, NowProj.nirrev, "NIRREV, number of irreversible reactions ( A + B  -->  C + D )")
  For i = 1 To NowProj.nirrev
    units = "1/(gmol/L)-s"
    If (Trim$(UCase$(NowProj.Fortran_IrrRxn(i).compa.name)) = Trim$(UCase$("NOM"))) Or _
       (Trim$(UCase$(NowProj.Fortran_IrrRxn(i).compb.name)) = Trim$(UCase$("NOM"))) Then
      units = "1/(mg/L)-s (note different units of 1/(mg/L)-s)"
    End If
    Call WriteFortranInput(f, NowProj.Fortran_IrrRxn(i).compa.idx, "::COMPA(" & qstr(i) & "), A for rxn#" & qstr(i) & " -- " & NowProj.Fortran_IrrRxn(i).compa.name)
    Call WriteFortranInput(f, NowProj.Fortran_IrrRxn(i).compb.idx, "  COMPB(" & qstr(i) & "), B for rxn#" & qstr(i) & " -- " & NowProj.Fortran_IrrRxn(i).compb.name)
    Call WriteFortranInput(f, NowProj.Fortran_IrrRxn(i).compc.idx, "  COMPC(" & qstr(i) & "), C for rxn#" & qstr(i) & " -- " & NowProj.Fortran_IrrRxn(i).compc.name)
    Call WriteFortranInput(f, NowProj.Fortran_IrrRxn(i).compd.idx, "  COMPD(" & qstr(i) & "), D for rxn#" & qstr(i) & " -- " & NowProj.Fortran_IrrRxn(i).compd.name)
    Call WriteFortranInput(f, NowProj.Fortran_IrrRxn(i).xk, "  XK(" & qstr(i) & "), 2nd order rate constant for rxn#" & qstr(i) & ", " & units)
  Next i
  '---- REVERSIBLE REACTIONS --------------------------------------------------------------------------------------------
  Call WriteFortranInput(f, NowProj.nmultiacid, "NMULTIACID, number of multiprotic acids ( E <==> F )")
  For i = 1 To NowProj.NUM_REV
    Call WriteFortranInput(f, NowProj.Fortran_RevRxn(i).compe.idx, "::COMPE(" & qstr(i) & "), E for rxn#" & qstr(i) & " -- " & NowProj.Fortran_RevRxn(i).compe.name)
    Call WriteFortranInput(f, NowProj.Fortran_RevRxn(i).compf.idx, "  COMPF(" & qstr(i) & "), F for rxn#" & qstr(i) & " -- " & NowProj.Fortran_RevRxn(i).compf.name)
    Call WriteFortranInput(f, NowProj.Fortran_RevRxn(i).xke, "  XKE(" & qstr(i) & "), equilibrium constant of reversible rxn#" & qstr(i))
  Next i
  '---- PHOTOLYSIS REACTIONS --------------------------------------------------------------------------------------------
  Call WriteFortranInput(f, NowProj.nphot, "NPHOT, number of photolysis reactions ( G --> h H )")
  For i = 1 To NowProj.nphot
    Call WriteFortranInput(f, NowProj.Fortran_PhotRxn(i).compg.idx, "::COMPG(" & qstr(i) & "), G for rxn#" & qstr(i) & " -- " & NowProj.Fortran_PhotRxn(i).compg.name)
    Call WriteFortranInput(f, NowProj.Fortran_PhotRxn(i).comph.idx, "  COMPH(" & qstr(i) & "), H for rxn#" & qstr(i) & " -- " & NowProj.Fortran_PhotRxn(i).comph.name)
    Call WriteFortranInput(f, NowProj.Fortran_PhotRxn(i).stocphot, "  STOCPHOT(" & qstr(i) & "), moles H produced when 1 mole G destroyed in rxn#" & qstr(i))
    For j = 1 To NowProj.nwvlen
      Call WriteFortranInput(f, NowProj.Fortran_PhotRxn(i).extcoef(j), _
          "  EXTCOEF(" & qstr(i) & "," & qstr(j) & "), extinction coefficient for rxn#" & _
          qstr(i) & " at " & qstr(NowProj.Wavelengths(j).lwave) & " nm")
    Next j
    For j = 1 To NowProj.nwvlen
      Call WriteFortranInput(f, NowProj.Fortran_PhotRxn(i).quatyd(j), _
          "  QUATYD(" & qstr(i) & "," & qstr(j) & "), quantum yield for rxn#" & _
          qstr(i) & " at " & qstr(NowProj.Wavelengths(j).lwave) & " nm")
    Next j
  Next i
  '---- COMPOUNDS #2 --------------------------------------------------------------------------------------------
  For i = 1 To NowProj.ntarget
    Call WriteFortranInput(f, NowProj.Fortran_Comp2(i).ncarbn, "::NCARBN(" & qstr(i) & "), number of carbon atoms per molecule of compound " & NowProj.TargetCompounds(i + 1).comname)
    Call WriteFortranInput(f, NowProj.Fortran_Comp2(i).nsubstt, "  NSUBSTT(" & qstr(i) & "), number of hydrogen substituted atoms (e.g. Cl,Br,etc.) per molecule of compound " & NowProj.TargetCompounds(i + 1).comname)
  Next i
  
  Close #f
  

End Sub

