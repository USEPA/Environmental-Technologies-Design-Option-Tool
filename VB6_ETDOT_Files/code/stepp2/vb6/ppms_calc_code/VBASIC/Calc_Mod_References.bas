Attribute VB_Name = "Calc_Mod_References"
Option Explicit




Const Calc_Mod_References_decl_end = True


Function Calc_Mod_GetRefText( _
    inout_TechDat As TechniqueData_Type) _
    As Boolean
On Error GoTo err_ThisFunc
Dim Ret As String
Dim RefText As String
  Ret = vbCrLf
  RefText = ""
  '
  ' HANDLE FIRST SET OF TECHNIQUE CODES.
  '
  Select Case inout_TechDat.Technique_Code
    ''''Case "MTU Fire & Explosion Data"
    Case TECHCODE_023_026d_MTU_FIREEXP_DATA, _
        TECHCODE_024_030d_MTU_FIREEXP_DATA, _
        TECHCODE_025_035d_MTU_FIREEXP_DATA, _
        TECHCODE_026_040d_MTU_FIREEXP_DATA:
      RefText = "Source:" & Ret & Ret
      RefText = RefText & "Pintar, A.J., Lukowski, J.S., Penke, B., Prezkop, T.M., Surd," & Ret
      RefText = RefText & "S.M., The Use of Estimation Methods to Screen Fire and" & Ret
      RefText = RefText & "Explosion Data, Technical Support Document, Project 912" & Ret
      RefText = RefText & "Sponsor Release, July 1996, Design Institute for Physical" & Ret
      RefText = RefText & "Property Data, AIChE, New York, NY." & Ret
      GoTo exit_normally_ThisFunc
    ''''Case "MTU Linear Method"
    Case TECHCODE_026_042d_MTU_LINEAR_METHOD:
      RefText = "Source:" & Ret & Ret
      RefText = RefText & "Pintar, A.J., Estimation of Autoignition Temperature," & Ret
      RefText = RefText & "Technical Support Document, Project 912 Sponsor Release," & Ret
      RefText = RefText & "July 1996, Design Institute for Physical Property Data," & Ret
      RefText = RefText & "AIChE, New York, NY." & Ret
      GoTo exit_normally_ThisFunc
    ''''Case "MTU Logarithmic Method"
    Case TECHCODE_026_041d_MTU_LOG_METHOD:
      RefText = "Source:" & Ret & Ret
      RefText = RefText & "Pintar, A.J., Estimation of Autoignition Temperature," & Ret
      RefText = RefText & "Technical Support Document, Project 912 Sponsor Release," & Ret
      RefText = RefText & "July 1996, Design Institute for Physical Property Data," & Ret
      RefText = RefText & "AIChE, New York, NY." & Ret
      GoTo exit_normally_ThisFunc
    ''''Case "MTU Group Contribution"
    Case TECHCODE_023_027d_MTU_GROUP_CONTRIB, _
        TECHCODE_024_031d_MTU_GROUP_CONTRIB
      RefText = "Source:" & Ret & Ret
      RefText = RefText & "Pintar, A.J., Lukowski, J.S., Penke, B., Prezkop, T.M., Surd," & Ret
      RefText = RefText & "S.M., The Use of Estimation Methods to Screen Fire and" & Ret
      RefText = RefText & "Explosion Data, Technical Support Document, Project 912" & Ret
      RefText = RefText & "Sponsor Release, July 1996, Design Institute for Physical" & Ret
      RefText = RefText & "Property Data, AIChE, New York, NY." & Ret
      GoTo exit_normally_ThisFunc
    ''''Case "Penn State Group Contribu"
    Case TECHCODE_023_029d_PENN_GROUP_CONTRIB, _
        TECHCODE_024_032d_PENN_GROUP_CONTRIB, _
        TECHCODE_025_038d_PENN_GROUP_CONTRIB:
      RefText = "Source:" & Ret & Ret
      RefText = RefText & "Danner, R.P. and T.E. Daubert, " & Chr(34) & "Manual for Predicting Chemical" & Ret
      RefText = RefText & "Process Design Data," & Chr(34) & " Design Instutefor Physical Property Data," & Ret
      RefText = RefText & "AIChE, 1986" & Ret
      GoTo exit_normally_ThisFunc
    ''''Case "MTU LFL Group Contributio"
    Case TECHCODE_025_037d_MTU_LFL_GROUP_CONTRIB:
      RefText = "Source:" & Ret & Ret
      RefText = RefText & "Pintar, A.J., Estimation of FlashPoint, Tehcnical Support" & Ret
      RefText = RefText & "Document, Project 912 Sponsor Release, June 1996, Design" & Ret
      RefText = RefText & "Institute for Physical Property Data, AIChE, New York, NY." & Ret
      GoTo exit_normally_ThisFunc
    ''''Case "LFL Data"
    Case TECHCODE_025_036d_LFL_DATA:
      RefText = "Source:" & Ret & Ret
      RefText = RefText & "Pintar, A.J., Lukowski, J.S., Penke, B., Prezkop, T.M., Surd," & Ret
      RefText = RefText & "S.M., The Use of Estimation Methods to Screen Fire and" & Ret
      RefText = RefText & "Explosion Data, Technical Support Document, Project 912" & Ret
      RefText = RefText & "Sponsor Release, July 1996, Design Institute for Physical" & Ret
      RefText = RefText & "Property Data, AIChE, New York, NY." & Ret
      GoTo exit_normally_ThisFunc
    ''''Case "MTU FlashPoint Method"
    Case TECHCODE_024_034d_MTU_FLASHPOINT_METH:
      RefText = "Source:" & Ret & Ret
      RefText = RefText & "Pintar, A.J., Estimation of Lower Flammable Limit, Technical" & Ret
      RefText = RefText & "Support Document, Project 912 Sponsor Release, July 1996," & Ret
      RefText = RefText & "Design Institute for Physical Property Data, AIChE, New" & Ret
      RefText = RefText & "York, NY." & Ret
      GoTo exit_normally_ThisFunc
    ''''Case "MTU Combustion Reaction"
    Case TECHCODE_023_028d_MTU_COMBUSTION_RXN, _
        TECHCODE_024_033d_MTU_COMBUSTION_RXN
      RefText = "Source:" & Ret & Ret
      RefText = RefText & "Pintar, A.J., Estimation of Lower Flammable Limit, Technical" & Ret
      RefText = RefText & "Support Document, Project 912 Sponsor Release, July 1996," & Ret
      RefText = RefText & "Design Institute for Physical Property Data, AIChE, New" & Ret
      RefText = RefText & "York, NY." & Ret
      GoTo exit_normally_ThisFunc
    ''''Case "MTU LFL Combustion Reacti"
    Case TECHCODE_025_039d_MTU_LFL_COMBUSTION_RXN:
      RefText = "Source:" & Ret & Ret
      RefText = RefText & "Pintar, A.J., Estimation of FlashPoint, Tehcnical Support" & Ret
      RefText = RefText & "Document, Project 912 Sponsor Release, June 1996, Design" & Ret
      RefText = RefText & "Institute for Physical Property Data, AIChE, New York, NY." & Ret
      GoTo exit_normally_ThisFunc
    ''''Case "Antoine"
    Case TECHCODE_006_007d_ANTOINELIKE_EXPRESSION:
      RefText = "Source:" & Ret & Ret
      RefText = RefText & "Boublik, T., Fried, V. Hala, E., The Vapour Pressures of Pure Substances" & Ret
      RefText = RefText & "Yaws, C.L., Thermodynamic and Physical Property Data" & Ret
      RefText = RefText & "Dean, J.A., Lange's Handbook of Chemistry, 4th Edition" & Ret
      RefText = RefText & "Stephenson, R.M., Malanowski, S., Handbook of the Thermodynamics of Organic Compounds" & Ret
      RefText = RefText & "Mallard, W.G., Linstrom, P.J., eds., NIST Standard Reference Database Number 69"
      GoTo exit_normally_ThisFunc
  End Select
  '
  ' HANDLE TEMPERATURE-DEPENDENT EQUATION REFERENCES.
  '
  Select Case inout_TechDat.FofT_EqForm
    Case 100:
      RefText = "Equation Form 100:" & Ret & Ret
      RefText = RefText & "A + BT + CT^2 + DT^3 + ET^4" & Ret & Ret
      RefText = RefText & "Where:" & Ret & Ret
      RefText = RefText & "T = Operating Temperature" & Ret & Ret
      RefText = RefText & "Source:" & Ret & Ret
      RefText = RefText & "Sibul, H.M., Stebbins, C., Kendall, R., Wang, Y., Daubert, N.C., and T.E. Daubert" & Ret
      RefText = RefText & "Documentation of Policies and Procedures for Compilation, Prediction, and Correlation for DIPPR Project 801" & Ret
      RefText = RefText & "Pennsylvania State University, University Park, PA, July 1994"
      'GoTo exit_normally_ThisFunc
    Case 101:
      RefText = "Equation Form 101:" & Ret & Ret
      RefText = RefText & "EXP(A + B/T + C(LN(T)) + DT^E)" & Ret & Ret
      RefText = RefText & "Where:" & Ret & Ret
      RefText = RefText & "T = Operating Temperature" & Ret & Ret
      RefText = RefText & "Source:" & Ret & Ret
      RefText = RefText & "Sibul, H.M., Stebbins, C., Kendall, R., Wang, Y., Daubert, N.C., and T.E. Daubert" & Ret
      RefText = RefText & "Documentation of Policies and Procedures for Compilation, Prediction, and Correlation for DIPPR Project 801" & Ret
      RefText = RefText & "Pennsylvania State University, University Park, PA, July 1994"
      'GoTo exit_normally_ThisFunc
    Case 102
      RefText = "Equation Form 102:" & Ret & Ret
      RefText = RefText & "AT^B/(1 + C/T + D/T^2)" & Ret & Ret
      RefText = RefText & "Where:" & Ret & Ret
      RefText = RefText & "T = Operating Temperature" & Ret & Ret
      RefText = RefText & "Source:" & Ret & Ret
      RefText = RefText & "Sibul, H.M., Stebbins, C., Kendall, R., Wang, Y., Daubert, N.C., and T.E. Daubert" & Ret
      RefText = RefText & "Documentation of Policies and Procedures for Compilation, Prediction, and Correlation for DIPPR Project 801" & Ret
      RefText = RefText & "Pennsylvania State University, University Park, PA, July 1994"
      'GoTo exit_normally_ThisFunc
    Case 105:
      RefText = "Equation Form 105:" & Ret & Ret
      RefText = RefText & "A/B^(1 + (1 - T/C)^D)" & Ret & Ret
      RefText = RefText & "Where:" & Ret & Ret
      RefText = RefText & "T = Operating Temperature" & Ret & Ret
      RefText = RefText & "Source:" & Ret & Ret
      RefText = RefText & "Sibul, H.M., Stebbins, C., Kendall, R., Wang, Y., Daubert, N.C., and T.E. Daubert" & Ret
      RefText = RefText & "Documentation of Policies and Procedures for Compilation, Prediction, and Correlation for DIPPR Project 801" & Ret
      RefText = RefText & "Pennsylvania State University, University Park, PA, July 1994"
      'GoTo exit_normally_ThisFunc
    Case 106:
      RefText = "Equation Form 106:" & Ret & Ret
      RefText = RefText & "EXP(A(1 - Tr)^(B + CTr + DTr^2 + ETr^3)" & Ret & Ret
      RefText = RefText & "Where:" & Ret & Ret
      RefText = RefText & "Tr = Reduced Temperature" & Ret & Ret
      RefText = RefText & "Source:" & Ret & Ret
      RefText = RefText & "Sibul, H.M., Stebbins, C., Kendall, R., Wang, Y., Daubert, N.C., and T.E. Daubert" & Ret
      RefText = RefText & "Documentation of Policies and Procedures for Compilation, Prediction, and Correlation for DIPPR Project 801" & Ret
      RefText = RefText & "Pennsylvania State University, University Park, PA, July 1994"
      'GoTo exit_normally_ThisFunc
    Case 107:
      RefText = "Equation Form 107:" & Ret & Ret
      RefText = RefText & "A + B[(C/T)/SINH(C/T)]^2 + D[(E/T)/COSH(E/T)]^2" & Ret & Ret
      RefText = RefText & "Where:" & Ret & Ret
      RefText = RefText & "T = Operating Temperature" & Ret & Ret
      RefText = RefText & "Source:" & Ret & Ret
      RefText = RefText & "Sibul, H.M., Stebbins, C., Kendall, R., Wang, Y., Daubert, N.C., and T.E. Daubert" & Ret
      RefText = RefText & "Documentation of Policies and Procedures for Compilation, Prediction, and Correlation for DIPPR Project 801" & Ret
      RefText = RefText & "Pennsylvania State University, University Park, PA, July 1994"
      'GoTo exit_normally_ThisFunc
    Case 114:
      RefText = "Equation Form 114:" & Ret & Ret
      RefText = RefText & "A/^2/t + B - 2ACt - ADt^2 - C^2t^3/3 - CDt^4/2 - D^2t^5/5" & Ret & Ret
      RefText = RefText & "Where:" & Ret & Ret
      RefText = RefText & "T = Operating Temperature" & Ret & Ret
      RefText = RefText & "Source:" & Ret & Ret
      RefText = RefText & "Stebbins, C., Kendall, R., Crane, N., Copella, C., Gibson, M., and T.E. Daubert" & Ret
      RefText = RefText & "Documentation of Policies and Procedures for Compilation, Prediction, and Correlation for DIPPR Project 801" & Ret
      RefText = RefText & "Pennsylvania State University, University Park, PA, October 1997"
      'GoTo exit_normally_ThisFunc
    Case 115:
      RefText = "Equation Form 115:" & Ret & Ret
      RefText = RefText & "A + B/T + C ln(T) + DT^2 + E/T^2" & Ret & Ret
      RefText = RefText & "Where:" & Ret & Ret
      RefText = RefText & "T = Operating Temperature" & Ret & Ret
      RefText = RefText & "Source:" & Ret & Ret
      RefText = RefText & "Stebbins, C., Kendall, R., Crane, N., Copella, C., Gibson, M., and T.E. Daubert" & Ret
      RefText = RefText & "Documentation of Policies and Procedures for Compilation, Prediction, and Correlation for DIPPR Project 801" & Ret
      RefText = RefText & "Pennsylvania State University, University Park, PA, October 1997"
      'GoTo exit_normally_ThisFunc
    Case 116:
      RefText = "Equation Form 116:" & Ret & Ret
      RefText = RefText & "A + B(1-T)^.35 +  D(1-T) + E(1-T)^4/3" & Ret & Ret
      RefText = RefText & "Where:" & Ret & Ret
      RefText = RefText & "T = Operating Temperature" & Ret & Ret
      RefText = RefText & "Source:" & Ret & Ret
      RefText = RefText & "Stebbins, C., Kendall, R., Crane, N., Copella, C., Gibson, M., and T.E. Daubert" & Ret
      RefText = RefText & "Documentation of Policies and Procedures for Compilation, Prediction, and Correlation for DIPPR Project 801" & Ret
      RefText = RefText & "Pennsylvania State University, University Park, PA, October 1997"
      'GoTo exit_normally_ThisFunc
    Case 200:
      RefText = "Equation Form 200:" & Ret & Ret
      RefText = RefText & "A + BT + CT^2(LN(T)) + DT^2.5 + ET^3" & Ret & Ret
      RefText = RefText & "Where:" & Ret & Ret
      RefText = RefText & "T = Operating Temperature" & Ret & Ret
      RefText = RefText & "Source:" & Ret & Ret
      RefText = RefText & "Mullins, M.E., Rogers, T.N., Kline, A.A., Szydlik, C.R." & Ret
      RefText = RefText & "Environmental, Safety, and Health Data Compilation, Policy and Procedures Manual" & Ret
      RefText = RefText & "Michigan Technological University, Houghton, MI, November 1995"
      'GoTo exit_normally_ThisFunc
    Case 201:
      RefText = "Equation Form 201:" & Ret & Ret
      RefText = RefText & "A + BT^2(LN(T)) + CT^2.5 + DT^3" & Ret & Ret
      RefText = RefText & "Where:" & Ret & Ret
      RefText = RefText & "T = Operating Temperature" & Ret & Ret
      RefText = RefText & "Source:" & Ret & Ret
      RefText = RefText & "Mullins, M.E., Rogers, T.N., Kline, A.A., Szydlik, C.R." & Ret
      RefText = RefText & "Environmental, Safety, and Health Data Compilation, Policy and Procedures Manual" & Ret
      RefText = RefText & "Michigan Technological University, Houghton, MI, November 1995"
      'GoTo exit_normally_ThisFunc
  End Select
  '
  ' HANDLE SECOND SET OF TECHNIQUE CODES.
  '
  Select Case inout_TechDat.Technique_Code
    ''''Case "Yaws":
    Case TECHCODE_039_020d_YAWS:
      RefText = "Equation Form:" & Ret & Ret
      RefText = RefText & "log S = -7.861 + 103.032E-03 * Tb + -315.247E-06 * Tb^2 + 262.558E-09 Tb^3" & Ret & Ret
      RefText = RefText & "Where:" & Ret & Ret
      RefText = RefText & "S = Solubility in water at 25 degrees celcius, ppm(wt)" & Ret
      RefText = RefText & "Tb = Boiling point temperature, K" & Ret & Ret
      RefText = RefText & "Source:" & Ret & Ret
      RefText = RefText & "Yaws, C.L.," & Ret
      RefText = RefText & "Thermodynamics and Physical Property Data," & Ret
      RefText = RefText & "Houston: Gulf Publishing Co., 1992."
      'GoTo exit_normally_ThisFunc
    ''''Case "Superfund"
    Case -1000000#:
      RefText = RefText & "Source:" & Ret & Ret
      RefText = RefText & "EPA Superfund Public Health Evaluation Manual," & Ret
      RefText = RefText & "U.S. Environmental Protection Agency," & Ret
      RefText = RefText & "Office of Emergency and Remedial Response," & Ret
      RefText = RefText & "Washington, D.C., EPA/540/1-86/060, October 1986."
      'GoTo exit_normally_ThisFunc
    ''''Case "RTI"
    Case -1000000#:
      RefText = RefText & "Source:" & Ret & Ret
      RefText = RefText & "Ashworth, R.A., Howe, G.B., Mullins, M.E., Rogers, T.N.," & Ret
      RefText = RefText & Chr$(34) & "Air Water Partitioning Coefficients of Organics in Dilute Aqueous Solutions" & Chr$(34) & "," & Ret
      RefText = RefText & "Journal of Hazardous Material, 1988."
      'GoTo exit_normally_ThisFunc
    ''''Case "Ashworth"
    Case -1000000#:
      RefText = "Equation Form:" & Ret & Ret
      RefText = RefText & "ln H = A/T + B" & Ret & Ret
      RefText = RefText & "Where:" & Ret & Ret
      RefText = RefText & "H = Henry's Constant" & Ret
      RefText = RefText & "A = -3024" & Ret
      RefText = RefText & "B = 5.133" & Ret
      RefText = RefText & "T = Temperature (K)" & Ret & Ret
      RefText = RefText & "Source:" & Ret & Ret
      RefText = RefText & "Ashworth, R.A., Howe, G.B., Mullins, M.E., Rogers, T.N.," & Ret
      RefText = RefText & Chr$(34) & "Air Water Partitioning Coefficients of Organics in Dilute Aqueous Solutions" & Chr$(34) & "," & Ret
      RefText = RefText & "Journal of Hazardous Material, 1988."
      'GoTo exit_normally_ThisFunc
    ''''Case "Stephenson"
    Case -1000000#:
      RefText = RefText & "Source:" & Ret & Ret
      RefText = RefText & "Stephenson, R.M.," & Ret
      RefText = RefText & Chr$(34) & "Mutual Solubilities: Water-Ketones, Water-Ethers, and Water-Gasoline-Alcohols" & Chr$(34) & "," & Ret
      RefText = RefText & "Journal of Chemical and Engineering Data, Vol. 37, pp. 80-95, 1992."
      'GoTo exit_normally_ThisFunc
    ''''Case "Chen and Wagner"
    Case -1000000#:
      RefText = RefText & "Source:" & Ret & Ret
      RefText = RefText & "Chen, H., Wagner, J." & Ret
      RefText = RefText & Chr$(34) & "An Efficient and Reliable Gas Chromatographic Method for Measuring" & Ret
      RefText = RefText & "Liquid-Liquid Mutual Solubilities in Alkylbenzene and Water Mixtures" & Chr$(34) & "," & Ret
      RefText = RefText & "Journal of Chemical and Engineering Data, Vol. 39, pp. 475-9, July 1994."
      'GoTo exit_normally_ThisFunc
    ''''Case "Hwang"
    Case -1000000#:
      RefText = RefText & "Source:" & Ret & Ret
      RefText = RefText & "Hwang, Y.L., Keller, G.E. II, Olson J.D.," & Ret
      RefText = RefText & Chr$(34) & "Steam Stripping for Removal of Organic Pollutants from Water" & Chr$(34) & "," & Ret
      RefText = RefText & "Industrial and Engineering Chemistry Research, Vol. 31, No. 7, pp. 1753-1768, 1992."
      'GoTo exit_normally_ThisFunc
    ''''Case "Persichetti"
    Case -1000000#:
      RefText = "Equation Form:" & Ret & Ret
      RefText = RefText & "ln H = C1 + C2/T" & Ret & Ret
      RefText = RefText & "Where:" & Ret & Ret
      RefText = RefText & "H = Henry's Constant" & Ret
      RefText = RefText & "C1 = 20.65" & Ret
      RefText = RefText & "C2 = -4401" & Ret
      RefText = RefText & "T = Temperature (K)" & Ret & Ret
      RefText = RefText & "Source:" & Ret & Ret
      RefText = RefText & "Persichetti, J.M., Coon, J.E., and Twu, C.H.," & Ret
      RefText = RefText & Chr$(34) & "Thermodynamic Model Preparation for Steam Stripping of VOCs" & Chr$(34) & "," & Ret
      RefText = RefText & "Simulation Sciences Inc., Fullerton, CA."
      'GoTo exit_normally_ThisFunc
    ''''Case "Gmehling"
    Case -1000000#:
      RefText = RefText & "Source:" & Ret & Ret
      RefText = RefText & "Gmehling, J., Menke, J., and Schiller, M.," & Ret
      RefText = RefText & "Activity Coefficients at Infinite Dilution C10 - C36 with O2S and H2O," & Ret
      RefText = RefText & "DECHEMA, Vol. IX, Part 4, 1994."
      'GoTo exit_normally_ThisFunc
    ''''Case "ASPEN (AENV)"
    Case -1000000#:
      RefText = RefText & "Source:" & Ret & Ret
      RefText = RefText & "Spencer, J.R.," & Ret
      RefText = RefText & Chr$(34) & "Preliminary Design of a Pilot Scale Steam Stripper" & Chr$(34) & "," & Ret
      RefText = RefText & "MTU Department of Chemical Engineering, October 1996."
      'GoTo exit_normally_ThisFunc
    ''''Case "PEARLS (AENV)", "PEARLS (AVLE)", "PEARLS (AENV)", "PEARLS (ALLE)"
    Case -1000000#:
      RefText = "Equation Form:" & Ret & Ret
      RefText = RefText & "H = Gamma*Pvap" & Ret & Ret
      RefText = RefText & "Where:" & Ret & Ret
      RefText = RefText & "H = Henry's Constant" & Ret
      RefText = RefText & "Gamma = Infinite Dilution Activity Coefficient" & Ret
      RefText = RefText & "Pvap = Vapor Pressure" & Ret
      'GoTo exit_normally_ThisFunc
''''    Case "801 Database"
''''      ' ??? this needs to be fixed
''''      'EquationText = RefText
''''      If Not is_f_of_t(CurProp) Then
''''      RefText = "Double click on the '801 Database' row to see" & Ret
''''      'Else
''''      '    RefText = "DIPPR 801 Database"
''''      End If
''''    Case "911 Database"
''''      ' ??? this needs to be fixed
''''      'EquationText = RefText
''''      If Not is_f_of_t(CurProp) Then
''''      RefText = "Double click on the '911 Database' row to see" & Ret
''''      'Else
''''      '    RefText = "DIPPR 911 Database"
''''      End If
    ''''Case "Bhiruds (1978)"
    Case TECHCODE_001_003e_BHIRUDS_1978:
      RefText = "Average percent error of 10.9 for 55 chemicals from the DIPPR 911 Database" & Ret
      RefText = RefText & "Equation Form:" & Ret & Ret
      RefText = RefText & "Rho = (M*Pc)/(R*T*e^a+Omega*b)" & Ret & Ret
      RefText = RefText & "Where:" & Ret & Ret
      RefText = RefText & "Rho = Liquid Density (mL/g)" & Ret
      RefText = RefText & "M = Molecular Weight (g/mol)" & Ret
      RefText = RefText & "Pc = Critical Pressure (atm)" & Ret
      RefText = RefText & "R = Universal Gas Constant" & Ret
      RefText = RefText & "T = Temperature (K)" & Ret & Ret
      RefText = RefText & "a = 1.39644 - 24.076*Tr + 102.615*Tr^2 - 255.719*Tr^3 + 355.805*Tr^4 - 256.671*Tr^5 + 75.18088*Tr^6" & Ret
      RefText = RefText & "b = 13.4412 - 135.7437*Tr + 533.380*Tr^2 - 1091.453*Tr^3 + 1231.43*Tr^4 - 728.227*Tr^5 + 176.737*Tr^6" & Ret
      RefText = RefText & "   Where:" & Ret
      RefText = RefText & "           Tr = Reduced Temperature, T/Tc" & Ret
      RefText = RefText & "   Omega = (3/7)[Tbr/(1-Tbr)]log(Pc) - 1" & Ret
      RefText = RefText & "   Where:" & Ret
      RefText = RefText & "           Tbr = Reduced Boiling Point Temperature, Tb/Tc" & Ret
      RefText = RefText & "           Pc = Critical Pressure (atm)" & Ret & Ret
      RefText = RefText & "Source:" & Ret & Ret
      RefText = RefText & "Bhirud, Vasant L., " & Chr(34) & "Saturated Liquid" & Ret
      RefText = RefText & "Densities of normal Fluids," & Chr(34) & " AIChE Journal," & Ret
      RefText = RefText & "24(6): 1127-1131 (November, 1978)." & Ret
      'GoTo exit_normally_ThisFunc
    ''''case "Modified Rackett (1978)"
    Case TECHCODE_001_004e_RACKETT_1978:
      RefText = "Average percent error of 11.3 for 55 chemicals from the DIPPR 911 database" & Ret & Ret
      RefText = RefText & "Equation Form:" & Ret & Ret
      RefText = RefText & "1/Rho = (R*Tc/MW*Pc)*Z^[1+(1-Tr)^(2/7)]" & Ret & Ret
      RefText = RefText & "Where:" & Ret & Ret
      RefText = RefText & "R = gas constant" & Ret
      RefText = RefText & "Tc = Critical Temperature, K" & Ret
      RefText = RefText & "Tr = Reduced Temperature, K" & Ret
      RefText = RefText & "Pc = Critical Pressure" & Ret
      RefText = RefText & "Z = Compressibility Factor" & Ret & Ret
      RefText = RefText & "Source:" & Ret & Ret
      RefText = RefText & "Spencer, C.F., Danner, R.P., " & Chr(34) & "Improved Equation" & Ret
      RefText = RefText & "for Prediction of Saturated Liquid Density," & Ret
      RefText = RefText & Chr(34) & " Journal of Chemical and Engineering Data, 17(2): 236-241 (1972)" & Ret
      'GoTo exit_normally_ThisFunc
    ''''Case "Watson"
    Case TECHCODE_010_008e_WATSON:
      RefText = "Equation Form:" & Ret & Ret
      RefText = RefText & "H1 = H2*[(Tc-T2)/(Tc-T1)]^-0.38" & Ret & Ret
      RefText = RefText & "Where:" & Ret & Ret
      RefText = RefText & "H = Enthalpy of vaporization, kJ/mol" & Ret
      RefText = RefText & "Tc = Critical Temperature, K" & Ret
      RefText = RefText & "T1 = Temperature where H is known, K" & Ret
      RefText = RefText & "T2 = Temperature of desired H, K" & Ret & Ret
      RefText = RefText & "Source:" & Ret & Ret
      RefText = RefText & "Reid, R.C., J.M. Prausnitz, and B.E. Poling," & Ret
      RefText = RefText & "The Properties of Gases and Liquids, 4th Edition," & Ret
      RefText = RefText & "McGraw Hill (1987), p.228" & Ret
      'GoTo exit_normally_ThisFunc
    ''''Case "Klein (1949)"
    Case TECHCODE_011_009e_KLEIN_1949:
      RefText = "Average percent error of 10.8 for 61 chemicals from the DIPPR 911 database" & Ret & Ret
      RefText = RefText & "Equation Form:" & Ret & Ret
      RefText = RefText & "H = R*K*Tb*ln(Pc)*[(1-1/[Pc*(Tb/Tc)^3])^(1/2)]/[1-Tb/Tc]" & Ret & Ret
      RefText = RefText & "Where:" & Ret & Ret
      RefText = RefText & "H = Enthalpy of vaporization, cal/mol" & Ret
      RefText = RefText & "R = Ideal Gas Constant = 1.9872 cal/K-mol" & Ret
      RefText = RefText & "K = Klein Constant, unit-less" & Ret
      RefText = RefText & "Tb = Normal Boiling Point, K" & Ret
      RefText = RefText & "Pc = Critical Pressure, atm" & Ret & Ret
      RefText = RefText & "Source:" & Ret & Ret
      RefText = RefText & "Lyman, W., W. Reehl, and D. Rosenblatt, Handbook" & Ret
      RefText = RefText & "of Chemical Property Estimation Methods, McGraw" & Ret
      RefText = RefText & "Hill, (1982)" & Ret
      'GoTo exit_normally_ThisFunc
    ''''Case "Chen and Pitzer (1965)"
    Case TECHCODE_011_010e_CHEN_PITZER_1965:
      RefText = "Average percent error of 10.9 for 61 chemicals from the DIPPR 911 database" & Ret & Ret
      RefText = RefText & "Equation Form:" & Ret & Ret
      RefText = RefText & "H = [Tb(7.11*log(Pc)-7.82+7.9*Tbr)]/(1.07-Tbr)" & Ret & Ret
      RefText = RefText & "Where:" & Ret & Ret
      RefText = RefText & "H = Enthalpy of vaporization, J/mol" & Ret
      RefText = RefText & "Tb = Boiling Point Temperature, K" & Ret
      RefText = RefText & "Tbr = Reduced Boiling Point Temperature, unit-less" & Ret
      RefText = RefText & "Pc = Critical Pressure, atm" & Ret & Ret
      RefText = RefText & "Source:" & Ret & Ret
      RefText = RefText & "Chen, N.H., " & Chr(34) & "Generalized Correlation For" & Ret
      RefText = RefText & "Latent Heat of Vaporization," & Chr(34) & " Journal of" & Ret
      RefText = RefText & "Chemical and Engineering Data, 10: 207-210" & Ret
      RefText = RefText & "(April 1965)." & Ret
      'GoTo exit_normally_ThisFunc
    ''''Case "Lorenz and Herz (1922)"
    Case TECHCODE_003_006e_LORENZ_HERZ_1922:
      RefText = "Average percent error of 15.7 for 94 chemicals from the DIPPR 911 database" & Ret & Ret
      RefText = RefText & "Equation Form:" & Ret & Ret
      RefText = RefText & "Tm = .5839*Tb" & Ret & Ret
      RefText = RefText & "Where:" & Ret & Ret
      RefText = RefText & "Tm = Melting Point Temperature (K)" & Ret
      RefText = RefText & "Tb = Boiling Point Temperature (K)" & Ret & Ret
      RefText = RefText & "Source:" & Ret & Ret
      RefText = RefText & "Horvath, Ari, Molecular Design: Chemical Structure" & Ret
      RefText = RefText & "Generation from the Properties of Pure Organic" & Ret
      RefText = RefText & "Compounds, Elsevier, Amsterdam, (1992)." & Ret
      'GoTo exit_normally_ThisFunc
    ''''Case "Taft and Starek (1930)"
    Case TECHCODE_003_005e_TAFT_STAREK_1930:
      RefText = "Average percent error of 14.0 for 76 chemicals from the DIPPR 911 database" & Ret & Ret
      RefText = RefText & "Equation Form:" & Ret & Ret
      RefText = RefText & "Tm = Tc - Tb" & Ret & Ret
      RefText = RefText & "Where:" & Ret & Ret
      RefText = RefText & "Tm = Melting Point Temperature, K" & Ret
      RefText = RefText & "Tc = Critical Temperature, K" & Ret
      RefText = RefText & "Tb = Boiling Point Temperature, K" & Ret & Ret
      RefText = RefText & "Source:" & Ret & Ret
      RefText = RefText & "Horvath, Ari, Molecular Design: Chemical Structure" & Ret
      RefText = RefText & "Generation from the Properties of Pure Organic" & Ret
      RefText = RefText & "Compounds, Elsevier, Amsterdam, (1992)." & Ret
      'GoTo exit_normally_ThisFunc
    ''''Case "Hansch (1968)"
    Case TECHCODE_034_017e_HANSCH_1968:
      RefText = "Equation Form:" & Ret & Ret
      RefText = RefText & "log(sigma) = 1.214*log(Kow) - 0.850 + 1.744" & Ret & Ret
      RefText = RefText & "Where:" & Ret & Ret
      RefText = RefText & "sigma = activity coeffiecient" & Ret
      RefText = RefText & "Kow = Octanol/Water Partitioning" & Ret & Ret
      RefText = RefText & "Source:" & Ret & Ret
      RefText = RefText & "Hansch, C., Quninlan, J.E., Lawrence, G.L., " & Ret
      RefText = RefText & Chr(34) & "The Linear Free Energy Relationship" & Ret
      RefText = RefText & "Between Partition Coefficients and the Aqueous" & Ret
      RefText = RefText & "Solubility of Organic Liquids," & Chr(34) & " Journal of" & Ret
      RefText = RefText & "Organic Chemistry, 33: 347-350 (1968)." & Ret
      'GoTo exit_normally_ThisFunc
    ''''Case "Kobayshi (1981)"
    Case TECHCODE_037_024e_KOBAYSHI_1981:
      RefText = "Equation Form:" & Ret & Ret
      RefText = RefText & "log BCF = 0.98 log Kow - 0.77" & Ret & Ret
      RefText = RefText & "Where:" & Ret & Ret
      RefText = RefText & "BCF = Bioconcentration Factor" & Ret
      RefText = RefText & "Kow = Octanol-Water Partitioning" & Ret & Ret
      RefText = RefText & "Source:" & Ret & Ret
      RefText = RefText & "Kobayashi, K., Workshop on the Control of Existing" & Ret
      RefText = RefText & "Chemicals Under the Patronage of the Organisation" & Ret
      RefText = RefText & "for Economic Co-operation and Development: Proceedings," & Ret
      RefText = RefText & "Reichstagsgebaude, Berlin(West), 141-63 (1981)." & Ret
      'GoTo exit_normally_ThisFunc
    ''''Case "Kenaga and Goring (1980)"
    Case TECHCODE_037_025e_KENAGA_GORING_1980:
      RefText = "Equation Form:" & Ret & Ret
      RefText = RefText & "log BCF = - 0.564 log S + 2.791" & Ret & Ret
      RefText = RefText & "Where:" & Ret & Ret
      RefText = RefText & "BCF = Bioconcentration Factor" & Ret
      RefText = RefText & "S = Solubility, micro-mol/L" & Ret & Ret
      RefText = RefText & "Source:" & Ret & Ret
      RefText = RefText & "Kenaga, E.E. and C.A.I. Goring, " & Chr(34) & "Relationship" & Ret
      RefText = RefText & "Between Water Solubility, Soil Sorption," & Ret
      RefText = RefText & "Octanol-Water Partitioning, and Concentration" & Ret
      RefText = RefText & "of Chemicals in Biota," & Chr(34) & " In J.G. Eaton, P.R." & Ret
      RefText = RefText & "Parrish, and A.C. Hendricks (eds.) Aquatic" & Ret
      RefText = RefText & "Toxicology, ASTM STP 707, American Society for" & Ret
      RefText = RefText & "Testing Materials, 78-115 (1980)." & Ret
      'GoTo exit_normally_ThisFunc
    ''''Case "Yalkowsky (1990)"
    Case TECHCODE_039_021e_YALKOWSKY_1990:
      RefText = "log Saq = -0.01(Tmp) - log Kow + 0.8" & Ret & Ret
      RefText = RefText & "Where:" & Ret & Ret
      RefText = RefText & "Saq = Aqueous solubility, M/L" & Ret
      RefText = RefText & "Tmp = Solute melting point, liquids assigned value of 25 degrees Celcius" & Ret
      RefText = RefText & "Kow = Octonal-water partition coefficient" & Ret & Ret
      RefText = RefText & "Source:" & Ret & Ret
      RefText = RefText & "Yalkowski, S.H.; Mishra, D.S. " & Chr(34) & "Comment on 'Prediction" & Ret
      RefText = RefText & "of aqueous Solubility of Organic Chemicals Based" & Ret
      RefText = RefText & "on Molecular Structure. 2. Application to PNAs," & Ret
      RefText = RefText & "PCBs, PCDDs, etc.'," & Chr(34) & " Environ. Sci. Tech., 24: 927-929" & Ret
      RefText = RefText & "(1990)." & Ret
      'GoTo exit_normally_ThisFunc
    ''''Case "Kenaga and Goring (1978)"
    Case TECHCODE_035_022e_KENAGA_GORING_1978:
      RefText = "Average percent error of 24.2 for 260 chemicals from the DIPPR 911 database" & Ret & Ret
      RefText = RefText & "Equation Form:" & Ret & Ret
      RefText = RefText & "log Kow = -1.085*log(S) + 4.538" & Ret & Ret
      RefText = RefText & "Where:" & Ret & Ret
      RefText = RefText & "S = Solubility, ppm(wt)" & Ret
      RefText = RefText & "Kow = Octanal-water partition coefficient" & Ret & Ret
      RefText = RefText & "Source:" & Ret & Ret
      RefText = RefText & "Kenaga E.E. and C.A.I. Goring. " & Chr(34) & "Relationship Between" & Ret
      RefText = RefText & "Water Solubility, Soil Sorption, Octanol-Water" & Ret
      RefText = RefText & "Partitioning, and Bioconcentration of Chemicals" & Ret
      RefText = RefText & "in Biota," & Chr(34) & " Aquatic Toxicology, ASTM STP 707," & Ret
      RefText = RefText & "78-115, (1980)." & Ret
      'GoTo exit_normally_ThisFunc
    ''''Case "Hansch KOW (1968)"
    Case TECHCODE_034_017e_HANSCH_1968:
      RefText = "Equation Form:" & Ret & Ret
      RefText = RefText & "log Kow = 0.747*log(1/S) + 0.730" & Ret & Ret
      RefText = RefText & "Where:" & Ret & Ret
      RefText = RefText & "S = Solubility, mol/L" & Ret
      RefText = RefText & "Kow = Octanal-water partition coefficient" & Ret & Ret
      RefText = RefText & "Source:" & Ret & Ret
      RefText = RefText & "Hansch C., J.E. Quinlan and G.L. Lawrence, " & Ret
      RefText = RefText & Chr(34) & "The Linear Free-Energy Relationships" & Ret
      RefText = RefText & "between Partition Coefficients and the Aqueous" & Ret
      RefText = RefText & "Solubility of Organic Liquids," & Chr(34) & " J. Org. Chem.," & Ret
      RefText = RefText & "33:347-50, (1968)." & Ret
      'GoTo exit_normally_ThisFunc
    ''''Case "Hayduk and Laudie (1974)"
    Case TECHCODE_015_012e_HAYDUK_LAUDIE_1974:
      RefText = "Equation Form:" & Ret & Ret
      RefText = RefText & "D = (13.26*10^5)/(mu^1.14 * V^0.589)" & Ret & Ret
      RefText = RefText & "Where:" & Ret & Ret
      RefText = RefText & "D = diffusivity of dilute solute (cm^2/s)" & Ret
      RefText = RefText & "V = molar volume at normal boiling point of solute (cm^3/mole)" & Ret
      RefText = RefText & "mu = solvent viscosity (cp)" & Ret & Ret
      RefText = RefText & "Source:" & Ret & Ret
      RefText = RefText & "Hayduk, W. and H. Laudie, " & Chr(34) & "Prediction of" & Ret
      RefText = RefText & "Diffusion Coefficients for Nonelectrolytes in" & Ret
      RefText = RefText & "Dilute Aqueous Solutions," & Chr(34) & "AIChE Journal 20(3)," & Ret
      RefText = RefText & "611-615, 1974." & Ret
      'GoTo exit_normally_ThisFunc
    ''''Case "Hayduk and Minhas (1982)"
    Case TECHCODE_015_011e_HAYDUK_MINHAS_1982:
      RefText = "Equation Form:" & Ret & Ret
      RefText = RefText & "For normal paraffin solutions:" & Ret
      RefText = RefText & "D = 13.3*10^-8 * T^1.47 * mu^epsilon * V^-0.71" & Ret
      RefText = RefText & "epsilon = (10.2/V) - 0.791" & Ret & Ret
      RefText = RefText & "For aqueous solutions:" & Ret
      RefText = RefText & "D = 1.25*10^-8 * (V^-0.19 - 0.292) * mu^epsilon * T^1.52" & Ret
      RefText = RefText & "epsilon = (9.58/V) - 1.12" & Ret & Ret
      RefText = RefText & "Where:" & Ret & Ret
      RefText = RefText & "D = diffusivity at infinite dilution of 1 into 2 (cm^2/s)" & Ret
      RefText = RefText & "T = absolute temperature (K)" & Ret
      RefText = RefText & "V = molar volume at the normal boiling point (cm^3/mol)" & Ret
      RefText = RefText & "mu = solvent viscosity (mPa s)" & Ret & Ret
      RefText = RefText & "Source:" & Ret & Ret
      RefText = RefText & "Hayduk, W., and B.S. Minhas, " & Chr(34) & "Correlations for" & Ret
      RefText = RefText & "Prediction of Molecular Diffusivities in Liquids," & Ret
      RefText = RefText & Chr(34) & " The Canadian Journal of Chemical Engineering 60," & Ret
      RefText = RefText & "295-299, 1982." & Ret
      'GoTo exit_normally_ThisFunc
    ''''Case "Wilke and Chang"
    Case TECHCODE_015_013e_WILKE_CHANG:
      RefText = "Equation Form:" & Ret & Ret
      RefText = RefText & "D = 7.4*10^-8 * [(phi*M)^1/2 * T]/[mu * V^0.6]" & Ret & Ret
      RefText = RefText & "phi = 2.6 if the solvent is water" & Ret
      RefText = RefText & "    = 1.9 if the solvent is methanol" & Ret
      RefText = RefText & "    = 1.5 if the solvent is ethanol" & Ret
      RefText = RefText & "    = 1.0 if the solvent is unassociated" & Ret & Ret
      RefText = RefText & "Where:" & Ret & Ret
      RefText = RefText & "D = diffusion coefficient (cm^2/s)" & Ret
      RefText = RefText & "M = Molecular weight of the solvent (g/mol)" & Ret
      RefText = RefText & "T = Temperature (K)" & Ret
      RefText = RefText & "V = molar volume at the normal boiling point of the solute (cm^3 mol)" & Ret
      RefText = RefText & "phi = association parameter, multiple of nominal molecular weight of the solvent to give effective value" & Ret
      RefText = RefText & "mu = viscosity of the solvent (cp)" & Ret & Ret
      RefText = RefText & "Source:" & Ret & Ret
      RefText = RefText & "Wilke, C.R. and P. Chang, " & Chr(34) & "Correlation of" & Ret
      RefText = RefText & "Diffusion Coefficients in Dilute Solutions," & Ret
      RefText = RefText & Chr(34) & "AIChE Journal 1, 264-270, 1955." & Ret
      'GoTo exit_normally_ThisFunc
    ''''Case "Brock and Bird (1983)"
    Case TECHCODE_017_015e_BROCK_BIRD_1983:
      RefText = "Average percent error of 6.0 for 134 chemicals from the DIPPR 911 database" & Ret & Ret
      RefText = RefText & "Equation Form:" & Ret & Ret
      RefText = RefText & "sigma = 4.6*10^-4 * [ Pc^2/3 * Tc^1/3 * Q * (1 - Tr)^11/9 ]" & Ret
      RefText = RefText & "Q = 0.1207 * [ 1 + Tbr * (ln(Pc) - 11.526)/(1 - Tbr) ] - 0.281" & Ret & Ret
      RefText = RefText & "Where:" & Ret & Ret
      RefText = RefText & "sigma = surface tension, mN/m" & Ret
      RefText = RefText & "Pc = critical pressure, Pa" & Ret
      RefText = RefText & "Tc = critical temperature, K" & Ret
      RefText = RefText & "Tr = reduced temperature, T/Tc" & Ret
      RefText = RefText & "Tbr = reduce normal boiling point, Tb/Tc" & Ret
      RefText = RefText & "Tb = normal boiling point, K" & Ret & Ret
      RefText = RefText & "Source:" & Ret & Ret
      RefText = RefText & "Danner, R.P. and T.E. Daubert, Manual for Predicting" & Ret
      RefText = RefText & "Chemical Process Design Data, Design Institute for" & Ret
      RefText = RefText & "Physical Property Data, AIChE, New York, 7-1 (1983)." & Ret
      'GoTo exit_normally_ThisFunc
    ''''Case "Wilke and Lee Modificatio"
    Case TECHCODE_016_014e_WILKE_LEE_MOD:
      RefText = "Equation Form:" & Ret & Ret
      RefText = RefText & "D = [ 3.03 - 0.98/(M^1/2) ] * (10^-3) * (T^3/2)/[ P * M^1/2 * sigma^2 * omega ]" & Ret & Ret
      RefText = RefText & "omega = 1.06036/T'^0.15610 + 0.19300/exp(0.47635*T') + 1.03587/exp(1.52996*T') + 1.76474/exp(3.89411*T')" & Ret
      RefText = RefText & "T' = k*T/epsilon" & Ret
      RefText = RefText & "epsilon = (epsilon1 * epsilon2)^1/2" & Ret
      RefText = RefText & "sigma = (sigma1 + sigma2)/2" & Ret
      RefText = RefText & "M = 2 * (1/M1 + 1/M2)^-1" & Ret & Ret
      RefText = RefText & "Where:" & Ret & Ret
      RefText = RefText & "D = Binary diffusion coefficient (cm^2/s)" & Ret
      RefText = RefText & "T = Temperature (K)" & Ret
      RefText = RefText & "M = Molecular Weight (g/mol)" & Ret
      RefText = RefText & "P = Pressure (bar)" & Ret
      RefText = RefText & "sigma = characteristic length (A)" & Ret
      RefText = RefText & "omega = diffusion collision integral (dimensionless)" & Ret & Ret
      RefText = RefText & "Source:" & Ret & Ret
      RefText = RefText & "Wilke, C.R., C.Y. Lee, " & Chr(34) & "Estimation of Diffusion" & Ret
      RefText = RefText & "Coefficients for Gases and Vapors," & Chr(34) & " Industrial" & Ret
      RefText = RefText & "and Engineering Chemistry, 1253-1257, 1955." & Ret
      'GoTo exit_normally_ThisFunc
    ''''Case "Baker (1994)"
    Case TECHCODE_036_023e_BAKER_1994:
      RefText = "Equation Form:" & Ret & Ret
      RefText = RefText & "log(Koc) = 0.904 log(Kow) + 0.086" & Ret & Ret
      RefText = RefText & "Where:" & Ret & Ret
      RefText = RefText & "Kow = Octanon-Water Partitioning" & Ret
      RefText = RefText & "Koc = Organic Carbon-Water Partitioning" & Ret & Ret
      RefText = RefText & "Source:" & Ret & Ret
      RefText = RefText & "Baker, J.R., J.R. Mihelcic, " & Chr(34) & "Estimation" & Ret
      RefText = RefText & "of Organic Carbon Normalized Soil Water Partition" & Ret
      RefText = RefText & "Coefficients," & Chr(34) & " Environmental, Safety, and" & Ret
      RefText = RefText & "Health Data Estimation Manual, (A.A. Kline, T.N." & Ret
      RefText = RefText & "Rogers, and M.E. Mullins, editors), Project 912" & Ret
      RefText = RefText & "Sponsor Release, July 1994, Design Institute for" & Ret
      RefText = RefText & "Physical Property Data, AIChE, New York, NY." & Ret
      'GoTo exit_normally_ThisFunc
    ''''Case "MTU DIPPR"
    Case TECHCODE_030_043e_MTU_DIPPR:
      RefText = "Equation Form:" & Ret & Ret
      RefText = RefText & "COD = ThOD" & Ret & Ret
      RefText = RefText & "Source:" & Ret & Ret
      RefText = RefText & "Kline, A.A., T.N. Rogers, A.J. Pintar, J.R. Mihelcic, E.V. Lutz, M.D. Miller, and M.E. Mullins, " & Chr(34) & "Project 912 Progress Report," & Chr(34) & " Environmental, Safety, and Data Estimation Manual, Project 912 Sponsor Release, December 1996, Design Institute for Physical Property Data, AIChE, New York, NY." & Ret
      'GoTo exit_normally_ThisFunc
    ''''Case "UNIFAC"
    Case -1000000#:
      RefText = "Source:" & Ret & Ret
      RefText = RefText & "Fredenslund, A., Jones, R.L., Prausnitz, J.M., " & Ret
      RefText = RefText & Chr(34) & "Group-Contribution Estimation of" & Ret
      RefText = RefText & "Activity Coefficients in Nonideal Liquid Mixtures" & Ret
      RefText = RefText & Chr(34) & ", AIChE Journal, 31, pp. 1086-1099, 1975." & Ret
      'GoTo exit_normally_ThisFunc
    ''''Case "ASPEN"
    Case -1000000#:
      RefText = "Source:" & Ret & Ret
      RefText = RefText & "ASPEN, Aspen Technology, Inc., Cambridge, MA (1982)." & Ret
      'GoTo exit_normally_ThisFunc
  End Select
exit_normally_ThisFunc:
  Calc_Mod_GetRefText = True
  If (inout_TechDat.ReferenceText <> "") Then
    inout_TechDat.ReferenceText = _
        inout_TechDat.ReferenceText & _
        Ret & Ret
  End If
  inout_TechDat.ReferenceText = _
      inout_TechDat.ReferenceText & _
      RefText
  Exit Function
exit_err_ThisFunc:
  Calc_Mod_GetRefText = False
  Exit Function
err_ThisFunc:
  ''''Call Show_Trapped_Error("Calc_Mod_GetRefText")
  With inout_TechDat
    .ReferenceText = Get_Trapped_Error_String( _
        "Calc_Mod_GetRefText")
  End With
  Resume exit_err_ThisFunc
End Function



