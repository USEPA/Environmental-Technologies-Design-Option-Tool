Attribute VB_Name = "modantoine"
Option Explicit
Const Limit = 10
Dim called As Integer

Dim temphigh As Double
Dim templow As Double
Dim EqNum(1 To Limit) As String
Dim AXE As Double
Dim BXE As Double
Dim CXE As Double
Dim VPANT As Double
Dim tdft As Double
Dim N As Integer    ' wasn't global
Dim X As Double ' wasn't global
'Dim Y As Double
Dim fx As Double
Dim AA As Double
Dim BB As Double
Dim CC As Double
Dim DD As Double
Dim EE As Double

Dim tk As Double

Dim ANTA As Double
Dim ANTB As Double
Dim ANTC As Double

Dim IOBJ As Integer
Dim NL As Integer
'Global TBEGK As Double
'Global TENDK As Double
Dim RSQD As Double
Dim RMPE As Double
Dim rmse As Double

Dim A_ANT As Double
Dim B_ANT As Double
Dim C_ANT As Double
Dim CoeffA(1 To Limit) As Double
Dim CoeffB(1 To Limit) As Double
Dim CoeffC(1 To Limit) As Double
Dim CoeffD(1 To Limit) As Double
Dim CoeffE(1 To Limit) As Double
Dim MinT(1 To Limit) As Double
Dim MaxT(1 To Limit) As Double

Dim units(1 To Limit) As String
Dim TUnits(1 To Limit) As String
Dim VPUnits(1 To Limit) As String
Dim LogForm(1 To Limit) As String
Dim EquationForm(1 To Limit) As String
Dim Pressure_Units(1 To 12) As String
Dim Temperature_Units(1 To 4) As String
Dim from_recalc As Integer

Global Temp_Ant As Double
Public Function get_antione_params(cas_arg As Long, antione_params() As Double) As Boolean

    Dim dbtable As Recordset
    Dim Criteria As Long
    Dim foundinant As Boolean

    'can we calculate antoine params from 911?
    'if so use them if not continue on...
    If (get_antione_911_params(cas_arg)) Then
            Call get_antione_regress(100, "Ln", "Pa", "K", "A - B / (T + C)")
            antione_params(0) = A_ANT
            antione_params(1) = B_ANT
            antione_params(2) = C_ANT
            antione_params(3) = templow
            antione_params(4) = temphigh
            get_antione_params = True
            Exit Function
    End If
    
    'try the antione database for the data now
    On Error GoTo error_db_closed
    Criteria = cas_arg
    Set dbtable = DBJetMaster.OpenRecordset("Antoine List", dbOpenTable)
    On Error GoTo error_db_open
    dbtable.MoveFirst
    foundinant = False
    Do While Not dbtable.EOF
        If dbtable("CAS") = cas_arg Then
            foundinant = True
            GoTo end_of_loop
        End If
        dbtable.MoveNext
    Loop
end_of_loop:
    If dbtable.EOF Then
        'MsgBox ("Chemical not found in Master Database")
        dbtable.Close
        get_antione_params = False
        Exit Function
    End If
    
    antione_params(0) = dbtable("A")
    antione_params(1) = dbtable("B")
    antione_params(2) = dbtable("C")
    antione_params(3) = dbtable("Tmin")
    antione_params(4) = dbtable("Tmax")
    
    get_antione_params = True
    dbtable.Close
    Exit Function
error_db_open:
    get_antione_params = False
    dbtable.Close
error_db_closed:
    get_antione_params = False
End Function

Public Function get_antione_911_params(cas_arg As Long) As Boolean
Dim DBTbl As Recordset
Dim Cas_Number As Double
Dim Code As Integer
Dim Counter As Integer
Dim found As Boolean
Dim name As String
Dim i As Integer

Cas_Number = CDbl(cas_arg)

' Now Search the 911 table for the correct chemical
    Set DBTbl = DBJetMaster.OpenRecordset("DIPPR911", dbOpenTable)
    
    DBTbl.Index = "PrimaryKey1"
    DBTbl.Seek "=", Cas_Number
    
    If DBTbl.NoMatch Then
        MsgBox ("Chemical Not Found in Master Database")
        DBTbl.Close
        Exit Function
    End If
Counter = 1
found = False
' Search all ocurrences of the chemical for PEARLS Code 6 (VP as a F(T))
    Do While (DBTbl("CAS #") = Cas_Number And Counter <= Limit)
        Code = DBTbl("PEARLS Code")
        If Code = 6 Then
            found = True
            MinT(Counter) = DBTbl("Value")
            units(Counter) = ""
            On Error Resume Next
            units(Counter) = DBTbl("Units")
            On Error GoTo Get911DataError
            MaxT(Counter) = DBTbl("Temperature")
            CoeffA(Counter) = DBTbl("Coef1")
            CoeffB(Counter) = DBTbl("Coef2")
            CoeffC(Counter) = DBTbl("Coef3")
            CoeffD(Counter) = DBTbl("Coef4")
            CoeffE(Counter) = DBTbl("Coef5")
            EqNum(Counter) = DBTbl("Equation")
            EqNum(Counter) = EqNum(Counter) & ", " & Counter
            ' All Equations will be 101, so to distinguish they will apear as (101,1  101,2  101,3  ......)
            Counter = Counter + 1
        End If
        DBTbl.MoveNext
    Loop
                   
    DBTbl.Close
On Error GoTo DIPPRError

If (found = False) Then
    get_antione_911_params = False
    MsgBox ("No occurances of vapor pressure as an F(t) where found for " & Trim(name) & ".")
    Exit Function
End If

'IT WORKED
get_antione_911_params = True

Exit Function

Get911DataError:
   
    If Err = 94 Then Resume Next
    MsgBox "Error loading data from 911 database", 48, "Error"
    DBTbl.Close
    get_antione_911_params = False
    Exit Function

DIPPRError:
    MsgBox ("Error with loading data to main form")
    get_antione_911_params = False
End Function
Public Sub get_antione_regress(N_Points As Integer, Log_Type As String, p_Unit As String, temp_Unit As String, Eq_Form As String)
'ANTOINE function takes parameters:
'IFIT, IOBJ, CDFT, TDFT, CLIM, IMAX, TOL, NL, IUNIT, JUNIT, KUNIT, IEQN, TLOW, THGH, NEQN, PUNIT, TUNIT, TLOW, THGH, AA, BB, CC, DD, EE
 

Dim input_IUNIT As Integer
Dim input_JUNIT As Integer
Dim input_KUNIT As Integer
Dim input_IEQN As Integer
Dim input_NEQN As Integer
Dim input_PUNIT As Double
Dim input_TUNIT As Double
Dim input_OTLOW As Double
Dim input_OTHGH As Double
Dim KFLAG As Integer


Screen.MousePointer = 11
called = 1
    
'determine log type to use
If (Trim(Log_Type) = "Ln") Then
    input_IUNIT = 1
ElseIf (Trim(Log_Type) = "Log10") Then
    input_IUNIT = 2
End If

'determine pressure units
Select Case Trim(p_Unit)
    Case "Pa"
        input_JUNIT = 1
    Case "atm"
        input_JUNIT = 2
    Case "bars"
        input_JUNIT = 3
    Case "mm Hg"
        input_JUNIT = 4
    Case "cm Hg"
        input_JUNIT = 5
    Case "m Hg"
        input_JUNIT = 6
    Case "in Hg"
        input_JUNIT = 7
    Case "in H2O"
        input_JUNIT = 8
    Case "ft H2O"
        input_JUNIT = 9
    Case "m H2O"
        input_JUNIT = 10
    Case "psia"
        input_JUNIT = 11
    Case "kPa"
        input_JUNIT = 12
End Select

Select Case Trim(temp_Unit)
    Case "K"
        input_KUNIT = 1
    Case "C"
        input_KUNIT = 2
    Case "F"
        input_KUNIT = 3
    Case "R"
        input_KUNIT = 4
End Select
    
If (input_IUNIT = 1) Then
    input_NEQN = 304
ElseIf (input_IUNIT = 2) Then
    input_NEQN = 300
End If

Select Case Trim(Eq_Form)
    Case "A - B / (T + C)"
        input_NEQN = input_NEQN
    Case "A - B / (T - C)"
        input_NEQN = input_NEQN + 1
    Case "A + B / (T - C)"
        input_NEQN = input_NEQN + 2
    Case "A + B / (T + C)"
        input_NEQN = input_NEQN + 3
End Select

input_IEQN = CInt(Left(EqNum(1), 3))
input_PUNIT = 1 'Find out the difference between
input_TUNIT = 1 'these and JUNIT and KUNIT
input_OTLOW = CDbl(MinT(1))
input_OTHGH = CDbl(MaxT(1))

templow = input_OTLOW
temphigh = input_OTHGH

AA = CoeffA(1)
BB = CoeffB(1)
CC = CoeffC(1)
DD = CoeffD(1)
EE = CoeffE(1)

NL = N_Points
 
Call CalcAntoine(1, 0, 500, 1000, 0.000001, input_IUNIT, input_JUNIT, input_KUNIT, input_NEQN, -999, 999, input_IEQN, input_PUNIT, input_TUNIT, input_OTLOW, input_OTHGH)

Screen.MousePointer = 0
End Sub



Public Sub CalcAntoine(IFIT As Integer, CDFT As Double, CLIM As Double, IMAX As Integer, TOL As Double, IUNIT As Integer, JUNIT As Integer, KUNIT As Integer, NEQN As Integer, TLOW As Double, THGH As Double, IEQN As Integer, PUNIT As Double, TUNIT As Double, OTLOW As Double, OTHGH As Double)
'IOBJ , NL, ANTC, TBEGk, TENDK, ANTA, ANTB, RSQD, RMPE

'IFIT, IOBJ, CDFT, TDFT,                                                                   CLIM, IMAX, TOL, NL, IUNIT,                                                      JUNIT, KUNIT, IEQN, TLOW, THGH, NEQN,                                                                     PUNIT, TUNIT, TLOW, THGH, AA, BB, CC, DD, EE

Dim Conv(12) As Double
Dim inam(2) As String
Dim jnam(12) As String
Dim knam(4) As String
Dim t1 As Double
Dim t2 As Double
'Dim IUNIT As Double
Dim BSIGN2 As Double
Dim CSIGN2 As Double
Dim KLIT As Integer
Dim JLIT As Integer

Dim BSIGN1 As Double
Dim CSIGN1 As Double

'Dim NEQN As Integer
Dim TBEGK As Double
Dim TENDK As Double

Dim ILIT As Integer
Dim AOUT As Double
Dim BOUT As Double
Dim COUT As Double


'Dim rmse As Double


Dim VP_In As Double

Dim KK As Integer
Dim TXE As Double
Dim AXL As Double
Dim BXL As Double
Dim CXL As Double
Dim TXL As Double

'Dim A_ANT As Double
'Dim B_ANT As Double
'Dim C_ANT As Double



Dim temp1 As Double
Dim temp2 As Double



JLIT = PUNIT
KLIT = TUNIT
t1 = TLOW
t2 = THGH
IOBJ = 1
If from_recalc <> 1 Then
    tdft = 25
End If
'...Pa to Pa
    Conv(1) = 1#
    '...Pa to atm
    Conv(2) = 1# / 1.01325 * 10 ^ 5#
    ' ...Pa to bars
    Conv(3) = 1.01325 / 1.10325 * 10 ^ 5#


'     ...['Pa' to 'mm_Hg' (torr)]:
      Conv(4) = 760# / (1.01325 * 10 ^ 5#)
'     ...['Pa' to 'cm_Hg']:
      Conv(5) = 0.1 * Conv(4)
'     ...['Pa' to 'm_Hg']:
      Conv(6) = 0.1 * Conv(5)
'     ...['Pa' to 'in_Hg']:
      Conv(7) = Conv(4) * (39.37 / 1000#)
'     ...['Pa' to 'in_H2O']:
      Conv(8) = (33.9 * 12#) / (1.01325 * 10 ^ 5#)
'     ...['Pa' to 'ft_H2O']:
      Conv(9) = 33.9 / 101325#
'     ...['Pa' to 'm_H2O']:
      Conv(10) = 10.333 / 101325#
'     ...['Pa' to 'psia']:
      Conv(11) = 14.696 / 101325#
'     ...['Pa' to 'kPa']:
      Conv(12) = 0.001

'        -- Units Labels --

      inam(1) = "{    Ln, "
      inam(2) = "{ Log10, "
      jnam(1) = "    Pa, "
      jnam(2) = "   atm, "
      jnam(3) = "  bars, "
      jnam(4) = " mm_Hg, "
      jnam(5) = " cm_Hg, "
      jnam(6) = "  m_Hg, "
      jnam(7) = " in_Hg, "
      jnam(8) = "in_H2O, "
      jnam(9) = "ft_H2O, "
      jnam(10) = " m_H2O, "
      jnam(11) = "  psia, "
      jnam(12) = "   kPa, "
      knam(1) = "K }"
      knam(2) = "C }"
      knam(3) = "F }"
      knam(4) = "R }"
      
      t1 = t1 + 273.15
      t2 = t2 + 273.15
      
        If NEQN = 300 Then
            IUNIT = 2
            BSIGN2 = -1#
            CSIGN2 = 1#
        ElseIf NEQN = 301 Then
                IUNIT = 2
                BSIGN2 = -1#
                CSIGN2 = -1#
        ElseIf NEQN = 302 Then
                IUNIT = 2
                BSIGN2 = 1#
                CSIGN2 = -1#
        ElseIf NEQN = 303 Then
                IUNIT = 2
                BSIGN2 = 1#
                CSIGN2 = 1#
        ElseIf NEQN = 304 Then
                IUNIT = 1
                BSIGN2 = -1#
                CSIGN2 = 1#
        ElseIf NEQN = 305 Then
                IUNIT = 1
                BSIGN2 = -1#
                CSIGN2 = -1#
         ElseIf NEQN = 306 Then
                IUNIT = 1
                BSIGN2 = 1#
                CSIGN2 = -1#
         ElseIf NEQN = 307 Then
                IUNIT = 1
                BSIGN2 = 1#
                CSIGN2 = 1#
        End If

     If IEQN = 101 Then
            TBEGK = OTLOW
            TENDK = OTHGH
            GoTo out1
         End If
         If IEQN = 300 Then
            ILIT = 2
            BSIGN1 = -1#
            CSIGN1 = 1#
         ElseIf IEQN = 301 Then
            ILIT = 2
            BSIGN1 = -1#
            CSIGN1 = -1#
         ElseIf IEQN = 302 Then
            ILIT = 2
            BSIGN1 = 1#
            CSIGN1 = -1#
         ElseIf IEQN = 303 Then
            ILIT = 2
            BSIGN1 = 1#
            CSIGN1 = 1#
         ElseIf IEQN = 304 Then
            ILIT = 1
            BSIGN1 = -1#
            CSIGN1 = 1#
         ElseIf IEQN = 305 Then
            ILIT = 1
            BSIGN1 = -1#
            CSIGN1 = -1#
         ElseIf IEQN = 306 Then
            ILIT = 1
            BSIGN1 = 1#
            CSIGN1 = -1#
         ElseIf IEQN = 307 Then
            ILIT = 1
            BSIGN1 = 1#
            CSIGN1 = 1#
         End If
         BB = BB * BSIGN1
         CC = CC * CSIGN1
out1:
'        -- Convert Antoine Coeff. Inputs to {Ln,Pa,K} Units --

         If (IEQN >= 300) And (IEQN <= 307) Then
            If KLIT = 1 Then
               BOUT = BB
               COUT = CC
               TBEGK = OTLOW
               TENDK = OTHGH
            ElseIf KLIT = 2 Then
               BOUT = BB
               COUT = CC - 273.15
               TBEGK = OTLOW + 273.15
               TENDK = OTHGH + 273.15
            ElseIf KLIT = 3 Then
               BOUT = BB / 1.8
               COUT = ((CC + 32#) - (1.8 * 273.15)) / 1.8
               TBEGK = ((OTLOW - 32#) / 1.8) + 273.15
               TENDK = ((OTHGH - 32#) / 1.8) + 273.15
            ElseIf KLIT = 4 Then
               BOUT = BB / 1.8
               COUT = CC / 1.8
               TBEGK = OTLOW / 1.8
               TENDK = OTHGH / 1.8
            End If
            If ILIT = 1 Then
               temp2 = Conv(JLIT)
               temp1 = Log(temp2)
               AOUT = AA - temp1
            ElseIf ILIT = 2 Then
               AOUT = (AA - (Log(Conv(JLIT)) / Log(10))) * 2.3025851
               BOUT = BOUT * 2.3025851
            End If
            ANTA = AOUT
            ANTB = BOUT
            ANTC = COUT
            RSQD = 0#
            RMPE = 0#
            rmse = 0#
            GoTo convertout
         ElseIf IEQN = 101 Then

'        -- Optimize Antoine "ANTA", "ANTB", and "ANTC" Parameters --

                If t1 > TBEGK Then
                    TBEGK = t1
                End If
'         tbegk = Max(tbegk, t1)
                If t2 < TENDK Then
                    TENDK = t2
                End If
       '  tendk = Min(tendk, t2)
         ANTC = CDFT
         If (IFIT - 1 = 0) Then GoTo out4  '30
         Else: GoTo out3 '40
         End If
out4:
        
           Call golden(IMAX, TOL, CLIM, TENDK, TBEGK)
            If (N <= 0) Then GoTo out16 '10  '(exit function do i need to return
                                       'a value??)
            ANTC = X
'            End If
out3:
            rmse = rlinear(X, TENDK, TBEGK)
        ' End If

'        -- Default Vapor Pressure (VP_IN, in {Pa}) from Input Eqn. --
convertout:
         tk = tdft + 273.15
         If (IEQN = 101) Then
            VP_In = Exp(AA + (BB / tk) + (CC * Log(tk)) + (DD * (tk) ^ EE))
         Else
            VP_In = Exp(ANTA + ANTB / (tk + ANTC))
         End If

'        -- Units Conversions for "ANTA", "ANTB", and "ANTC" --

         For KK = 1 To 12 Step 1
            If ((JUNIT > 0) And (JUNIT <> KK)) Then GoTo out8 '60
            If ((KUNIT > 0) And (KUNIT <> 1)) Then GoTo out5   '75
            If ((IUNIT > 0) And (IUNIT <> 1)) Then GoTo out15  '70

'           * Temperature Units of [K] *

            AXE = ANTA + Log(Conv(KK))
            BXE = ANTB
            CXE = ANTC
            TXE = tk
            TLOW = TBEGK
            THGH = TENDK
            VPANT = Exp(AXE + BXE / (TXE + CXE)) / Conv(KK)
            BXE = BXE * BSIGN2
            CXE = CXE * CSIGN2
'C            WRITE(7,901) NAME,ICAS,TLOW,THGH,AXE,BXE,CXE,RSQD,RMPE,RMSE,
'C     &                   TXE,VPANT,VP_IN,INAM(1),JNAM(KK),KNAM(1),NEQN

    
out15:
            
            AXL = (ANTA + Log(Conv(KK))) / 2.3025851
            'FOR SOME REASON IF jUNIT IS MM HG THEN A COEFF IS OFF BY 10
           
            
            BXL = ANTB / 2.3025851
            CXL = ANTC
            TXL = tk
            TLOW = TBEGK
            THGH = TENDK
            VPANT = (10# ^ (AXL + BXL / (TXL + CXL))) / Conv(KK)
            BXL = BXL * BSIGN2
            CXL = CXL * CSIGN2
'C            WRITE(7,901) NAME,ICAS,TLOW,THGH,AXL,BXL,CXL,RSQD,RMPE,RMSE,
'C     &                   TXL,VPANT,VP_IN,INAM(2),JNAM(KK),KNAM(1),NEQN
out5:
            
            If ((KUNIT > 0) And (KUNIT <> 2)) Then GoTo out10  '85
            If ((IUNIT > 0) And (IUNIT <> 1)) Then GoTo out9 '80
        
        '           * Temperature Units of [C] *

            AXE = ANTA + Log(Conv(KK))
            BXE = ANTB
            CXE = ANTC + 273.15
            TXE = tk - 273.15
            TLOW = TBEGK - 273.15
            THGH = TENDK - 273.15
            VPANT = Exp(AXE + BXE / (TXE + CXE)) / Conv(KK)
            BXE = BXE * BSIGN2
            CXE = CXE * CSIGN2
'C            WRITE(7,901) NAME,ICAS,TLOW,THGH,AXE,BXE,CXE,RSQD,RMPE,RMSE,
'C     &                   TXE,VPANT,VP_IN,INAM(1),JNAM(KK),KNAM(2),NEQN
out9:
        
            AXL = (ANTA + Log(Conv(KK))) / 2.3025851
            'FOR SOME REASON IF jUNIT IS MM HG THEN A COEFF IS OFF BY 10
           
            
            BXL = ANTB / 2.3025851
            CXL = ANTC + 273.15
            TXL = tk - 273.15
            TLOW = TBEGK - 273.5
            THGH = TENDK - 273.5
            VPANT = (10# ^ (AXL + BXL / (TXL + CXL))) / Conv(KK)
            BXL = BXL * BSIGN2
            CXL = CXL * CSIGN2
'C            WRITE(7,901) NAME,ICAS,TLOW,THGH,AXL,BXL,CXL,RSQD,RMPE,RMSE,
'C     &                   TXL,VPANT,VP_IN,INAM(2),JNAM(KK),KNAM(2),NEQN
out10:
            
            If ((KUNIT > 0) And (KUNIT <> 3)) Then GoTo out12  '95
            If ((IUNIT > 0) And (IUNIT <> 1)) Then GoTo out11  '90

'C           * Temperature Units of [F] *

            AXE = ANTA + Log(Conv(KK))
            BXE = ANTB * 1.8
            CXE = ANTC * 1.8 + (1.8 * 273.15) - 32#
            TXE = 1.8 * (tk - 273.15) + 32#
            TLOW = 1.8 * (TBEGK - 273.15) + 32#
            THGH = 1.8 * (TENDK - 273.15) + 32#
            VPANT = Exp(AXE + BXE / (TXE + CXE)) / Conv(KK)
            BXE = BXE * BSIGN2
            CXE = CXE * CSIGN2
'C            WRITE(7,901) NAME,ICAS,TLOW,THGH,AXE,BXE,CXE,RSQD,RMPE,RMSE,
'C     &                   TXE,VPANT,VP_IN,INAM(1),JNAM(KK),KNAM(3),NEQN
out11:
            
            AXL = (ANTA + Log(Conv(KK))) / 2.3025851
            'FOR SOME REASON IF jUNIT IS MM HG THEN A COEFF IS OFF BY 10
            
            
            BXL = (ANTB * 1.8) / 2.3025851
            CXL = ANTC * 1.8 + (1.8 * 273.15) - 32#
            TXL = 1.8 * (tk - 273.15) + 32#
            TLOW = 1.8 * (TBEGK - 273.15) + 32#
            THGH = 1.8 * (TENDK - 273.15) + 32#
            VPANT = 10# ^ (AXL + BXL / (TXL + CXL)) / Conv(KK)
            BXL = BXL * BSIGN2
            CXL = CXL * CSIGN2
'C            WRITE(7,901) NAME,ICAS,TLOW,THGH,AXL,BXL,CXL,RSQD,RMPE,RMSE,
'C     &                   TXL,VPANT,VP_IN,INAM(2),JNAM(KK),KNAM(3),NEQN
out12:
            
            If ((KUNIT > 0) And (KUNIT <> 4)) Then GoTo out7  '105
            If ((IUNIT > 0) And (IUNIT <> 1)) Then GoTo out13  '100

'C           * Temperature Units of [R] *

            AXE = ANTA + Log(Conv(KK))
            BXE = ANTB * 1.8
            CXE = ANTC * 1.8
            TXE = tk * 1.8
            TLOW = TBEGK * 1.8
            THGH = TENDK * 1.8
            VPANT = Exp(AXE + BXE / (TXE + CXE)) / Conv(KK)
            BXE = BXE * BSIGN2
            CXE = CXE * CSIGN2
'C            WRITE(7,901) NAME,ICAS,TLOW,THGH,AXE,BXE,CXE,RSQD,RMPE,RMSE,
'C     &                   TXE,VPANT,VP_IN,INAM(1),JNAM(KK),KNAM(4),NEQN
out13:
            
            AXL = (ANTA + Log(Conv(KK))) / 2.3025851
            'FOR SOME REASON IF jUNIT IS MM HG THEN A COEFF IS OFF BY 10
            
            BXL = (ANTB * 1.8) / 2.3025851
            CXL = ANTC * 1.8
            TXL = tk * 1.8
            TLOW = TBEGK * 1.8
            THGH = TENDK * 1.8
            VPANT = 10# ^ (AXL + BXL / (TXL + CXL)) / Conv(KK)
            BXL = BXL * BSIGN2
            CXL = CXL * CSIGN2
'C            WRITE(7,901) NAME,ICAS,TLOW,THGH,AXL,BXL,CXL,RSQD,RMPE,RMSE,
'C     &                   TXL,VPANT,VP_IN,INAM(2),JNAM(KK),KNAM(4),NEQN
out7:
         
out8:
        Next KK
 '        WRITE(7,'()')
         If ((IUNIT = 0) Or (JUNIT = 0) Or (KUNIT = 0)) Then GoTo out6 '10

'C        -- Store Antoine Coefficients in Results Vectors --

         If (IUNIT = 1) Then
            A_ANT = AXE
            B_ANT = BXE
            C_ANT = CXE
            Temp_Ant = TXE
         ElseIf (IUNIT = 2) Then
            A_ANT = AXL
            B_ANT = BXL
            C_ANT = CXL
            Temp_Ant = TXL
         End If
out6:
      
out14:
        
'C         WRITE(7,'(//)')

out16:
'      Return


'        -- Output Format --

' 901  FORMAT(1X,A25,I10,11(1PE12.4),5X,A9,A8,A3,2X,I5)

        
        
End Sub

Public Sub golden(IMAX As Integer, TOL As Double, xlim As Double, TENDK As Double, TBEGK As Double)
' *********************************************************************
'             "SECTION" SEARCH SUBROUTINE
'             ----------------------------------
'
'     SOLVES FOR THE NON-LINEAR ANTOINE "ANTC" PARAMETER
'
' *********************************************************************
     
' *********************************************************************
 
'        -- OBJECTIVE FUNCTION TO BE MINIMIZED --


'Dim Y As Double
Dim KFLAG As Integer
Dim A As Double
Dim B As Double
Dim ftest2 As Double
Dim ftest1 As Double
Dim UNC As Double
Dim X1 As Double
Dim x2 As Double
Dim fx1 As Double
Dim fx2 As Double
Dim i4 As Integer
     ' object(x) = rlinear(IOBJ, NL, x, TBEGK, TENDK, ANTA, ANTB, RSQD, RMPE)
'
'had variable x after NL not sure what that was..  repleaced with antc
'-- STATEMENT FUNCTION TO IMPLEMENT SECTION --

      'sect(x, y) = x + 0.618 * y

      N = 0
      KFLAG = 0
      If (TBEGK - 10#) < xlim Then
         A = -(TBEGK - 10#)
      Else
          A = -xlim
    End If
 '     a = -Min((TBEGK - 10#), xlim)
      B = A

'        -- BRACKET ROOT & SET UNCERTAINTY INTERVAL (UNC) --
         
         ftest2 = object(B, TENDK, TBEGK)
out12:
         
         ftest1 = ftest2
         B = B + (xlim / CDbl(IMAX))
         'DBLE(IMAX))
         If (B >= xlim) Then GoTo out11  '998
         
         ftest2 = object(B, TENDK, TBEGK)
         If (ftest1 - ftest2) < 0 Then
            GoTo out5  '10
          End If
          If (ftest1 - ftest2) >= 0 Then
            GoTo out12   '5
          End If
out5:
      
      UNC = B - A
'Tony write statement
      
      If (UNC <= TOL) Then GoTo out4  '45
      If (N = IMAX) Then GoTo out8
      If (N = 0) Then GoTo out1  '15
      If (KFLAG = 1) Then GoTo out2 '30
      
      If (KFLAG = 2) Then GoTo out6  '40
    
out1:
      X1 = sect(B, -UNC)
      
      fx1 = object(X1, TENDK, TBEGK)
      If (N > 0) Then GoTo out3  '25
out10:    'goto 20 label
      x2 = sect(A, UNC)
      fx2 = object(x2, TENDK, TBEGK)
out3:  'goto 25 label
      N = N + 1
      If (fx1 >= fx2) Then GoTo out7   '35
'        -- BRANCH FOR F(X1) < F(X2) --

      KFLAG = 1
      B = x2
        GoTo out5  '10
out2:
      x2 = X1
      fx2 = fx1
        GoTo out1  '15

'        -- BRANCH FOR F(X1) > F(X2) --

out7:  'goto 35 label
      KFLAG = 2
      A = X1
        GoTo out5   '10
out6:  'goto 40 label
      X1 = x2
      fx1 = fx2
        GoTo out10  '20

out4:
      X = (A + B) / 2#
      fx = object(X, TENDK, TBEGK)
'      Return

'        -- NON-CONVERGENCE FAILURE MESSAGES --
Exit Sub
out11:
     
    '  WRITE (7,201)
    '  WRITE (*,201)
 '201  FORMAT(//,1X,'******  ERROR :  "ANTC" ROOT IS NOT BOUNDED',/)
    MsgBox ("ANTC ROOT is not  Bounded ")
    GoTo out9  '1000
out8:
      
   '   WRITE (7,200) N
   '   WRITE (*,200) N
 MsgBox ("error: Subroutine Golden did not find the ANTC root after" & CStr(i4) & "iterations")
 
    
' 200  FORMAT(//,1X,'******  ERROR :  SUBROUTINE GOLDEN DID NOT FIND THE
  '   & "ANTC" ROOT AFTER ',I4,'  ITERATIONS',/)
out9:

'        -- RETURN "FAILURE" DEFAULTS --

      N = 0
      X = 0#
      fx = 0#
'      Return
     
End Sub

Public Function rlinear(CEST As Double, TENDK As Double, TBEGK As Double)
'***********************************************************************

'     -- FOR "ANTC" ESTIMATE, COMPUTE "ANTA" "ANTB" "RMSE" --

'***********************************************************************
'      REAL*8 FUNCTION RLINEAR(IOBJ,NL,CEST,TBEGK,TENDK,ANTA,ANTB,RSQD,
 '***********************************************************************
'      IMPLICIT REAL*8(A-H,O-Z)
 '     COMMON /DIPPR/ AA,BB,CC,DD,EE
    Dim dp As Double
    Dim XBEG As Double
    Dim XEND As Double
    Dim XINC As Double
    Dim SUMX As Double
    Dim SUMY As Double
    Dim SUMXX As Double
    Dim SUMYY As Double
    Dim SUMXY As Double
    Dim K As Integer
    Dim xnow As Double
    Dim tnow As Double
    Dim XCOORD As Double
    Dim VPLN As Double
    Dim B0 As Double
    Dim B1 As Double
    Dim rmse As Double
    Dim J As Integer
    Dim vpcal As Double
    Dim rsav As Double
    'Dim CEST As Double
    
    
'        -- DIPPR 801/911 Vapor Pressure Function --

'      vpdippr(tk) = AA + BB / tk + CC * Log(tk) + DD * (tk) ^ EE

      ANTC = CEST
      dp = CDbl(NL)
      XBEG = 1# / TENDK
      XEND = 1# / TBEGK
      XINC = (XEND - XBEG) / CDbl(NL - 1)

         SUMX = 0#
         SUMY = 0#
         SUMXX = 0#
         SUMYY = 0#
         SUMXY = 0#
      For K = 1 To NL Step 1
         xnow = XBEG + CDbl(K - 1) * XINC
         tnow = (1# / xnow)
         XCOORD = 1# / (tnow + ANTC)
         VPLN = vpdippr(tnow)

         SUMX = SUMX + XCOORD
         SUMY = SUMY + VPLN
         SUMXX = SUMXX + (XCOORD) ^ 2#
         SUMYY = SUMYY + (VPLN) ^ 2#
         SUMXY = SUMXY + (XCOORD * VPLN)
      Next K

'        -- Antoine "ANTA" and "ANTB" by Linear Regression --

      B0 = (SUMY / dp) - (dp * SUMX * SUMXY - SUMX ^ 2# * SUMY) / dp / (dp * SUMXX - SUMX ^ 2#)
      B1 = (dp * SUMXY - SUMX * SUMY) / (dp * SUMXX - SUMX ^ 2#)
      ANTA = B0
      ANTB = B1
      RSQD = 1# - (SUMYY - B0 * SUMY - B1 * SUMXY) / ((dp * SUMYY - SUMY ^ 2#) / dp)

'        -- Root-Mean-Square Error (RMSE) --

         rmse = 0#
         RMPE = 0#
      
      
      For J = 1 To NL Step 1
         
         xnow = XBEG + CDbl(J - 1) * XINC
         tnow = (1# / xnow)
         XCOORD = 1# / (tnow + ANTC)
         VPLN = vpdippr(tnow)
         vpcal = ANTA + ANTB * XCOORD

'        * Absolute Logarithmic {log10} Error *

         rmse = rmse + ((vpcal - VPLN) / 2.3025851) ^ 2

'        * Percentage Difference {%} Error *

         RMPE = RMPE + (100# * (Exp(vpcal) - Exp(VPLN)) / Exp(VPLN)) ^ 2#

         Next J
         
         
         
         
         rmse = Sqr(rmse / dp)
         RMPE = Sqr(RMPE / dp)
         If (IOBJ = 2) Then
            rsav = rmse
            rmse = RMPE
            RMPE = rsav
         End If
         'TOny's write statement
         rlinear = rmse
'      Return
End Function

Public Function sect(inval1 As Double, inval2 As Double)

 sect = inval1 + 0.618 * inval2


End Function

Public Function object(inval As Double, TENDK As Double, TBEGK As Double)


object = rlinear(inval, TENDK, TBEGK)

End Function

Public Function vpdippr(datapoint As Double)

vpdippr = AA + BB / datapoint + CC * Log(datapoint) + DD * (datapoint) ^ EE

End Function

Public Sub load_frm_antoine()
'Clear Frame Regress
Call Init_Antoine_FrameVP
'Call Initialise_FrameInputs
Call Init_Antoine_FrameInputs
' Puts the starting output frame on top before the user
' desides whether they want to use regression or unit
'conversion
Call Antoine_FrameStartOnTop
If InfoMethod(VP).Enabled(InfoMethod(VP).CurMethod) = True Then
    FRMAntoine!txtVP.caption = InfoMethod(VP).value(InfoMethod(VP).CurMethod)
End If
FRMAntoine!TXTStartingCoeffA.Enabled = False
FRMAntoine!TXTStartingCoeffB.Enabled = False
FRMAntoine!TXTStartingCoeffC.Enabled = False
FRMAntoine!TXTStartingCoeffD.Enabled = False
FRMAntoine!TXTStartingCoeffE.Enabled = False
FRMAntoine!CMBEquationNumber.Enabled = False
FRMAntoine!TXTEquation.Enabled = False

FRMAntoine!LBLStartingCoeffA.Enabled = False
FRMAntoine!LBLStartingCoeffB.Enabled = False
FRMAntoine!LBLStartingCoeffC.Enabled = False
FRMAntoine!LBLStartingCoeffD.Enabled = False
FRMAntoine!LBLStartingCoeffE.Enabled = False
FRMAntoine!LBLEquationNumber.Enabled = False
FRMAntoine!LBLEquation.Enabled = False
FRMAntoine!VertScrollEquation.Enabled = False

FRMAntoine!LBLVPUnits.Enabled = False
FRMAntoine!LBLTempUnits.Enabled = False
FRMAntoine!LBLLogForm.Enabled = False
FRMAntoine!LBLEquationForm.Enabled = False
FRMAntoine!LBLTempRange.Enabled = False
FRMAntoine!LBLTempTo.Enabled = False
FRMAntoine!LBLRegPoints.Enabled = False
FRMAntoine!CMDRegress.Enabled = False
FRMAntoine!CMDConvert.Enabled = False

FRMAntoine!CMBVPUnits.Enabled = False
FRMAntoine!CMBTempUnits.Enabled = False
FRMAntoine!CMBLogForm.Enabled = False
FRMAntoine!CMBEquationForm.Enabled = False
FRMAntoine!TXTTempFrom.Enabled = False
FRMAntoine!TXTTempTo.Enabled = False
FRMAntoine!TXTRegressionPoints.Enabled = False

End Sub

Public Sub Init_Antoine_FrameVP()

FRMAntoine!TXTChemName = Cur_Info.name
FRMAntoine!TXTStartingCoeffA = ""
FRMAntoine!TXTStartingCoeffB = ""
FRMAntoine!TXTStartingCoeffC = ""
FRMAntoine!TXTStartingCoeffD = ""
FRMAntoine!TXTStartingCoeffE = ""

FRMAntoine!CMBEquationNumber = ""
FRMAntoine!TXTEquation = ""
End Sub

Public Sub Init_Antoine_FrameInputs()


FRMAntoine!CMBVPUnits = ""
FRMAntoine!CMBTempUnits = ""
FRMAntoine!CMBLogForm = ""
FRMAntoine!CMBEquationForm = ""
FRMAntoine!TXTTempFrom = ""
FRMAntoine!TXTTempTo = ""
FRMAntoine!TXTRegressionPoints = ""


End Sub

Public Sub Antoine_FrameStartOnTop()
' This routine makes the starting Frame be visible and
' the others not
FRMAntoine!FrameStart.Visible = True

FRMAntoine!FrameRegress.Visible = False
FRMAntoine!FrameStatistics.Visible = False
End Sub

Public Sub do_antoine_regress()
'ANTOINE function takes parameters:
'IFIT, IOBJ, CDFT, TDFT, CLIM, IMAX, TOL, NL, IUNIT, JUNIT, KUNIT, IEQN, TLOW, THGH, NEQN, PUNIT, TUNIT, TLOW, THGH, AA, BB, CC, DD, EE
Dim input_NL As Integer
Dim input_IUNIT As Integer
Dim input_JUNIT As Integer
Dim input_KUNIT As Integer
Dim input_IEQN As Integer
Dim input_NEQN As Integer
Dim input_PUNIT As Double
Dim input_TUNIT As Double
Dim input_OTLOW As Double
Dim input_OTHGH As Double
Dim input_AA As Double
Dim input_BB As Double
Dim input_CC As Double
Dim input_DD As Double
Dim input_EE As Double
Dim KFLAG As Integer

Screen.MousePointer = 11
called = 1
input_NL = CInt(Trim(FRMAntoine!TXTRegressionPoints))
If (Trim(FRMAntoine!CMBLogForm) = "Ln") Then
    input_IUNIT = 1
ElseIf (Trim(FRMAntoine!CMBLogForm) = "Log10") Then
    input_IUNIT = 2
End If
Select Case Trim(FRMAntoine!CMBVPUnits)
    Case "Pa"
        input_JUNIT = 1
    Case "atm"
        input_JUNIT = 2
    Case "bars"
        input_JUNIT = 3
    Case "mm Hg"
        input_JUNIT = 4
    Case "cm Hg"
        input_JUNIT = 5
    Case "m Hg"
        input_JUNIT = 6
    Case "in Hg"
        input_JUNIT = 7
    Case "in H2O"
        input_JUNIT = 8
    Case "ft H2O"
        input_JUNIT = 9
    Case "m H2O"
        input_JUNIT = 10
    Case "psia"
        input_JUNIT = 11
    Case "kPa"
        input_JUNIT = 12
End Select

Select Case Trim(FRMAntoine!CMBTempUnits)
    Case "K"
        input_KUNIT = 1
    Case "C"
        input_KUNIT = 2
    Case "F"
        input_KUNIT = 3
    Case "R"
        input_KUNIT = 4
End Select
    
input_IEQN = CInt(Left(FRMAntoine!CMBEquationNumber, 3))
If (input_IUNIT = 1) Then
    input_NEQN = 304
ElseIf (input_IUNIT = 2) Then
    input_NEQN = 300
End If
Select Case Trim(FRMAntoine!CMBEquationForm)
    Case "A - B / (T + C)"
        input_NEQN = input_NEQN
    Case "A - B / (T - C)"
        input_NEQN = input_NEQN + 1
    Case "A + B / (T - C)"
        input_NEQN = input_NEQN + 2
    Case "A + B / (T + C)"
        input_NEQN = input_NEQN + 3
End Select
input_PUNIT = 1 'Find out the difference between
input_TUNIT = 1 'these and JUNIT and KUNIT
input_OTLOW = CDbl(Trim(FRMAntoine!TXTTempFrom))
input_OTHGH = CDbl(Trim(FRMAntoine!TXTTempTo))
input_AA = CDbl(Trim(FRMAntoine!TXTStartingCoeffA))
input_BB = CDbl(Trim(FRMAntoine!TXTStartingCoeffB))
input_CC = CDbl(Trim(FRMAntoine!TXTStartingCoeffC))
input_DD = CDbl(Trim(FRMAntoine!TXTStartingCoeffD))
input_EE = CDbl(Trim(FRMAntoine!TXTStartingCoeffE))


'Call AntoineFunction(1, 1, 0#, 25#, 500#, 100, 0.000001, input_NL, input_IUNIT, input_JUNIT, input_KUNIT, input_NEQN, -999#, 999#, input_IEQN, input_PUNIT, input_TUNIT, input_OTLOW, input_OTHGH, input_AA, input_BB, input_CC, input_DD, input_EE)
NL = input_NL
AA = input_AA
BB = input_BB
CC = input_CC
DD = input_DD
EE = input_EE

temphigh = input_OTHGH
templow = input_OTLOW
Call CalcAntoine(1, 0, 500, 1000, 0.000001, input_IUNIT, input_JUNIT, input_KUNIT, input_NEQN, -999, 999, input_IEQN, input_PUNIT, input_TUNIT, input_OTLOW, input_OTHGH)



A_ANT = Format(A_ANT, "####0.0000")
B_ANT = Format(B_ANT, "####0.0000")
C_ANT = Format(C_ANT, "####0.0000")

FRMAntoine!TXTANTA = A_ANT
FRMAntoine!TXTANTB = B_ANT
FRMAntoine!TXTANTC = C_ANT
VPANT = Format(VPANT, "####0.0000")
FRMAntoine!TXTVaporPressure = VPANT
FRMAntoine!TXTANTEqnNum = FRMAntoine!CMBEquationNumber
FRMAntoine!TXTANTEquation = FRMAntoine!TXTEquation
RSQD = Format(RSQD, "####0.000000")
FRMAntoine!TXTRSQR = RSQD
rmse = Format(rmse, "####0.0000")
FRMAntoine!TXTRMSE = rmse
RMPE = Format(RMPE, "###0.0000")
FRMAntoine!TXTRMPE = RMPE
'If CMBTempUnits.Text = "K" Then
 '   tdft = tdft + 273.15
'End If
FRMAntoine!TXTANTTemp = Temp_Ant
FRMAntoine!TXTANTTempUnits = FRMAntoine!CMBTempUnits
FRMAntoine!TXTANTPressUnits = FRMAntoine!CMBVPUnits
Call FRMAntoine.FrameRegressOnTop
Screen.MousePointer = 0
End Sub

Public Sub do_antoine_to_db()
' do we really want to edit the master?
Dim foundinant As Integer
Dim tableantoine As Recordset
' for now let's do this
MsgBox ("Database not editable")
Exit Sub
Set tableantoine = DBJetMaster.OpenRecordset("Antoine List", dbOpenTable)
foundinant = 0

tableantoine.MoveFirst
Do While Not tableantoine.EOF
    If tableantoine("CAS") = Cur_Info.CAS Then
        foundinant = 1
        GoTo out1
    End If
    tableantoine.MoveNext
Loop
    
out1:
    If foundinant = 0 Then
        FRMAntoine!TXTErr1.caption = "ERROR!!"
        FRMAntoine!TXTErr1.Refresh
        tableantoine.Close
    Exit Sub
    Else
    foundinant = 0
    tableantoine.Edit
    tableantoine("ANTA") = A_ANT
    tableantoine("ANTB") = B_ANT
    tableantoine("ANTC") = C_ANT
    tableantoine("TMIN") = templow - 273.15
    tableantoine("TMAX") = temphigh - 273.15
    tableantoine("PUNIT") = "mm Hg"
    tableantoine("TUNIT") = "C"
    tableantoine.Update
    FRMAntoine!TXTErr1.caption = "Done"
    FRMAntoine!TXTErr1.Refresh
    End If
    tableantoine.Close
End Sub

Public Sub recalc_antoine()
    Dim newtemp As Double
    
    newtemp = CDbl(FRMAntoine!TXTANTTemp.Text)

    If FRMAntoine!TXTANTTempUnits.Text = "K" Then
        newtemp = newtemp - 273.15
    End If
    
    If FRMAntoine!TXTANTTempUnits.Text = "F" Then
        newtemp = ((newtemp + 459.67) / 1.8) - 273.15
    End If
        
        
    If FRMAntoine!TXTANTTempUnits.Text = "R" Then
        newtemp = (newtemp / 1.8) - 273.15
    End If
        
    tdft = newtemp
    from_recalc = 1
    
End Sub

Public Sub recalc_one_antoine()
    Dim newtemp As Double
    
    newtemp = CDbl(FRMAntoine!TXTTemp.Text)
    
    
    If FRMAntoine!TXTTempUnits.Text = "K" Then
        newtemp = newtemp - 273.15
    End If
    
    If FRMAntoine!TXTTempUnits.Text = "F" Then
        newtemp = ((newtemp + 459.67) / 1.8) - 273.15
    End If
        
        
    If FRMAntoine!TXTTempUnits.Text = "R" Then
        newtemp = (newtemp / 1.8) - 273.15
    End If
        
    tdft = newtemp
        
        
    
End Sub

Public Sub antoine_into_db()
Dim foundinant As Integer
Dim tableantoine As Recordset
' not sure we want to edit db, for now just exit
MsgBox ("Database not editable")
Exit Sub
foundinant = 0
tableantoine.MoveFirst
Do While Not tableantoine.EOF
    If tableantoine("CAS") = Cur_Info.CAS Then
        foundinant = 1
        GoTo out1
    End If
    tableantoine.MoveNext
Loop
    
out1:
    If foundinant = 0 Then
        FRMAntoine!TXTErr.caption = "ERROR!!"
        FRMAntoine!TXTErr.Refresh
    Exit Sub
    Else
    foundinant = 0
    tableantoine.Edit
    tableantoine("ANTA") = A_ANT
    tableantoine("ANTB") = B_ANT
    tableantoine("ANTC") = C_ANT
    tableantoine("TMIN") = templow - 273.15
    tableantoine("TMAX") = temphigh - 273.15
    tableantoine("PUNIT") = "mm Hg"
    tableantoine("TUNIT") = "C"
    tableantoine.Update
    FRMAntoine!TXTErr.caption = "Done"
    FRMAntoine!TXTErr.Refresh
    End If
End Sub

Public Sub do_antoine_convert()
Dim input_NL As Integer
Dim input_IUNIT As Integer
Dim input_JUNIT As Integer
Dim input_KUNIT As Integer
Dim input_IEQN As Integer
Dim input_NEQN As Integer
Dim input_PUNIT As Double
Dim input_TUNIT As Double
Dim input_OTLOW As Double
Dim input_OTHGH As Double
Dim input_AA As Double
Dim input_BB As Double
Dim input_CC As Double
Dim input_DD As Double
Dim input_EE As Double
Dim Counter As Integer
Dim temperature As Double

called = 2  'used for recalculate
Counter = 1
Do While ((Counter <= 10) And Not (Trim(FRMAntoine!CMBEquationNumber) = Trim(EqNum(Counter))))
    Counter = Counter + 1
Loop
input_NL = 0
If (Trim(FRMAntoine!CMBLogForm) = "Ln") Then
    input_IUNIT = 1
ElseIf (Trim(FRMAntoine!CMBLogForm) = "Log10") Then
    input_IUNIT = 2
End If
Select Case Trim(FRMAntoine!CMBVPUnits)
    Case "Pa"
        input_JUNIT = 1
    Case "atm"
        input_JUNIT = 2
    Case "bars"
        input_JUNIT = 3
    Case "mm Hg"
        input_JUNIT = 4
    Case "cm Hg"
        input_JUNIT = 5
    Case "m Hg"
        input_JUNIT = 6
    Case "in Hg"
        input_JUNIT = 7
    Case "in H2O"
        input_JUNIT = 8
    Case "ft H2O"
        input_JUNIT = 9
    Case "m H2O"
        input_JUNIT = 10
    Case "psia"
        input_JUNIT = 11
    Case "kPa"
        input_JUNIT = 12
End Select

Select Case Trim(FRMAntoine!CMBTempUnits)
    Case "K"
        input_KUNIT = 1
    Case "C"
        input_KUNIT = 2
    Case "F"
        input_KUNIT = 3
    Case "R"
        input_KUNIT = 4
End Select
    
input_IEQN = CInt(Left(FRMAntoine!CMBEquationNumber, 3))
If (input_IUNIT = 1) Then
    input_NEQN = 304
ElseIf (input_IUNIT = 2) Then
    input_NEQN = 300
End If
Select Case Trim(FRMAntoine!CMBEquationForm)
    Case "A - B / (T + C)"
        input_NEQN = input_NEQN
    Case "A - B / (T - C)"
        input_NEQN = input_NEQN + 1
    Case "A + B / (T - C)"
        input_NEQN = input_NEQN + 2
    Case "A + B / (T + C)"
        input_NEQN = input_NEQN + 3
End Select
input_PUNIT = 1 'Find out the difference between
input_TUNIT = 1 'these and JUNIT and KUNIT
input_OTLOW = MinT(Counter)
input_OTHGH = MaxT(Counter)
input_AA = CDbl(Trim(FRMAntoine!TXTStartingCoeffA))
input_BB = CDbl(Trim(FRMAntoine!TXTStartingCoeffB))
input_CC = CDbl(Trim(FRMAntoine!TXTStartingCoeffC))
input_DD = 0
input_EE = 0

temphigh = input_OTHGH
templow = input_OTLOW

NL = input_NL
AA = input_AA
BB = input_BB
CC = input_CC
DD = input_DD
EE = input_EE



Call CalcAntoine(1, 0, 500, 1000, 0.000001, input_IUNIT, input_JUNIT, input_KUNIT, input_NEQN, -999#, 999#, input_IEQN, input_PUNIT, input_TUNIT, input_OTLOW, input_OTHGH)
'Call CalcAntoine(1, 1, 0#, 25#, 500#, 100, 0.000001, input_NL, input_IUNIT, input_JUNIT, input_KUNIT, input_NEQN, -999#, 999#, input_IEQN, input_PUNIT, input_TUNIT, input_OTLOW, input_OTHGH, input_AA, input_BB, input_CC, input_DD, input_EE)

A_ANT = Format(A_ANT, "####0.0000")
B_ANT = Format(B_ANT, "####0.0000")
C_ANT = Format(C_ANT, "####0.0000")

FRMAntoine!TXTAA = A_ANT
FRMAntoine!TXTBB = B_ANT
FRMAntoine!TXTCC = C_ANT
VPANT = Format(VPANT, "######0.0000")
FRMAntoine!txtVP = VPANT
FRMAntoine!TXTEquationNumber = FRMAntoine!CMBEquationNumber
FRMAntoine!TXTEquation = FRMAntoine!TXTEquation
FRMAntoine!TXTTemp = Temp_Ant
FRMAntoine!TXTVPUnits = FRMAntoine!CMBVPUnits
FRMAntoine!TXTTempUnits = FRMAntoine!CMBTempUnits



End Sub

Public Sub antoine_equation_number()
Dim Counter As Integer
Dim Number As Integer
If (Trim(FRMAntoine!CMBEquationNumber) = "") Then
    Exit Sub
End If


Counter = 1
Do While ((Counter <= 10) And Not (Trim(FRMAntoine!CMBEquationNumber) = Trim(EqNum(Counter))))
    Counter = Counter + 1
Loop
Number = CInt(Left(FRMAntoine!CMBEquationNumber, 3))

FRMAntoine!TXTStartingCoeffA = CoeffA(Counter)
FRMAntoine!TXTStartingCoeffB = CoeffB(Counter)
FRMAntoine!TXTStartingCoeffC = CoeffC(Counter)
FRMAntoine!TXTStartingCoeffD = CoeffD(Counter)
FRMAntoine!TXTStartingCoeffE = CoeffE(Counter)


Call fill_antoine_input_defaults(MinT(Counter), MaxT(Counter), Number)
FRMAntoine!TXTEquation = fill_antoine_equation_box(EqNum(Counter))
End Sub

Public Sub antoine_antoine()
Dim DBTbl As Recordset
Dim Code As Integer
Dim Counter As Integer
Dim found As Boolean
Dim name As String
Dim i As Integer

On Error GoTo GetAntoineDataError
name = Cur_Info.name

' Find the CAS # associated with the chemical in order to search the 911 table
    Set DBTbl = DBJetMaster.OpenRecordset("Antoine List", dbOpenTable)
    
'    DBTbl.Index = "Name"
'    DBTbl.Seek "=", Trim(Name)
'
'    If DBTbl.NoMatch Then
'        MsgBox ("Chemical not found in Master Database")
'        DBTbl.Close
'        Exit Sub
'    End If

'    If (Trim(DBTbl("Name")) = Trim(Name)) Then
'        CAS_Number = DBTbl("CAS")
'    End If
'    DBTbl.Close
    
    DBTbl.MoveFirst
Do While (UCase(Trim(DBTbl("Name"))) <> UCase(Trim(name)))
    DBTbl.MoveNext
Loop

Counter = 1
found = False
    Do While (UCase(Trim(DBTbl("Name"))) = UCase(Trim(name)) And Counter <= Limit)
        EqNum(Counter) = DBTbl("Equation")
        If (EqNum(Counter) <> 0) Then
            If (DBTbl("Equation") >= 300 And DBTbl("Equation") <= 303) Then
                LogForm(Counter) = "Log10"
            Else
                LogForm(Counter) = "ln"
            End If
            If (DBTbl("Equation") = 300 Or DBTbl("Equation") = 304) Then
                EquationForm(Counter) = "A - B / (T + C)"
            ElseIf (DBTbl("Equation") = 301 Or DBTbl("Equation") = 305) Then
                EquationForm(Counter) = "A - B / (T - C)"
            ElseIf (DBTbl("Equation") = 302 Or DBTbl("Equation") = 306) Then
                EquationForm(Counter) = "A + B / (T - C)"
            ElseIf (DBTbl("Equation") = 303 Or DBTbl("Equation") = 307) Then
                EquationForm(Counter) = "A + B / (T + C)"
            End If
            found = True
'            units(Counter) = ""
'            On Error Resume Next
'            units(Counter) = DBTbl("Units")
'            On Error GoTo GetAntoineDataError
            MinT(Counter) = DBTbl("Tmin")
            MaxT(Counter) = DBTbl("Tmax")
            CoeffA(Counter) = DBTbl("A")
            CoeffB(Counter) = DBTbl("B")
            CoeffC(Counter) = DBTbl("C")
            CoeffD(Counter) = 0
            CoeffE(Counter) = 0
            TUnits(Counter) = DBTbl("TempUnits")
            VPUnits(Counter) = DBTbl("VPUnits")
            EqNum(Counter) = EqNum(Counter) & ", " & Counter
            ' All Equations will be 101, so to distinguish they will apear as (101,1  101,2  101,3  ......)
            Counter = Counter + 1
        End If
        DBTbl.MoveNext
    Loop
                   
    DBTbl.Close
On Error GoTo AntoineError

If (found = False) Then
    MsgBox ("No occurances of vapor pressure as an F(t) were found for " & Trim(name) & ".")
    Exit Sub
End If
    
FRMAntoine!TXTStartingCoeffA = CoeffA(1)
FRMAntoine!TXTStartingCoeffB = CoeffB(1)
FRMAntoine!TXTStartingCoeffC = CoeffC(1)
FRMAntoine!TXTStartingCoeffD = CoeffD(1)
FRMAntoine!TXTStartingCoeffE = CoeffE(1)

FRMAntoine!CMBEquationNumber = EqNum(1)
i = 1
Do While (Trim(EqNum(i)) <> "" And Trim(EqNum(i)) <> "0")
    FRMAntoine!CMBEquationNumber.AddItem Trim(EqNum(i))
    i = i + 1
Loop

FRMAntoine!TXTStartingCoeffA.Enabled = True
FRMAntoine!TXTStartingCoeffB.Enabled = True
FRMAntoine!TXTStartingCoeffC.Enabled = True
FRMAntoine!TXTStartingCoeffD.Enabled = False
FRMAntoine!TXTStartingCoeffE.Enabled = False
FRMAntoine!CMBEquationNumber.Enabled = True
FRMAntoine!TXTEquation.Enabled = True
FRMAntoine!LBLStartingCoeffA.Enabled = True
FRMAntoine!LBLStartingCoeffB.Enabled = True
FRMAntoine!LBLStartingCoeffC.Enabled = True
FRMAntoine!LBLStartingCoeffE.Enabled = False
FRMAntoine!LBLStartingCoeffD.Enabled = False

FRMAntoine!LBLEquationNumber.Enabled = True
FRMAntoine!LBLEquation.Enabled = True
FRMAntoine!VertScrollEquation.Enabled = True

FRMAntoine!CMBVPUnits = "mm Hg"
FRMAntoine!CMBTempUnits = TUnits(1)
FRMAntoine!CMBLogForm = LogForm(1)
FRMAntoine!CMBEquationForm = EquationForm(1)

Call fill_antoine_input_defaults(MinT(1), MaxT(1), CInt(Left(EqNum(1), 3)))
FRMAntoine!TXTEquation = fill_antoine_equation_box(EqNum(1))

FRMAntoine!CMDConvert.Enabled = True

FRMAntoine!CMDConvert.Enabled = True
FRMAntoine!CMDRegress.Enabled = False

Exit Sub

GetAntoineDataError:
   
    If Err = 94 Then Resume Next
    MsgBox "Error loading data from Antoine database", 48, "Error"
    DBTbl.Close
    Exit Sub

AntoineError:
    MsgBox ("Error with loading data to main form")
End Sub

Public Sub antoine_dippr()
Dim DBTbl As Recordset
Dim Cas_Number As Double
Dim Code As Integer
Dim Counter As Integer
Dim found As Boolean
Dim name As String
Dim i As Integer
   
On Error GoTo Get911DataError
name = FRMAntoine!TXTChemName.Text


' Find the CAS # associated with the chemical in order to search the 911 table
    Set DBTbl = DBJetMaster.OpenRecordset("Chemical Name", dbOpenTable)
    
'    DBTbl.Index = "Name"
'    DBTbl.Seek "=", Trim(Name)
'
'    If DBTbl.NoMatch Then
'        MsgBox ("Chemical not found in Master Database")
'        DBTbl.Close
'        Exit Sub
'    End If

'    If (Trim(DBTbl("Name")) = Trim(Name)) Then
'        CAS_Number = DBTbl("CAS")
'    End If
'    DBTbl.Close
    DBTbl.MoveFirst
Do While (UCase(Trim(DBTbl("Name"))) <> UCase(Trim(name)))
    DBTbl.MoveNext
Loop

If (UCase(Trim(DBTbl("Name"))) = UCase(Trim(name))) Then
    Cas_Number = DBTbl("CAS")
End If
DBTbl.Close
' Now Search the 911 table for the correct chemical
    Set DBTbl = DBJetMaster.OpenRecordset("DIPPR911", dbOpenTable)
    
    DBTbl.Index = "PrimaryKey1"
    DBTbl.Seek "=", Cas_Number
    
    If DBTbl.NoMatch Then
        MsgBox ("Chemical Not Found in Master Database")
        DBTbl.Close
        Exit Sub
    End If
Counter = 1
found = False
' Search all ocurrences of the chemical for PEARLS Code 6 (VP as a F(T))
    Do While (DBTbl("CAS #") = Cas_Number And Counter <= Limit)
        Code = DBTbl("PEARLS Code")
        If Code = 6 Then
            found = True
            MinT(Counter) = DBTbl("Value")
            units(Counter) = ""
            On Error Resume Next
            units(Counter) = DBTbl("Units")
            On Error GoTo Get911DataError
            MaxT(Counter) = DBTbl("Temperature")
            CoeffA(Counter) = DBTbl("Coef1")
            CoeffB(Counter) = DBTbl("Coef2")
            CoeffC(Counter) = DBTbl("Coef3")
            CoeffD(Counter) = DBTbl("Coef4")
            CoeffE(Counter) = DBTbl("Coef5")
            EqNum(Counter) = DBTbl("Equation")
            EqNum(Counter) = EqNum(Counter) & ", " & Counter
            ' All Equations will be 101, so to distinguish they will apear as (101,1  101,2  101,3  ......)
            Counter = Counter + 1
        End If
        DBTbl.MoveNext
    Loop
                   
    DBTbl.Close
On Error GoTo DIPPRError

If (found = False) Then
    MsgBox ("No occurances of vapor pressure as an F(t) where found for " & Trim(name) & ".")
    Exit Sub
End If
    
FRMAntoine!TXTStartingCoeffA = CoeffA(1)
FRMAntoine!TXTStartingCoeffB = CoeffB(1)
FRMAntoine!TXTStartingCoeffC = CoeffC(1)
FRMAntoine!TXTStartingCoeffD = CoeffD(1)
FRMAntoine!TXTStartingCoeffE = CoeffE(1)

FRMAntoine!CMBEquationNumber = EqNum(1)
i = 1
Do While (Trim(EqNum(i)) <> "")
    FRMAntoine!CMBEquationNumber.AddItem Trim(EqNum(i))
    i = i + 1
Loop

FRMAntoine!TXTStartingCoeffA.Enabled = True
FRMAntoine!TXTStartingCoeffB.Enabled = True
FRMAntoine!TXTStartingCoeffC.Enabled = True
FRMAntoine!TXTStartingCoeffD.Enabled = True
FRMAntoine!TXTStartingCoeffE.Enabled = True
FRMAntoine!CMBEquationNumber.Enabled = True
FRMAntoine!TXTEquation.Enabled = True
FRMAntoine!LBLStartingCoeffA.Enabled = True
FRMAntoine!LBLStartingCoeffB.Enabled = True
FRMAntoine!LBLStartingCoeffC.Enabled = True
FRMAntoine!LBLStartingCoeffD.Enabled = True
FRMAntoine!LBLStartingCoeffE.Enabled = True
FRMAntoine!LBLEquationNumber.Enabled = True
FRMAntoine!LBLEquation.Enabled = True
FRMAntoine!VertScrollEquation.Enabled = True

FRMAntoine!CMBVPUnits = "mm Hg"
FRMAntoine!CMBTempUnits = "C"
FRMAntoine!CMBLogForm = "Log10"
FRMAntoine!CMBEquationForm = "A - B / (T + C)"

Call fill_antoine_input_defaults(MinT(1), MaxT(1), 101)
FRMAntoine!TXTEquation = fill_antoine_equation_box(EqNum(1))

FRMAntoine!CMDConvert.Enabled = False
FRMAntoine!CMDRegress.Enabled = True

Exit Sub

Get911DataError:
   
    If Err = 94 Then Resume Next
    MsgBox "Error loading data from 911 database", 48, "Error"
    DBTbl.Close
    Exit Sub

DIPPRError:
    MsgBox ("Error with loading data to main form")
End Sub

Public Sub fill_antoine_input_defaults(TempFrom As Double, TempTo As Double, EqNum As Integer)
' TempFrom to TempTo is valid temperature range,
' EqNum is Equation # of inputs
Dim i As Integer

Pressure_Units(1) = "mm Hg"
Pressure_Units(2) = "Pa"
Pressure_Units(3) = "atm"
Pressure_Units(4) = "bars"
Pressure_Units(5) = "cm Hg"
Pressure_Units(6) = "m Hg"
Pressure_Units(7) = "in Hg"
Pressure_Units(8) = "in H2O"
Pressure_Units(9) = "ft H2O"
Pressure_Units(10) = "m H2O"
Pressure_Units(11) = "psia"
Pressure_Units(12) = "kPa"

Temperature_Units(1) = "C"
Temperature_Units(2) = "K"
Temperature_Units(3) = "F"
Temperature_Units(4) = "R"

FRMAntoine!LBLVPUnits.Enabled = True
FRMAntoine!LBLTempUnits.Enabled = True
FRMAntoine!LBLLogForm.Enabled = True
FRMAntoine!LBLEquationForm.Enabled = True
FRMAntoine!CMBVPUnits.Enabled = True
FRMAntoine!CMBTempUnits.Enabled = True
FRMAntoine!CMBLogForm.Enabled = True
FRMAntoine!CMBEquationForm.Enabled = True

FRMAntoine!CMBVPUnits.Clear     'remove old data
FRMAntoine!CMBTempUnits.Clear
FRMAntoine!CMBLogForm.Clear
FRMAntoine!CMBEquationForm.Clear


For i = 1 To 12 Step 1
    FRMAntoine!CMBVPUnits.AddItem Pressure_Units(i)
Next i

For i = 1 To 4 Step 1
    FRMAntoine!CMBTempUnits.AddItem Temperature_Units(i)
Next i

FRMAntoine!CMBLogForm.AddItem "Ln"
FRMAntoine!CMBLogForm.AddItem "Log10"

FRMAntoine!CMBEquationForm.AddItem "A - B / (T + C)"
FRMAntoine!CMBEquationForm.AddItem "A - B / (T - C)"
FRMAntoine!CMBEquationForm.AddItem "A + B / (T - C)"
FRMAntoine!CMBEquationForm.AddItem "A + B / (T + C)"

If (EqNum = 101) Then
    FRMAntoine!LBLTempRange.Enabled = True
    FRMAntoine!LBLRegPoints.Enabled = True
    FRMAntoine!LBLTempTo.Enabled = True
    FRMAntoine!TXTTempFrom.Enabled = True
    FRMAntoine!TXTTempTo.Enabled = True
    FRMAntoine!TXTRegressionPoints.Enabled = True
    FRMAntoine!CMDRegress.Enabled = True

    FRMAntoine!TXTTempFrom = TempFrom
    FRMAntoine!TXTTempTo = TempTo
    FRMAntoine!TXTRegressionPoints = 100
ElseIf (EqNum >= 300 And EqNum <= 307) Then
    
    FRMAntoine!TXTTempTo = TempTo
    FRMAntoine!TXTRegressionPoints = 100

End If

FRMAntoine!CMBVPUnits.ListIndex = 0
FRMAntoine!CMBTempUnits.ListIndex = 0
FRMAntoine!CMBLogForm.ListIndex = 0
FRMAntoine!CMBEquationForm.ListIndex = 0


End Sub


Public Function fill_antoine_equation_box(Number As String) As String
Dim EquationNumber As Integer

fill_antoine_equation_box = ""
' Converts the Equation Number string to a integer
' Number has a format of 101,1 .. 301, 2 ....
EquationNumber = CInt(Left(Number, 3))

Select Case EquationNumber
    Case 101
        fill_antoine_equation_box = "EXP(A + B/T + C(LN(T)) + DT^E)"
    Case 300
        fill_antoine_equation_box = "Log10P = A - B / (T + C)" & Chr(10) & Chr(10)
    Case 301
        fill_antoine_equation_box = "Log10P = A - B / (T - C)" & Chr(10) & Chr(10)
    Case 302
        fill_antoine_equation_box = "Log10P = A + B / (T - C)" & Chr(10) & Chr(10)
    Case 303
        fill_antoine_equation_box = "Log10P = A + B / (T + C)" & Chr(10) & Chr(10)
    Case 304
        fill_antoine_equation_box = "LnP = A - B / (T + C)" & Chr(10) & Chr(10)
    Case 305
        fill_antoine_equation_box = "LnP = A - B / (T - C)" & Chr(10) & Chr(10)
    Case 306
        fill_antoine_equation_box = "LnP = A + B / (T - C)" & Chr(10) & Chr(10)
    Case 307
        fill_antoine_equation_box = "LnP = A + B / (T + C)" & Chr(10) & Chr(10)
End Select

End Function

'********************************************************************
'mrt - Antoine stuff for printing 3/20/99
'********************************************************************

Sub antoine_check_update_udb(info As AntoineInfoType)
    'mrt 3/20/99
    'Check to see if calculations have been performed with antoine stuff
    'If calculations have been done, update dbuser database for printing
    
    If (Trim(FRMAntoine!TXTANTA.Text) = "" Or Trim(FRMAntoine!TXTANTA.Text) = "TXTANTA") And (Trim(FRMAntoine!TXTANTB.Text) = "" Or Trim(FRMAntoine!TXTANTB.Text) = "TXTANTB") And (Trim(FRMAntoine!TXTANTC.Text) = "" Or Trim(FRMAntoine!TXTANTC.Text) = "TXTANTC") Then
        If (Trim(FRMAntoine!TXTStartingCoeffA.Text) = "" Or Trim(FRMAntoine!TXTStartingCoeffA.Text) = "TXTStartingCoeffA") And (Trim(FRMAntoine!TXTStartingCoeffB.Text) = "" Or Trim(FRMAntoine!TXTStartingCoeffB.Text) = "TXTStartingCoeffB") And (Trim(FRMAntoine!TXTStartingCoeffC.Text) = "" Or Trim(FRMAntoine!TXTStartingCoeffC.Text) = "TXTStartingCoeffC") Then
            Exit Sub
        Else
            Call antoine_dbupdate(info)
        End If
    Else
        Call antoine_dbupdate_regress(info)
    End If
End Sub

Sub antoine_dbupdate(info As AntoineInfoType)
    'mrt - this sub updates dbuser for printing. It assumes that the user
    '       has already found coefficients, but hasn't done a regression.

    With info
        .AntCalc = True
        .AntType = True 'indicates there was no regression performed
        .A = FRMAntoine!TXTStartingCoeffA.Text
        .B = FRMAntoine!TXTStartingCoeffB.Text
        .C = FRMAntoine!TXTStartingCoeffC.Text
        .D = FRMAntoine!TXTStartingCoeffD.Text
        .E = FRMAntoine!TXTStartingCoeffE.Text
        .EqNum = FRMAntoine!CMBEquationNumber.Text
        .TMin = FRMAntoine!TXTTempFrom.Text
        .TMax = FRMAntoine!TXTTempTo.Text
        .TFT = ""
        .TFTUnit = FRMAntoine!CMBTempUnits.Text
        .value = ""
        .Unit = ""
    End With
End Sub

Sub antoine_dbupdate_regress(info As AntoineInfoType)
    'mrt - this sub updates dbuser for printing. It assumes that the user
    '       has found coefficients and has done a regression.

    With info
        .AntCalc = True
        .AntType = False 'indicates there was a regression performed
        .A = FRMAntoine!TXTANTA.Text
        .B = FRMAntoine!TXTANTB.Text
        .C = FRMAntoine!TXTANTC.Text
        .D = ""
        .E = ""
        .EqNum = FRMAntoine!TXTANTEqnNum.Text
        .TMin = FRMAntoine!TXTTempFrom.Text
        .TMax = FRMAntoine!TXTTempTo.Text
        .TFT = FRMAntoine!TXTANTTemp.Text
        .TFTUnit = FRMAntoine!TXTANTTempUnits.Text
        .value = FRMAntoine!TXTVaporPressure.caption
        .Unit = FRMAntoine!TXTANTPressUnits.Text
    End With
End Sub

Public Sub run_default_ant_calc()
    On Error Resume Next
    Call antoine_antoine
End Sub
