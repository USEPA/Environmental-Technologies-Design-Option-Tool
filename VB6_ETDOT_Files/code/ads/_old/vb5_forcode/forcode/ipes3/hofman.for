CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC
CC
CC  Program Name:       LIQUID-PHASE ADSORPTION ISOTHERM CALCULATOR
CC                      USING MANES/HOFER MODEL
CC  Author:             Michigan Tech University - 1994
CC  Intended Platform:  Compiled with Microsoft FORTRAN and linked
CC                      to the Visual Basic code of the Adsorption
CC                      Simulation Software.
CC
CC  Modification History:
CC  =====================
CC  03/11/1994: Fred Gobin (?)
CC  - Modified for DLL
CC  06/28/1996: Eric Oman
CC  - Added code to output LNCPTS and LNQPTS for output graph
CC    (Note maximum of 200 regression points.)
CC  09/02/1998: Eric Oman
CC  - Removed LNCPTS,LNQPTS,NUMPTS parameters.
CC
CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC

C---- List of Input and Output Parameters ------------------
C
C     BB     : Polanyi Parameter                    Input      
C     CBEG   : Correlation lower bound (ug/l)       Ouput
C     CBULK  : Bulk Concentration (ug/l)            I
C     CEND   : Correlation upper bound (ug/l)       O
C     CSAV   : Average bulk concentration (ug/l)    O
C     FWT    : Molecular weight (g/mol)             I
C     GM     : ????                                 I
C     NL     : Number of regression points (?)      I
C     OMAG   : Order of magnitude (?)               I
C     ORGDEN : Density g/cm3                        I
C     PVAP   : Vapor Pressure (mm Hg)               I
C     QSAV   : Polanyi Adsorption Capacity (ug/g)   O
C     RMSE   : Root Mean Square Error               O
C     RNDX   : Refractive Index                     I
C     RSQD   : Regression R-Squared                 O
C     SOLUB  : Solubility (mg/l)                    I
C     TT     : Temperature (Kelvin)                 I
C     XK1    : Freundlich K in  (ug/g)...           O
C     XK2    : Freundlich K in (umol/g)...          O
C     XNF    : Freundlich 1/n                       O
C     W0     : Polanyi parameter                    I
C     ----------------------------------------------------------
C     HOERR  : Error flag for HOFMAN
C     ERRMAT : Matrix of error flags
C     ----------------------------------------------------------
C     Note 1:
C     All Double Precision Input parameters are stored in the In_Data Array
C     Note 2:
C      The integers passed as parameters must be INTEGER*2 if declared as Integer
C      in Visual Basic app., or as INTEGER*4 if declared as Long in VB app.

C      SUBROUTINE HOFMAN(IN_DATA,NL,CSAV,QSAV,
C     &                  XK1,XK2,XNF,CBEG,CEND,RSQD,RMSE,
C     &                  ERRMAT,HOERR,LNCPTS,LNQPTS,NUMPTS)
      SUBROUTINE HOFMAN(IN_DATA,NL,CSAV,QSAV,
     &                  XK1,XK2,XNF,CBEG,CEND,RSQD,RMSE,
     &                  ERRMAT,HOERR)
      IMPLICIT NONE
C
C------ PARAMETERS FOR HOFMAN CALCULATION MODULE.
C
C...INPUTS...:
      DOUBLE PRECISION IN_DATA(11)
      INTEGER*2 NL
C...OUTPUTS...:
      DOUBLE PRECISION CSAV
      DOUBLE PRECISION QSAV
      DOUBLE PRECISION XK1
      DOUBLE PRECISION XK2
      DOUBLE PRECISION XNF
      DOUBLE PRECISION CBEG
      DOUBLE PRECISION CEND
      DOUBLE PRECISION RSQD
      DOUBLE PRECISION RMSE
      INTEGER*2 ERRMAT(30)
      INTEGER*2 HOERR
C
C------ LOCAL VARIABLES.
C
      DOUBLE PRECISION ADSP,ADSPRF,B0,B1,BB,CBULK,CINC,
     &                 CLNC,CONC,CS,CSS,CZERO,DENS,DIFF,DP
      DOUBLE PRECISION FWT,GAMMA,GM,ORGDEN,OMAG,OREF,PS,PVAP
      DOUBLE PRECISION QCAP,QCAL
      DOUBLE PRECISION QLNQ,REFDEN,REFMW,REFVOL,RGAS,RHOM,RIREF
      DOUBLE PRECISION RNDX,
     &                 SOLUB,SUMX,SUMY,SUMXX,SUMYY,SUMXY,
     &                 TT,VOLAD,VOLM,W0

      INTEGER I,J,K,NJ
      INTEGER*2 ERRNUM

C      DOUBLE PRECISION LNCPTS(200),LNQPTS(200)
C      INTEGER*2 NUMPTS


C    -- INITIALIZE VARIABLES
      ERRNUM = 0 
      DO I=1, 30
        ERRMAT(I) = 0
      ENDDO

      BB    =IN_DATA(1)
      W0    = IN_DATA(2)
      TT    = IN_DATA(3)
      CBULK = IN_DATA(4)
      ORGDEN= IN_DATA(5)
      FWT   = IN_DATA(6)
      PVAP  = IN_DATA(7)
      SOLUB = IN_DATA(8)
      RNDX  = IN_DATA(9)
      GM    = IN_DATA(10)
      OMAG  = IN_DATA(11)

      CSAV=0
      QSAV=0
      SUMX=0
      SUMY=0
      SUMXX=0
      SUMYY=0
      SUMXY=0
      
      DIFF = SOLUB-(CBULK/1000)
      
      IF (DIFF.LE.0) THEN

           HOERR = -1
           CALL ERROR (ERRMAT,ERRNUM,11)
           GOTO 9999

      END IF

      DENS = ORGDEN

C    -- CS AND CZERO IN {uM/L}

      CZERO = CBULK/FWT
      CS = (SOLUB*1000)/FWT
      VOLM = FWT/DENS
      RHOM = (DENS*1.0E+06)/FWT
      RGAS = 1.987
      NL = NL+1
      NJ = NL/2
      NL = 2*NJ
      DP = DBLE(NL)
      OMAG = OMAG/2
      CBEG = DLOG(CZERO/10**OMAG)
      CEND = DLOG(CZERO*10**OMAG)
      DIFF = CS-DEXP(CEND)

      IF (DIFF.LE.0) THEN

           CEND = DLOG(0.99*CS)
           CALL ERROR (ERRMAT,ERRNUM,12)

      END IF

      IF (CBEG.GE.CEND) THEN

           HOERR = -1
           CALL ERROR (ERRMAT,ERRNUM,13)
           GOTO 9999

      END IF

      CINC = (CEND-CBEG)/DP

      PS = PVAP/760

      CSS = (PS/0.08206/TT)*1.0E+06
      DIFF = (CSS*FWT)-CBULK

      IF (DIFF.LE.0) THEN

           HOERR = -1
           CALL ERROR (ERRMAT,ERRNUM,18)
           GOTO 9999

      END IF

C    -- GAS PHASE D-R CORRELATION

C    -- REFERENCE CHEMICAL: TOLUENE

C    --               MW = 92.140 {g/gmol}
C    --          DENSITY = 0.8623 {g/cc}
C    -- REFRACTIVE INDEX = 1.4941

      REFMW = 92.14
      REFDEN = 0.8623
      REFVOL = REFMW/REFDEN
      RIREF = 1.4941
      OREF = (RIREF**2-1)/(RIREF**2+2)

C    -- CORRELATING DIVISOR FOR WATER VAPOR ISOTHERM IS 0.28

      GAMMA = ((RNDX**2-1)/(RNDX**2+2))/OREF-0.28

      DO 10 K=1,NL

           CLNC = CBEG+DBLE(K)*CINC
           CONC = DEXP(CLNC)
           ADSP = (RGAS*TT)*DLOG(CS/CONC)
           ADSPRF = (ADSP/VOLM)*REFVOL/GAMMA

C    -- CHECK FOR DEXP BOUNDARY

           IF ((-BB*(ADSPRF)**GM).LT.-710) THEN

                HOERR = -1
                CALL ERROR (ERRMAT,ERRNUM,14)
                GOTO 9999

           END IF

           VOLAD = W0*DEXP(-BB*(ADSPRF)**GM)
           QCAP = RHOM*VOLAD

           IF (QCAP.LE.1.0E-03) THEN
       
                HOERR = -1
                CALL ERROR (ERRMAT,ERRNUM,21)
                GOTO 9999

           END IF
 
           QLNQ = DLOG(QCAP)
Cc
Cc Code added on 6/28/96 to output ln(C) and ln(Q) for analysis.
Cc
C           LNCPTS(K) = CLNC
C           LNQPTS(K) = QLNQ
Cc
           SUMX = SUMX+CLNC
           SUMY = SUMY+QLNQ
           SUMXX = SUMXX+(CLNC)**2
           SUMYY = SUMYY+(QLNQ)**2
           SUMXY = SUMXY+(CLNC*QLNQ)
        
           IF (K.EQ.NJ) THEN

                CSAV = CONC*FWT
                QSAV = QCAP*FWT
    
           END IF

  10  CONTINUE

Cc
Cc Code added on 6/28/96 to output ln(C) and ln(Q) for analysis.
Cc
C      IF (NL .LT. 200) THEN
C           LNCPTS(NL+1) = 0D0
C           LNQPTS(NL+1) = 0D0
C           NUMPTS = NL
C      ELSE
C           NUMPTS = 200
C      END IF
Cc

C    -- CALCULATE FREUNDLICH "K" AND 1/n BY LINEAR REGRESSION
C    -- FREUNDLICH "K" (XK1) IS IN {(ug/g)(L/ug)**(1/n)}

      B0 = (SUMY/DP)-(DP*SUMX*SUMXY-SUMX**2*SUMY)/DP/(DP*SUMXX-SUMX**2)
      B1 = (DP*SUMXY-SUMX*SUMY)/(DP*SUMXX-SUMX**2)
      XNF = B1
      XK2 = DEXP(B0)
      XK1 = (XK2*FWT)*(1/FWT)**XNF
      RSQD = 1-(SUMYY-B0*SUMY-B1*SUMXY)/((DP*SUMYY-SUMY**2)/DP)

C    -- CALCULATE ROOT-MEAN-SQUARE ERROR (RMSE)

      RMSE = 0

      DO 20 J=1,NL

           CLNC = CBEG+DBLE(K)*CINC
           CONC = DEXP(CLNC)
           ADSP = (RGAS*TT)*DLOG(CS/CONC)
           ADSPRF = (ADSP/VOLM)*REFVOL/GAMMA
           VOLAD = W0*DEXP(-BB*(ADSPRF)**GM)
           QCAP = RHOM*VOLAD
           QCAL = XK2*(CONC)**XNF
           RMSE = RMSE+((QCAL-QCAP)/QCAP)**2

  20  CONTINUE

      RMSE = DSQRT(RMSE/DP)**100
      CBEG = DEXP(CBEG)*FWT
      CEND = DEXP(CEND)*FWT
 9999 CONTINUE
      RETURN
      END
