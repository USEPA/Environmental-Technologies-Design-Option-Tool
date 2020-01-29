CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC
C
C     MODULE NAME: ADLIQ()
C     LONGER NAME: LIQUID-PHASE ADSORPTION ISOTHERM CALCULATOR
C
C     AUTHOR: Michigan Tech University - 1994
C
C     PLATFORM: DIGITAL VISUAL FORTRAN V6.0
C
C     HISTORY:
C         1994-MAR-11 - FRED GOBIN - MODIFIED FOR DLL
C         1996-JUN-28 - OMAN - ADDED CODE FOR LNCPTS/LNQPTS GRAPH (200 POINTS MAX)
C         1998-SEP-02 - OMAN - REMOVED LNCPTS/LNQPTS GRAPH
C         1999-MAY-14 - OMAN - MODIFIED ADLIQ() ROUTINE TO USE
C           MOLAR VOLUME @ NBP (VOLM_NBP), INSTEAD OF USING MOLAR VOLUME
C           @ OPERATING TEMPERATURE (VOLM=FWT/DENS).
C         1999-MAY-20 - OMAN - MODIFIED ADLIQ() ROUTINE TO USE
C           MOLAR VOLUME @ OPERATING TEMPERATURE (VOLM), INSTEAD OF 
C           USING MOLAR VOLUME @ NBP (VOLM_NBP).
C
CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC
C                                                                              *
C234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567
C-----789012345678901234567890123456789012345678901234567890123456789012CCCCCCCC-----------------

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
C     QSAV   : Polanyi Adsorption Capacity (ug/g)   O
C     RMSE   : Root Mean Square Error               O
C     RSQD   : Regression R-Squared                 O
C     SOLUB  : Solubility (mg/l)                    I
C     TT     : Temperature (Kelvin)                 I
C     VOLM_NBP: Molar volume @ normal boiling pt.   I
C     W0     : Polanyi parameter                    I
C     XK1    : Freundlich K in  (ug/g)...           O
C     XK2    : Freundlich K in (umol/g)...          O
C     XNF    : Freundlich 1/n                       O
C     ----------------------------------------------------------
C     ALERR  : Error flag for ADLIQ                 O
C     ERRMAT : Matrix of error flags                O
C     ----------------------------------------------------------
C     Note :
C      The integers passed as parameters must be INTEGER*2 if declared as Integer
C      in Visual Basic app., or as INTEGER*4 if declared as Long in VB app.

cC      SUBROUTINE ADLIQ (BB,W0,GM,CBULK,ORGDEN,TT,FWT,SOLUB,NL,OMAG,
cC     &                  CSAV,QSAV,XK1,XK2,XNF,CBEG,CEND,RSQD,RMSE,
cC     &                  ERRMAT,ALERR,LNCPTS,LNQPTS,NUMPTS)
c      SUBROUTINE ADLIQ (BB,W0,GM,CBULK,ORGDEN,TT,FWT,SOLUB,NL,OMAG,
c     &                  CSAV,QSAV,XK1,XK2,XNF,CBEG,CEND,RSQD,RMSE,
c     &                  ERRMAT,ALERR)
      SUBROUTINE ADLIQ (BB,W0,GM,CBULK,ORGDEN,TT,FWT,SOLUB,NL,OMAG,
     &                  VOLM_NBP,
     &                  CSAV,QSAV,XK1,XK2,XNF,CBEG,CEND,RSQD,RMSE,
     &                  ERRMAT,ALERR)

      IMPLICIT NONE
C
C------ PARAMETERS FOR ADLIQ CALCULATION MODULE.
C
      DOUBLE PRECISION BB
      DOUBLE PRECISION W0
      DOUBLE PRECISION GM
      DOUBLE PRECISION CBULK
      DOUBLE PRECISION ORGDEN
      DOUBLE PRECISION TT
      DOUBLE PRECISION FWT
      DOUBLE PRECISION SOLUB
      INTEGER*2 NL
      DOUBLE PRECISION OMAG
	DOUBLE PRECISION VOLM_NBP
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
      INTEGER*2 ALERR
C      DOUBLE PRECISION LNCPTS(200)
C      DOUBLE PRECISION LNQPTS(200)
C      INTEGER*2 NUMPTS
C
C------ LOCAL VARIABLES.
C
      DOUBLE PRECISION ADSP,B0,B1,CINC,CLNC,CONC,CS,CZERO,DENS,DIFF,
     &                 DP,QLNQ,QCAL,QCAP,RGAS,RHOM,SUMX,SUMY,SUMYY,
     &                 SUMXX,SUMXY,VOLM
      INTEGER*2 ERRNUM
      INTEGER J,K,NJ
      INTEGER I
C      COMMON /INITIAL/ W0,BB,GM,CBULK,ORGDEN,OMAG,NL
C      COMMON /ADSORB/ CSAV,QSAV,XK1,XK2,XNF,CBEG,CEND,RSQD,RMSE
C      COMMON /ERR/ ERRMAT(30),ERRNUM

      ERRNUM = 0
      DO 14 I=1,30
        ERRMAT(I) = 0
  14  CONTINUE


C    -- INITIALIZE VARIABLES
      DIFF = SOLUB-(CBULK/1000)

      IF (DIFF.LE.0) THEN
        ALERR = -1
        CALL ERROR (ERRMAT,ERRNUM,11)
        RETURN
      END IF

      DENS = ORGDEN

C    -- POLANYI GENERALIZED ISOTHERM 
C    -- CS AND CZERO IN {umol/L}  

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
        ALERR = -1
        CALL ERROR (ERRMAT,ERRNUM,13)
        RETURN 
      END IF

      CINC = (CEND-CBEG)/DP
      CSAV = 0
      QSAV = 0
      SUMX = 0
      SUMY = 0
      SUMXX = 0
      SUMYY = 0
      SUMXY = 0

      DO 20 K=1,NL
        CLNC = CBEG+DBLE(K)*CINC
        CONC = DEXP(CLNC)
        ADSP = (RGAS*TT)*DLOG(CS/CONC)

C------ MODIFIED BY OMAN, 1999-MAY-20, NEW CODE BEGINS:
        QCAP = (RHOM*W0)*DEXP(-BB*(ADSP/VOLM)**GM)
C------ MODIFIED BY OMAN, 1999-MAY-20, NEW CODE ENDS.
C------ MODIFIED BY OMAN, 1999-MAY-20, OLD CODE BEGINS:
c  C------ MODIFIED BY OMAN, 1999-MAY-14, NEW CODE BEGINS:
c          QCAP = (RHOM*W0)*DEXP(-BB*(ADSP/VOLM_NBP)**GM)
c  C------ MODIFIED BY OMAN, 1999-MAY-14, NEW CODE ENDS.
c  C------ MODIFIED BY OMAN, 1999-MAY-14, OLD CODE BEGINS:
c  C        QCAP = (RHOM*W0)*DEXP(-BB*(ADSP/VOLM)**GM)
c  C------ MODIFIED BY OMAN, 1999-MAY-14, OLD CODE ENDS.
C------ MODIFIED BY OMAN, 1999-MAY-20, OLD CODE ENDS.

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
  20  CONTINUE
Cc
Cc Code added on 6/28/96 to output ln(C) and ln(Q) for analysis.
Cc
C      IF (NL .LT. 200) THEN
C        LNCPTS(NL+1) = 0D0
C        LNQPTS(NL+1) = 0D0
C        NUMPTS = NL
C      ELSE
C        NUMPTS = 200
C      END IF
Cc

C    -- CALCULATE FREUNDLICH "K" AND "1/n" BY LINEAR REGRESSION 
C    -- FREUNDLICH "K" (XK1) IS IN {ug/gm} OR {(L/ug)**(1/n)}

      B0 = (SUMY/DP)-(DP*SUMX*SUMXY-SUMX**2*SUMY)/DP/(DP*SUMXX-SUMX**2)
      B1 = (DP*SUMXY-SUMX*SUMY)/(DP*SUMXX-SUMX**2)
      XNF = B1
      XK2 = DEXP(B0)
      XK1 = (XK2*FWT)*(1/FWT)**XNF
      RSQD = 1-(SUMYY-B0*SUMY-B1*SUMXY)/((DP*SUMYY-SUMY**2)/DP)

C    -- CALCULATE THE ROOT MEAN SQUARE ERROR (RMSE) 

      RMSE = 0

      DO 30 J=1,NL
        CLNC = CBEG+DBLE(J)*CINC
        CONC = DEXP(CLNC)
        ADSP = (RGAS*TT)*DLOG(CS/CONC)

C------ MODIFIED BY OMAN, 1999-MAY-20, NEW CODE BEGINS:
        QCAP = (RHOM*W0)*DEXP(-BB*(ADSP/VOLM)**GM)
C------ MODIFIED BY OMAN, 1999-MAY-20, NEW CODE ENDS.
C------ MODIFIED BY OMAN, 1999-MAY-20, OLD CODE BEGINS:
c C------ MODIFIED BY OMAN, 1999-MAY-14, NEW CODE BEGINS:
c         QCAP = (RHOM*W0)*DEXP(-BB*(ADSP/VOLM_NBP)**GM)
c C------ MODIFIED BY OMAN, 1999-MAY-14, NEW CODE ENDS.
c C------ MODIFIED BY OMAN, 1999-MAY-14, OLD CODE BEGINS:
c C        QCAP = (RHOM*W0)*DEXP(-BB*(ADSP/VOLM)**GM)
c C------ MODIFIED BY OMAN, 1999-MAY-14, OLD CODE ENDS.
C------ MODIFIED BY OMAN, 1999-MAY-20, OLD CODE ENDS.

        QCAL = XK2*(CONC)**XNF
        RMSE = RMSE+((QCAL-QCAP)/QCAP)**2
  30  CONTINUE

      RMSE = DSQRT(RMSE/DP)*100
      CBEG = DEXP(CBEG)*FWT
      CEND = DEXP(CEND)*FWT
C      OPEN(UNIT=12,FILE='check2.txt')
C      WRITE(12,*) 'RMSE=',RMSE
C      WRITE(12,*) 'XK1=',XK1
C      WRITE(12,*) 'XK2=',XK2
C      CLOSE(12)
 9999 CONTINUE
      RETURN
      END


