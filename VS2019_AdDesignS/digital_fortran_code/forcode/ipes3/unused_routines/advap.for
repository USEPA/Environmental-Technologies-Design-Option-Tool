CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC
CC
CC  Program Name:       GAS-PHASE ADSORPTION ISOTHERM CALCULATOR
CC  Author:             Michigan Tech University - 1994
CC  Intended Platform:  Compiled with Microsoft FORTRAN and linked
CC                      to the Visual Basic code of the Adsorption
CC                      Simulation Software.
CC
CC  Modification History:
CC  =====================
CC  03/15/1994: Fred Gobin (?)
CC  - Modified for DLL
CC  03/25/1996: Eric Oman
CC  - Added code to output LNCPTS and LNQPTS for output graph
CC    (Note maximum of 200 regression points.)
CC  - Also added code to allow calling program to request certain
CC    data (via DATRQT and DATMAT).
CC
CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC

C---- List of Input and Output Parameters ------------------
C
C     BB     : Polanyi Parameter                    Input
C     CBULK  : Bulk Concentration (ug/l)            I
C     CBEG   : Correlation lower bound (ug/l)       Output
C     CEND   : Correlation upper bound (ug/l)       O
C     CSAV   : Average bulk concentration (ug/l)    O
C     DATMAT : Matrix to store returned data        O
C     DATRQT : Requested data to return in DATMAT   I 
C              0 -- Nothing
C              1 -- Element 1: QCAP (w, cm3/g)
C                   Element 2: ADSP (epsilon, cal/gmol)
C                   Element 3: PI (Pi, atm)
C     FWT    : Molecular weight (g/mol)             I
C     GM     : ????                                 I
C     IMOD   : Integer Flag for Equation            I
C     LNCPTS : Matrix of ln(C) points               Output
C     LNQPTS :
C     NL     : Number of regression points (?)      I
C     NUMPTS : Number of regression points          Output
C     OMAG   : Order of magnitude (?)               I
C     ORGDEN : Density g/cm3                        I
C     PVAP   : Vapor Pressure (mm Hg)               I
C     QSAV   : Polanyi Adsorption Capacity (ug/g)   O
C     RMSE   : Root Mean Square Error               O
C     RNDX   : Refractive Index                     I
C     RSQD   : Regression R-Squared                 O
C     RELHUM : Relative Humidity   (??)             I
C     TT     : Temperature (Kelvin)                 I
C     XK1    : Freundlich K in  (ug/g)...           O
C     XK2    : Freundlich K in (umol/g)...          O
C     XNF    : Freundlich 1/n                       O
C     W0     : Polanyi parameter (cm3/g) (??)       I
C     ----------------------------------------------------------
C     AVERR  : Error flag for ADVAP                 O
C     ERRMAT : Matrix of error flags                O
C     ----------------------------------------------------------
C     Note 1:
C     All Double Precision Input parameters are stored in the In_Data Array
C     Note 2:
C     The integers passed as parameters must be INTEGER*2 if declared as Integer
C     in Visual Basic app., or as INTEGER*4 if declared as Long in VB app.

      SUBROUTINE ADVAP (IN_DATA,IMOD,NL,
     &                  CSAV,QSAV,XK1,XK2,XNF,CBEG,CEND,RSQD,RMSE,
     &                  ERRMAT,AVERR,LNCPTS,LNQPTS,NUMPTS,
     &                  DATRQT,DATMAT)

      IMPLICIT NONE

      DOUBLE PRECISION IN_DATA(11),RELHUM,PVAP,TT,FWT,RNDX
      DOUBLE PRECISION W0,BB,CBULK,ORGDEN,GM,OMAG
      DOUBLE PRECISION CSAV,QSAV,XK1,XK2,XNF,CBEG,CEND,RSQD,RMSE
      DOUBLE PRECISION LNCPTS(200),LNQPTS(200)
      INTEGER*2 NUMPTS
      INTEGER*2 DATRQT
      DOUBLE PRECISION DATMAT(6,200)

      DOUBLE PRECISION SUMX,SUMY,SUMXX,SUMYY,SUMXY,PS,CS,DENS,BETA, 
     &                 DIFF,REFMW1,REFMW2,REFDN1,REFDN2,RIREF1,RIREF2,
     &                 OREF1,OREF2,SF1,SF2,CZERO,VOLM,VH2O,RHOM,RGAS,
     &                 DP,CINC,CLNC,CONC,PI,ADSP,QCAP,EL,YORD,XORD,
     &                 QLOG,QQ,QLNQ,RSAV,B0,B1,QCAL,RLIM,RATIO

      INTEGER*2 AVERR,ERRMAT(30),ERRNUM,IMOD,NL
      INTEGER K,J,NJ,NM

C      COMMON /INITIAL/ W0,BB,GM,CBULK,ORGDEN,OMAG,NL
C      COMMON /ADSORB/ CSAV,QSAV,XK1,XK2,XNF,CBEG,CEND,RSQD,RMSE
C      COMMON /ERR/ ERRMAT(30),ERRNUM

      INTEGER I

C    -- INITIALIZE VARIABLES
      BB    = IN_DATA(1)
      W0    = IN_DATA(2)
      TT    = IN_DATA(3)
      CBULK = IN_DATA(4)
      ORGDEN= IN_DATA(5)
      FWT   = IN_DATA(6)
      PVAP  = IN_DATA(7)
      RELHUM= IN_DATA(8)
      RNDX  = IN_DATA(9)
      GM    = IN_DATA(10)
      OMAG  = IN_DATA(11)
 
      CSAV = 0
      QSAV = 0
      SUMX = 0
      SUMY = 0
      SUMXX = 0
      SUMYY = 0
      SUMXY = 0
      RMSE = 0

      ERRNUM = 0
      DO 14 I=1,30
        ERRMAT(I)=0
  14  CONTINUE


C    -- PARAMETERS IN D-R EQUATION 

      NM = IDINT(GM)
      PS = PVAP/760
          CS = (PS/0.08206/TT)*1.0E+06
      DENS = ORGDEN
      BETA = 0.66
      DIFF = (CS*FWT)-CBULK

      IF (DIFF.LE.0) THEN

	   AVERR = -1
	   CALL ERROR (ERRMAT,ERRNUM,11)
	   GOTO 9999

      END IF

      IF (RELHUM.LT.60) THEN

	   RELHUM = 100
	   BETA = 1

      END IF

C    -- POLANYI GENERALIZED ISOTHERM 

C    -- THE EQUILIBRIUM ADSORBED CHEMICAL CONCENTRATION IS 
C    -- ESTIMATED FOR A CHEMICAL-AIR-WATER MIXTURE 

C    -- CALCULATE REFRACTIVE INDEX "SCALE FACTORS" 

C    -- REFERENCE CHEMICAL : TOLUENE [1] 

C    --               MW = 92.140 {g/gmol} 
C    --          DENSITY = 0.8623 {g/cc} 
C    -- REFRACTIVE INDEX = 1.4941 

C    -- REFERENCE CHEMICAL : N-HEPTANE [2] 

C    --               MW = 100.19 {g/gmol} 
C    --          DENSITY = 0.6795 {g/cc} 
C    -- REFRACTIVE INDEX = 1.3851 

      REFMW1 = 92.14
      REFMW2 = 100.19
      REFDN1 = 0.8623
      REFDN2 = 0.6795
      RIREF1 = 1.4941
      RIREF2 = 1.3851
      OREF1 = (REFMW1/REFDN1)*(RIREF1**2-1)/(RIREF1**2+2)
      OREF2 = (REFMW2/REFDN2)*(RIREF2**2-1)/(RIREF2**2+2)
      SF1 = (FWT/DENS)*(RNDX**2-1)/(RNDX**2+2)/OREF1
      SF2 = (FWT/DENS)*(RNDX**2-1)/(RNDX**2+2)/OREF2
    

C    -- CZERO IS IN {umol/L} AND CBULK IS IN {ug/L} 

      CZERO = CBULK/FWT
      VOLM = FWT/DENS
      VH2O = 18.015
      RHOM = (DENS/FWT)*1.0E+06
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

	   AVERR = -1
	   CALL ERROR (ERRMAT,ERRNUM,13)
	   GOTO 9999

      END IF

      CINC = (CEND-CBEG)/DP

      DO 20 K=1,NL

	   CLNC = CBEG+DBLE(K)*CINC
	   CONC = DEXP(CLNC)
	   PI = (CONC*1.0E-06)*(0.08206*TT)

	   IF (IMOD.EQ.2) THEN

C    --  DUBININ-RADUSHKEVICH (D-R) EQUATION 

		ADSP = (RGAS*TT)*DLOG(PS/PI)

C    -- THIS IF WAS ADDED TO CHECK FOR BOUNDARY OF DEXP FUNCTION 

		IF ((-BB*(ADSP/SF1)**NM).LT.-710) THEN

		     AVERR = -1
		     CALL ERROR (ERRMAT,ERRNUM,14)
		     GOTO 9999

		END IF

		QCAP = (RHOM*W0)*DEXP(-BB*(ADSP/SF1)**NM)
		
	   ELSE

C    --  CALGON CHARACTERISTIC EQUATION (BPL) 

		EL = DLOG10(PS/PI)
		YORD = (TT/BETA)*(EL/VOLM-DLOG10(100/RELHUM)/VH2O)
		XORD = YORD/SF2

		QLOG = 1.71 - (1.46E-02*XORD)
     &                 - (1.65E-03*XORD**2)
     &                 - (4.11E-04*XORD**3)
     &                 + (3.14E-05*XORD**4)
     &                 - (6.75E-07*XORD**5)

C    -- THIS IF WAS ADDED TO CHECK FOR EXPONENT BOUNDARY 

		IF (QLOG.LT.-307) THEN

		     AVERR = -1
		     CALL ERROR (ERRMAT,ERRNUM,15)
		     GOTO 9999

		END IF

		QQ = (10**QLOG)/100
		QCAP = (QQ*DENS/FWT)*1.0E+06

	   END IF

	   IF (QCAP.LE.1.0E-03) THEN

		AVERR = -1
		CALL ERROR (ERRMAT,ERRNUM,16)
		GOTO 9999

	   END IF

	   QLNQ = DLOG(QCAP)

c
c Code added on 3/25/96 to output ln(C) and ln(Q) for analysis.
c
           LNCPTS(K) = CLNC
           LNQPTS(K) = QLNQ
c

c
c Code added on 3/25/96 to output various variables for analysis.
c
           IF (DATRQT.EQ.1) THEN
             DATMAT(1,K) = QCAP
             DATMAT(2,K) = ADSP
             DATMAT(3,K) = PI
           END IF
c

           SUMX = SUMX+CLNC
	   SUMY = SUMY+QLNQ
	   SUMXX = SUMXX+(CLNC)**2
	   SUMYY = SUMYY+(QLNQ)**2
	   SUMXY = SUMXY+(CLNC*QLNQ)

	   IF (K.EQ.NJ) THEN

		CSAV = CONC*FWT
		QSAV = QCAP*FWT
		RSAV = PI/PS

	   END IF

  20  CONTINUE

c
c Code added on 3/25/96 to output ln(C) and ln(Q) for analysis.
c
      IF (NL .LT. 200) THEN
           LNCPTS(NL+1) = 0D0
           LNQPTS(NL+1) = 0D0
           NUMPTS = NL
      ELSE
           NUMPTS = 200
      END IF
c

C    -- CALCULATE FREUNDLICH "K" AND "1/n"  BY LINEAR REGRESSION
C    -- FREUNDLICH "K" (XK1) IS IN {ug/gm} OR {(L/ug)**(1/n)} 

      B0 = (SUMY/DP)-(DP*SUMX*SUMXY-SUMX**2*SUMY)/DP/(DP*SUMXX-SUMX**2)
      B1 = (DP*SUMXY-SUMX*SUMY)/(DP*SUMXX-SUMX**2)
      XNF = B1
      XK2 = DEXP(B0)
      XK1 = (XK2*FWT)*(1/FWT)**XNF
      RSQD = 1-(SUMYY-B0*SUMY-B1*SUMXY)/((DP*SUMYY-SUMY**2)/DP)

C    -- CALCULATE THE ROOT MEAN SQUARE ERROR (RMSE) 
	 
      DO 30 J=1,NL

	   CLNC = CBEG+DBLE(J)*CINC
	   CONC = DEXP(CLNC)
	   PI = (CONC*1.0E-06)*(0.08206*TT)

	   IF (IMOD.EQ.2) THEN

C    -- DUBININ-RADUSHKEVICH (D-R) EQUATION 

		ADSP = (RGAS*TT)*DLOG(PS/PI)
		QCAP = (RHOM*W0)*DEXP(-BB*(ADSP/SF1)**NM)
 
	   ELSE

C    -- CALGON CHARACTERISTIC EQUATION (BPL) 

		EL = DLOG10(PS/PI)
		YORD = (TT/BETA)*(EL/VOLM-DLOG10(100/RELHUM)/VH2O)
		XORD = YORD/SF2
		QLOG = 1.71 - (1.46E-02*XORD)
     &                 - (1.65E-03*XORD**2)
     &                 - (4.11E-04*XORD**3)
     &                 + (3.14E-05*XORD**4)
     &                 - (6.75E-07*XORD**5)
		QQ = (10**QLOG)/100
		QCAP = (QQ*DENS/FWT)*1.0E+06

	   END IF

	   QCAL = XK2*(CONC)**XNF
	   RMSE = RMSE+((QCAL-QCAP)/QCAP)**2

  30  CONTINUE

      RMSE = DSQRT(RMSE/DP)*100
      CBEG = DEXP(CBEG)*FWT
      CEND = DEXP(CEND)*FWT
      CS = (CS*FWT)/1000

C    -- WARN USER IF IN "PORE FILLING" REGIME {(Pi/Ps) > 0.2}

      RLIM = 0.2
      RATIO = PI/PS

      IF ((RSAV.GE.RLIM).OR.(RATIO.GE.RLIM)) THEN

	   CALL ERROR (ERRMAT,ERRNUM,17)
	   GOTO 9999

      END IF

 9999 OPEN(7,FILE="PIPO.TXT")
      WRITE(7,*) ERRNUM, ERRMAT(1),ERRMAT(2),ERRMAT(3),ERRMAT(4)
      WRITE(7,*) ERRMAT(5),ERRMAT(6),ERRMAT(7),ERRMAT(8),ERRMAT(9)
      WRITE(7,*) ERRMAT(10),ERRMAT(11),ERRMAT(12),ERRMAT(13),ERRMAT(14)
      WRITE(7,*) ERRMAT(15),ERRMAT(16),ERRMAT(17),ERRMAT(18),ERRMAT(19)
      CLOSE(7)

      RETURN
      END
