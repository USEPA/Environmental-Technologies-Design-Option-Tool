CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC
CC
CC  Program Name:       LIQUID-PHASE ADSORPTION ISOTHERM CALCULATOR
CC                      USING HANSEN/FACKLER MODEL
CC  Author:             Michigan Tech University - 1994
CC  Intended Platform:  Compiled with Microsoft FORTRAN and linked
CC                      to the Visual Basic code of the Adsorption
CC                      Simulation Software.
CC
CC  Modification History:
CC  =====================
CC  ??/??/19??: Fred Gobin (?)
CC  - Modified for DLL
CC  06/28/1996: Eric Oman
CC  - Added code to output LNCPTS and LNQPTS for output graph
CC    (Note maximum of 200 regression points.)
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
C     NG     : ?????
C     NL     : Number of regression points (?)      I
C     OMAG   : Order of magnitude (?)               I
C     ORGDEN : ????                                 I
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
C     HFERR  : Error flag for HANFAC                O
C     ERRMAT : Matrix of error flags                O
C     ----------------------------------------------------------
C     Note 1:
C     All Double Precision Input parameters are stored in the In_Data Array
C     Note 2:
C      The integers passed as parameters must be INTEGER*2 if declared as Integer
C      in Visual Basic app., or as INTEGER*4 if declared as Long in VB app.

      SUBROUTINE HANFAC (IN_DATA,NG,NL,IMAX,
     &                  CSAV,QSAV,XK1,XK2,XNF,CBEG,CEND,RSQD,RMSE,
     &                  ERRMAT,HFERR,LNCPTS,LNQPTS,NUMPTS)

      IMPLICIT NONE

      DOUBLE PRECISION TOL,IN_DATA(12)
      DOUBLE PRECISION SOLUB,TT,FWT,RNDX
      DOUBLE PRECISION W0,BB,GM,CBULK,ORGDEN,OMAG,
     &                 CSAV,QSAV,XK1,XK2,XNF,CBEG,CEND,RSQD,RMSE

      DOUBLE PRECISION XMOLFR,XMOLFW,RGAS,VOLM,VH2O,ADSP,DADSP,DADSPW,
     &                 PVAP,SUMX,SUMY,SUMXX,SUMYY,SUMXY,DIFF,DENS,
     &                 CZERO,
     &                 CS,RHOM,DP,CINC,PS,CSS,X,PH2O,REFMW,REFDEN,
     &                 REFVOL,RIREF,OREF,GAMMA1,GAMMA2,SF,SFW,QCAP, 
     &                 CLNC,CONC,ADP,ADSPRF,VOLAD,QC,ZVOLAD,ADSPW,PI

      DOUBLE PRECISION CON,PIW,CONW,ZDVOL,XUMVOL,SUMVOL,DELVOL,A,FX,
     &                 XADS,XADSW,DADS,QLNQ,B0,B1,QCAL 

      INTEGER IMAX,J,K,N,NJ,I
      INTEGER*2 ERRMAT(30),ERRNUM,HFERR,NG,NL

      DOUBLE PRECISION LNCPTS(200),LNQPTS(200)
      INTEGER*2 NUMPTS

      COMMON /INFO/ XMOLFR,XMOLFW,RGAS,VOLM,VH2O,DADSP,DADSPW 

C      COMMON /LIMITS/ TOL,IMAX
C      COMMON /INITIAL/ W0,BB,GM,CBULK,ORGDEN,OMAG,NL
C      COMMON /ADSORB/ CSAV,QSAV,XK1,XK2,XNF,CBEG,CEND,RSQD,RMSE
C      COMMON /INFO/ XMOLFR,XMOLFW,RGAS,VOLM,VH2O,ADSP,DADSP,DADSPW
C      COMMON /ERR/ ERRMAT(30),ERRNUM

C    -- INITIALIZE VARIABLES
      ERRNUM = 0
      DO 14 I=1,30
        ERRMAT(I)=0
  14  CONTINUE

      HFERR =0 

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
      TOL   = IN_DATA(12)

      CSAV=0
      QSAV=0
      SUMX=0
      SUMY=0
      SUMXX=0
      SUMYY=0
      SUMXY=0

      DIFF = SOLUB-(CBULK/1000)

      IF (DIFF.LE.0) THEN

           HFERR = -1
           CALL ERROR (ERRMAT,ERRNUM,11)
           GOTO 9999 

      END IF

      DENS = ORGDEN

C    -- CS AND CZERO IN {uM/L}

      CZERO = CBULK/FWT
      CS = (SOLUB*1000)/FWT
      VOLM = FWT/DENS
      VH2O = 18.015
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

           HFERR = -1
           CALL ERROR (ERRMAT,ERRNUM,13)
           GOTO 9999

      END IF

      CINC = (CEND-CBEG)/DP

      PS = PVAP/760

      CSS = (PS/0.08206/TT)*1.0E+06
      DIFF = (CSS*FWT)-CBULK

      IF (DIFF.LE.0) THEN

           HFERR = -1
           CALL ERROR (ERRMAT,ERRNUM,18)
           GOTO 9999

      END IF

C    -- EQUATION FOR VAPOR PRESSURE OF WATER OBTAINED FROM 
C    -- THE PROPERTIES OF GASES AND LIQUIDS BY REID ET AL. (1987)

      X = 1-(TT/647.3)
      PH2O = 221.2*DEXP((-7.76451*X+1.45838*X**1.5-2.77580*X**3-
     &                    1.23303*X**6)/(1-X))
      PH2O=PH2O/1.013

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
      GAMMA1 = ((RNDX**2-1)/(RNDX**2+2))/OREF-0.28
      GAMMA2 = ((RNDX**2-1)/(RNDX**2+2))/OREF
      SF = (VOLM/REFVOL)*GAMMA2
      SFW = (VH2O/REFVOL)*0.28

C    -- HANSEN-FACKLER MODIFICATION

      DO 10 K=1,NL

           QCAP = 0
           CLNC = CBEG+DBLE(K)*CINC
           CONC = DEXP(CLNC)
           ADP = (RGAS*TT)*DLOG(CS/CONC)
           ADSPRF = (ADP/VOLM)*REFVOL/GAMMA1

C     -- CHECK FOR DEXP BOUNDARY

           IF ((-BB*(ADSPRF)**GM).LT.-710) THEN

                HFERR = -1
                CALL ERROR (ERRMAT,ERRNUM,14)
                GOTO 9999

           END IF

           VOLAD = W0*DEXP(-BB*(ADSPRF)**GM)
           QC = VOLAD*RHOM
           ZVOLAD = DLOG(VOLAD)
           ADSP = (ADSPRF/REFVOL)*GAMMA2*VOLM
           ADSPW = (ADSPRF/REFVOL)*0.28*VH2O
           PI = PS*DEXP(-ADSP/(RGAS*TT))
           CON = (PI/0.08206/TT)*1.0E+06
           PIW = PH2O*DEXP(-ADSPW/(RGAS*TT))
           CONW = (PIW/0.08206/TT)*1.0E+06
           XMOLFR = CON/(CON+CONW)

           IF (XMOLFR.LE.1.0D-70) THEN

                XMOLFR = 1.0D-70

           END IF

           XMOLFW = CONW/(CON+CONW)
 
           DO 20 J=K,1,-1
    
                ZDVOL = ZVOLAD*DBLE(K)
                XUMVOL = ZVOLAD*DBLE(J)
                SUMVOL = DEXP(XUMVOL)

                IF (SUMVOL.LE.1.0D-70) THEN
   
                     SUMVOL=1.0D-70

                END IF

                IF (J.EQ.K) THEN

                     DELVOL = DEXP(ZDVOL)

                ELSE

                     DELVOL = SUMVOL-DELVOL

                END IF

                DADSP = SF*(-DLOG(SUMVOL/W0)/BB)**(1/GM)
                DADSPW = SFW*(-DLOG(SUMVOL/W0)/BB)**(1/GM)
                A = 1.0D-70
   
                CALL GOLDEN (TT,NG,A,IMAX,TOL,N,X,FX,HFERR)

                IF (HFERR.EQ.-1) THEN
                    
                     CALL ERROR (ERRMAT,ERRNUM,20)
                     GOTO 9999
 
                END IF
 
           XADS = X
           XADSW = 1-X
           DADS = (XADS*DELVOL)/(XADS*VOLM*1.0E-06+XADSW*VH2O*1.0E-06)-
     &     (XMOLFR*DELVOL)/(XMOLFR*VOLM*1.0E-06+XMOLFW*VH2O*1.0E-06)
           QCAP = QCAP+DADS
           DELVOL = SUMVOL

  20       CONTINUE

           QCAP = QCAP+QC

           IF (QCAP.LE.1.0E-03) THEN

                HFERR = -1
                CALL ERROR (ERRMAT,ERRNUM,21)
                GOTO 9999

           END IF

           QLNQ = DLOG(QCAP)
c
c Code added on 6/28/96 to output ln(C) and ln(Q) for analysis.
c
           LNCPTS(K) = CLNC
           LNQPTS(K) = QLNQ
c
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

c
c Code added on 6/28/96 to output ln(C) and ln(Q) for analysis.
c
      IF (NL .LT. 200) THEN
           LNCPTS(NL+1) = 0D0
           LNQPTS(NL+1) = 0D0
           NUMPTS = NL
      ELSE
           NUMPTS = 200
      END IF
c

C    -- CALCULATE FREUNDLICH "K" AND "1/n" BY LINEAR REGRESSION  
C    -- FREUNDLICH "K" (XK1) IS IN (ug/g)(L/ug)**1/n

      B0 = (SUMY/DP)-(DP*SUMX*SUMXY-SUMX**2*SUMY)/DP/(DP*SUMXX-SUMX**2)
      B1 = (DP*SUMXY-SUMX*SUMY)/(DP*SUMXX-SUMX**2)
      XNF = B1
      XK2 = DEXP(B0)
      XK1 = (XK2*FWT)*(1/FWT)**XNF
      RSQD = 1-(SUMYY-B0*SUMY-B1*SUMXY)/((DP*SUMYY-SUMY**2)/DP)

C    -- CALCULATE ROOT MEAN SQUARE ERROR (RMSE)
 
      RMSE=0
      K=0
      J=0

      DO 30 K=1,NL

           QCAP = 0
           CLNC = CBEG+DBLE(K)*CINC
           CONC = DEXP(CLNC)
           ADP = (RGAS*TT)*DLOG(CS/CONC)
           ADSPRF = (ADP/VOLM)*REFVOL/GAMMA1
           VOLAD = W0*DEXP(-BB*(ADSPRF)**GM)
           QC = VOLAD*RHOM
           ZVOLAD = DLOG(VOLAD)
           ADSP = (ADSPRF/REFVOL)*GAMMA2*VOLM
           ADSPW = (ADSPRF/REFVOL)*0.28*VH2O
           PI = PS*DEXP(-ADSP/(RGAS*TT))
           CON = (PI/0.08206/TT)*1.0E+06
           PIW = PH2O*DEXP(-ADSPW/(RGAS*TT))
           CONW = (PIW/0.08206/TT)*1.0E+06
           XMOLFR = CON/(CON+CONW)

           IF (XMOLFR.LE.1.0-70) THEN
 
                XMOLFR = 1.0D-70

           END IF

           XMOLFW = CONW/(CON+CONW)
 
           DO 35 J=K,1,-1

                ZDVOL= ZVOLAD*DBLE(K)
                XUMVOL = ZVOLAD*DBLE(J)
                SUMVOL = DEXP(XUMVOL)

                IF (SUMVOL.LE.1.0D-70) THEN

                     SUMVOL = 1.0D-70

                END IF
    
                IF (J.EQ.K) THEN

                     DELVOL = DEXP(ZDVOL)

                ELSE

                     DELVOL = SUMVOL-DELVOL

                END IF

                DADSP = SF*(-DLOG(SUMVOL/W0)/BB)**(1/GM)
                DADSPW = SFW*(-DLOG(SUMVOL/W0)/BB)**(1/GM)
                A = 1.0D-70

                CALL GOLDEN (TT,NG,A,IMAX,TOL,N,X,FX,HFERR)

                IF (HFERR.EQ.-1) THEN

                     CALL ERROR (ERRMAT,ERRNUM,20)
                     GOTO 9999
 
                END IF
            
           XADS = X
           XADSW = 1 - X
           DADS = (XADS*DELVOL)/(XADS*VOLM*1.0E-06+XADSW*VH2O*1.0E-06)-
     &     (XMOLFR*DELVOL)/(XMOLFR*VOLM*1.0E-06+XMOLFW*VH2O*1.0E-06)
           QCAP = QCAP + DADS
           DELVOL = SUMVOL

  35       CONTINUE

           QCAP = QCAP+QC

           IF (QCAP.LE.1.0D-70) GOTO 30

           QCAL = XK2*(CONC)**XNF
           RMSE = RMSE+((QCAL-QCAP)/QCAP)**2

  30  CONTINUE 

      RMSE = DSQRT(RMSE/DP)*100
      CBEG = DEXP(CBEG)*FWT
      CEND = DEXP(CEND)*FWT
9999  RETURN
      END
