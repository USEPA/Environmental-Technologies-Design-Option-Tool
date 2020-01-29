C *****************************************************************************
C *                                                                           *
C *                    CALCULATE FREUNDLICH PARAMETERS                        *
C *                                                                           *
C *    ALGORITHM DEVELOPED BY:  RANDY D. CORTRIGHT, GRADUATE STUDENT          *
C *                             DAVID W. HAND, SENIOR RESEARCH ENGINEER       *
C *                             [ June 1985 ]                                 *
C *                                                                           *
C *               MODIFIED BY:  TONY N. ROGERS, CHE GRADUATE STUDENT          *
C *                             [ July 1991 ]                                 *
C *               Last Modified 03/15/94 for DLL (F.Gobin)                    *
C *****************************************************************************
C     BB     : Polanyi Parameter                    Input 
C     CBULK  : bulk conc. (ug/l)                    I     
C     CBEG   : Correlation lower bound (ug/l)       Ouput
C     CEND   : Correlation upper bound (ug/l)       O
C     CSAV   : Average bulk concentration (ug/l)    O
C     FWT    : Molecular weight (g/mol)             I
C     GM     : ????                                 I
C     NL     : Number of regression points (?)      I
C     ORGDEN : Density g/cm3                        I
C     PVAP   : Vapor Pressure (mm Hg)               I
C     QSAV   : Polanyi Adsorption Capacity (ug/g)   O
C     RNDX   : Refractive Index                     I
C     TOL    : Tolerance                            I
C     TT     : Temperature (Kelvin)                 I
C     XK1    : Freundlich K in  (ug/g)...           O
C     XK2    : Freundlich K in (umol/g)...          O
C     XNF    : Freundlich 1/n                       O
C     W0     : Polanyi parameter                    I
C     ----------------------------------------------------------
C     XERR  : Error flag for ??                     O
C     SQERR  : Error FLag for ??                    O
C     ERRMAT : Matrix of error flags                O
C     ----------------------------------------------------------
C     Note 1:
C     All Double Precision Input parameters are stored in the In_Data Array
C     Note 2:
C      The integers passed as parameters must be INTEGER*2 if declared as Integer
C      in Visual Basic app., or as INTEGER*4 if declared as Long in VB app.

      SUBROUTINE SPEQ (IN_DATA,NL,
     &               CSAV,QSAV,XK1,XK2,XNF,CBEG,CEND,
     &               ERRMAT,XERR,SQERR)
      IMPLICIT NONE
C
C------ PARAMETERS FOR SPEQ CALCULATION MODULE.
C
C...INPUTS...:
      DOUBLE PRECISION IN_DATA(10)
      INTEGER*2 NL
      INTEGER*2 XERR
C...OUTPUTS...:
      DOUBLE PRECISION CSAV
      DOUBLE PRECISION QSAV
      DOUBLE PRECISION XK1
      DOUBLE PRECISION XK2
      DOUBLE PRECISION XNF
      DOUBLE PRECISION CBEG
      DOUBLE PRECISION CEND
      INTEGER*2 ERRMAT(30)
      INTEGER*2 SQERR
C
C------ COMMON VARIABLES.
C
      INCLUDE 'COMMON.FI'
C
C------ LOCAL VARIABLES.
C
      DOUBLE PRECISION TT,FWT,PVAP,RNDX,SPRD,BETA
      DOUBLE PRECISION TOL,W0,BB,CBULK,ORGDEN,GM
      DOUBLE PRECISION RGAS,DENS,PS,CS,DIFF,REFMW,REFDEN,RIREF,OREF,
     &                 QMIN,YMIN,PMAX,QMAX,SSAV,XSTEP,SUMY,QVAL,YVAL,
     &                 ARG,VALUE,PI,WEIGHT,YMAX,CONC,CINC,
     &                 QTST,QINT
      INTEGER*2 ERRNUM
      INTEGER ITST,NM,JMAX,ICOUNT,IFCN
      INTEGER I
C      COMMON /LIMITS/ TOL
C      COMMON /INITIAL/ W0,BB,GM,CBULK,ORGDEN,OMAG,NL
C      COMMON /ADSORB/ CSAV,QSAV,XK1,XK2,XNF,CBEG,CEND,RSQD,RMSE
C      COMMON /ERR/ ERRMAT(30),ERRNUM

C
C------ INITIALIZE ERROR PARAMETERS.
C
      ERRNUM = 0
      DO I=1,30
        ERRMAT(I) = 0
      ENDDO

C    -- PARAMETERS IN D-R EQUATION
      BB    = IN_DATA(1)
      W0    = IN_DATA(2)
      TT    = IN_DATA(3)
      CBULK = IN_DATA(4)
      ORGDEN= IN_DATA(5)
      FWT   = IN_DATA(6)
      PVAP  = IN_DATA(7)
      RNDX  = IN_DATA(8)
      GM    = IN_DATA(9)
      TOL   = IN_DATA(10)

      SPRD = 0
      RGAS = 1.987
      NM = IDINT(GM)
      DENS = ORGDEN

      PS = PVAP/760

C    -- CS AND CBULK IN {ug/L}

      CS = (PS/0.08206/TT)*(FWT*1.0E+06)
      DIFF = CS-CBULK

      IF (DIFF.LE.0) THEN
        SQERR = -1
        CALL ERROR (ERRMAT,ERRNUM,18)
        GOTO 9999
      END IF

C    -- DUBININ-RADUSKEVICH (D-R) CORRELATION

C    -- CALCULATE REFRACTIVE INDEX "SCALE FACTOR"

C    -- REFERENCE CHEMICAL : TOLUENE

C    --               MW = 92.140 {g/gmol}
C    --          DENSITY = 0.8623 {g/cc}
C    -- REFRACTIVE INDEX = 1.4941

      REFMW = 92.14
      REFDEN = 0.8623
      RIREF = 1.4941
      OREF = (REFMW/REFDEN)*(RIREF**2-1)/(RIREF**2+2)
      BETA = (FWT/DENS)*(RNDX**2-1)/(RNDX**2+2)/OREF

C    -- CALCULATE UPPER LIMIT OF SURFACE LOADING

      QMIN = 0
      YMIN = 0
      PMAX = (CBULK*1.0E-06/FWT)*(0.08206*TT)

C    -- CHECK FOR DEXP BOUNDARY

      IF ((-BB*((RGAS*TT/BETA)*DLOG(PS/PMAX))**NM).LT.-710) THEN
        SQERR = -1
        CALL ERROR (ERRMAT,ERRNUM,14)
        GOTO 9999
      END IF

      QMAX = W0*(1.0E+06*DENS/FWT)
     &    *DEXP(-BB*((RGAS*TT/BETA)*DLOG(PS/PMAX))**NM)

      IF (ALLOW_SCREENIO.EQ.1) THEN
        PRINT *, '- PERFORMING NUMERICAL INTEGRATION'
      ENDIF

C    -- NUMERICAL INTEGRATION BY SIMPSON'S RULE
C    -- NOTE: STEP SIZE IS REDUCED UNTIL AREA CONVERGES
         
      JMAX = NL+1
      JMAX = (JMAX/2)*2

      IF (JMAX.LE.0) JMAX=2

      SSAV = 0

   5  XSTEP = (QMAX-QMIN)/DBLE(JMAX)
      ICOUNT = 0
      IFCN = 0
      SUMY = 0
      QVAL = QMIN
      YVAL = YMIN

  10  IF (ICOUNT.LE.JMAX) THEN
        IF(ICOUNT.EQ.0) GOTO 15
C
C    -- CALCULATE THE VALUE OF d{ln C}/d{ln q} AT THE GAS CONCENTRATION
C    -- THIS IS EQUAL TO THE FREUNDLICH "N"
C
        QVAL = QVAL+XSTEP
        ARG = QVAL/W0/(1.0E+06*DENS/FWT)
        VALUE = DLOG(ARG*1.0D+25)-DLOG(1.0D+25)
        PI = PS*DEXP(-DSQRT(VALUE/(-BB*(RGAS*TT/BETA)**NM)))
        YVAL = 1/((2*BB)*DLOG(PS/PI)*(RGAS*TT/BETA)**NM)

  15    ICOUNT = ICOUNT+1
        ITST = ((ICOUNT/2)*2)/ICOUNT
        ITST = 2*(ITST+1)
        IFCN = IFCN+ITST
        WEIGHT = DBLE(ITST)
        SUMY = SUMY+(WEIGHT*YVAL)
      
        GOTO 10
      END IF

      YMAX = YVAL
      QMAX = QVAL
      IFCN = IFCN-2
      SUMY = SUMY-YMIN-YMAX
      SUMY = (SUMY/DBLE(IFCN))*(QMAX-QMIN)

      IF (DABS((SSAV-SUMY)/SUMY).GT.TOL) THEN
        JMAX = JMAX*2
        SSAV = SUMY

        IF (ALLOW_SCREENIO.EQ.1) THEN
          PRINT *, '  JMAX = ', JMAX
        ENDIF
        
        GOTO 5
      END IF

      SPRD = SUMY

      IF (JMAX.LE.100000) THEN
        NL = JMAX
      ELSE
        NL = 0
      END IF

C    -- CALCULATE FREUNDLICH PARAMETERS
C    -- FREUNDLICH "K" (XK1) IS IN {(ug/gm) OR (L/ug)**(1/n)}

      IF (ALLOW_SCREENIO.EQ.1) THEN
        PRINT *, '- CALCULATING FREUNDLICH PARAMETERS'
      ENDIF
      XNF = QMAX/SPRD
      XK2 = QMAX/(CBULK/FWT)**XNF
      XK1 = (XK2*FWT)*(1.D0/FWT)**XNF
      CSAV = (PI/0.08206D0/TT)*(FWT*1.D06)
      QSAV = XK1*(CSAV)**XNF

C    -- LOWEST CONC. WHERE ERROR IS LESS THAN (XERR*100) PERCENT
C    -- BETWEEN THE D-R AND FREUNDLICH PREDICTIONS

      CONC = CSAV
      CINC = CONC*TOL
      QTST = QSAV
      QINT = QMAX*FWT

  20  IF (DABS((QTST-QINT)/QINT).LE.XERR) THEN
        CONC = CONC-CINC
        PI = (CONC*1.0E-06/FWT)*(0.08206*TT)
        QINT = W0*(1.0E+06*DENS)
     &     *DEXP(-BB*((RGAS*TT/BETA)*DLOG(PS/PI))**NM)
        QTST = XK1*(CONC)**XNF

        GOTO 20
      END IF

      CBEG = CONC

C    -- HIGHEST CONC. WHERE ERROR IS LESS THAN (XERR*100) PERCENT
C    -- BETWEEN THE D-R AND FREUNDLICH PREDICTIONS

      CONC = CSAV
      CINC = CONC*TOL
      QTST = QSAV
      QINT = QMAX*FWT

  30  IF (DABS((QTST-QINT)/QINT).LE.XERR) THEN
        CONC = CONC+CINC
        PI = (CONC*1.0E-06/FWT)*(0.08206*TT)
        QINT = W0*(1.0E+06*DENS)
     &       *DEXP(-BB*((RGAS*TT/BETA)*DLOG(PS/PI))**NM)
        QTST = XK1*(CONC)**XNF

        GOTO 30
      END IF

      CEND = CONC

C    -- CHECK VALIDITY OF ISOTHERM LIMITS
      
      DIFF = CS-CEND
      CS = CS/1000

      IF (DIFF.LE.0.D0) THEN
        SQERR = -1
        CALL ERROR (ERRMAT,ERRNUM,19)
        GOTO 9999
      END IF

C
C------ COMMENTED OUT BY EJOMAN ON 9/3/98.
C
C      IF (CBEG.GE.CEND) THEN
C        SQERR = -1
C        CALL ERROR (ERRMAT,ERRNUM,13)
C        GOTO 9999
C      END IF
C
C------ END COMMENTED OUT BLOCK.
C

      IF (ALLOW_SCREENIO.EQ.1) THEN
        PRINT *, '- IPE MODULE CALCULATIONS COMPLETE'
      ENDIF
 
 9999 RETURN     
      END
