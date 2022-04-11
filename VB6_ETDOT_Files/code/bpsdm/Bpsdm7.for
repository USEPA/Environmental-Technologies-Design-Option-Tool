C
C....declaration block
C
      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DOUBLE PRECISION KF,MW
      DIMENSION TDATA(600),TIE(5),TINC(5),Y0(600),CD1(600),
     $          TDATA1(600),TIE1(5),TINC1(4),BR(14,14)
      COMMON/BLOCKA/TBAR,WR(14),BIC,DG,BEDS(14,14),BEDP(14,14),TCONV
      COMMON/BLOCKB/FMIN,TP(600),CP(600),CD(600),CINT(600)
      COMMON/BLOCKC/XNI,CBO,MW
      COMMON NC,NCOMP
      OPEN (3, FILE = 'D:\FOR\COL.TXT')
      OPEN (4, FILE = 'D:\FOR\BPSDM7.DAT')
      OPEN (7, FILE = 'D:\FOR\BPSDM7.OUT')
C
C.....read in data from namelist
C
      READ(4,*) WT,VOL,RAD,RHOP,EPS1,EPOR,NDATA
      READ(4,*) DT0,DSTEP,DTOL,DH0,DOUT1,NM
      READ(4,*) DS,SHIFT,DSTOT
      READ(4,*) KF
      READ(4,*) DP
      READ(4,*) XK
      READ(4,*) XN
      READ(4,*) CBO
      READ(4,*) MW
      CBO=CBO/MW
      READ(4,*) (TIE1(I),I=1,NM)
      READ(4,*) (TINC1(I),I=1,NM)
      DO 5 I = 1,NDATA
       READ (4,*) TDATA1(I),CD1(I)
       CD1(I)=CD1(I)/MW
5     CONTINUE
C
C.....read in collocation constants
C
      READ(3,*) NC
      READ(3,100) (WR(I),I = 1,NC)
      DO 10 I = 1,NC
       READ(3,100) (BR(I,J),J = 1,NC)
10    CONTINUE
      NEQ=NC + 1
      FMIN1=999.0D0
C
C.....print out model parameters
C
20    ECMBR=1.0D0-WT/(RHOP*VOL)
      WRITE(*,120) NC,NEQ,WT,VOL,ECMBR,EPOR,RAD,RHOP,DTOL
      WRITE(7,120) NC,NEQ,WT,VOL,ECMBR,EPOR,RAD,RHOP,DTOL
C
C.....calculate and print out dimensionless groups
C
      IFINAL=0
      QE=XK*CBO**XN
      YE=XK*CBO**XN+EPOR*1.0D-3*CBO/RHOP
      DGS=RHOP*QE*(1.0D0-ECMBR)*1000.0D0/(CBO*ECMBR)
      DGP=EPOR*(1.0D0-ECMBR)/ECMBR
      DG=DGS+DGP
      XNI=1.0D0/XN
22    CONTINUE
      RATE=DS*DG+DP*DGP
      BIC=RAD*KF*(1.0D0-ECMBR)/(RATE*ECMBR)
      X=DS*DG/RATE
      Z=(DP-DS)*DGP/RATE
      TBAR=60.0D0*(DP*CBO*EPOR/(RHOP*YE*1000.0D0)+DS)/RAD**2.0D0
C
C.....determine maximum time conversion constant
C
      TCONV=TBAR
C
C.....convert independent variables to dimensionless form
C
      H0=DH0*TCONV
      T0=DT0*TCONV
      TOUT=DOUT1*TCONV
      TTOL=DTOL*TCONV
      TSTEP=DSTEP*TCONV
      DO 35 I=1,NM
       TINC(I)=TINC1(I)*TCONV
       TIE(I)=TIE1(I)*TCONV
35    CONTINUE
      DO 45 J=1,NDATA
       TDATA(J)=TDATA1(J)*TCONV
       CD(J)=CD1(J)/CBO
45    CONTINUE
C
C.....combine collocation matrix with adsorbate parameters
C.....to decrease computation time
C
       DO 47 J = 1,NC
        DO 46 K = 1,NC
         BEDS(J,K)=BR(J,K)*X
         BEDP(J,K)=BR(J,K)*Z
46      CONTINUE
47     CONTINUE
C
C.....total number of differential equations
C
      N=NC+1
      INDEX=1
      MF=22
      EPS=EPS1
      NSTEPS=1000
C
C.....initialize dependent variables
C
      MM=0
      DO 50 J=1,NC
       MM=MM+1
       Y0(MM)=0.0D0
50    CONTINUE
      MM=MM+1
      Y0(MM)=1.0D0
C
C.....loop for calling GEAR to integrate differential equations
C
      MA=1
      ITP=0
60    ITP=ITP+1
      CALL GEAR(N,T0,H0,Y0,TOUT,EPS,MF,INDEX)
      CP(ITP)=Y0(NC+1)
      TP(ITP)=TOUT
      DOUT=TOUT/TCONV
      IF(ITP .LT. NSTEPS) THEN
       IF(TOUT .LT. TTOL) THEN
        IF(NM .NE. 0 .AND. TOUT .GT. TIE(MA))THEN
         TSTEP=TINC(MA)
         IF(MA .EQ. NM ) THEN
          NM = 0
         ELSE
          MA=MA+1
         ENDIF
        ENDIF
        TOUT=TOUT+TSTEP
        IF(TOUT .GT. TTOL) TOUT=TTOL
        GO TO 60
       ENDIF
      ELSE
       IF ( TOUT .NE. TTOL ) THEN
        WRITE(7,160) NSTEPS,DOUT
       ENDIF
      ENDIF
C
C.....if experimental data is given call OBJFUN to determine
C.....FMIN for each component and print out results
C
      IF(NDATA .EQ. 0) GOTO 85
      CALL OBJFUN(TDATA,NDATA,ITP)
      WRITE(*,200) DS,FMIN
      WRITE(7,200) DS,FMIN
      IF (DS .GT. DSTOT) GOTO 65
      DS=DS+SHIFT
      GOTO 22
65    CONTINUE
      DO 75 J=1,NDATA
       RES=((CINT(J)-CD(J))/CD(J))*100.0D0
       TDATA(J)=TDATA(J)/TCONV
75    CONTINUE
85    STOP
C
C
C                      ---- FORMAT BLOCK ----
C
C
100   FORMAT(4D20.12)
110   FORMAT(1X,4D20.12)
120   FORMAT(////
     $    ' NUMBER OF COLLOCATION POINTS, NC............ =',I15/
     $    ' TOTAL NO. OF DIFFERENTIAL EQUATIONS, NEQ.... =',I15/
     $    ' MASS OF ADSORBENT, WT (GRAMS)............... =',G15.5/
     $    ' VOLUME OF REACTOR, VOL (CM**3).............. =',G15.5/
     $    ' VOID FRACTION OF REACTOR, ECMBR (DIM.)...... =',G15.5/
     $    ' VOID FRACTION OF ADSORBENT, EPOR (DIM.)..... =',G15.5/
     $    ' RADIUS OF ADSORBENT PARTICLE, RAD (CM)...... =',G15.5/
     $    ' APPARENT PARTICLE DENSITY, RHOP (GM/CM**3).. =',G15.5/
     $    ' TOTAL RUN TIME, DTOL (MIN).................. =',G15.5/)
160   FORMAT('MORE STEPS ATTEMPTED THAN NSTEPS',16X,
     $       'NSTEPS =',I3/,'AND TOUT(MIN)=',G10.6)
200   FORMAT('DS = ',G20.10,' FMIN = ',G20.10)
      END
C
                  SUBROUTINE DIFFUN ( N,T,Y0,YDOT )
C
C   *****************************************************************
C   * This subroutine is called by GEAR in the integration process. *
C   * It receives the values of the dependent variables from GEAR   *
C   * and returns the values of the derivatives of the dependent    *
C   * variables.  This continues until the total run time is met.   *
C   *****************************************************************
C
      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DIMENSION YDOT(600),BB(14),Y0(600),CPORE(600)
      DOUBLE PRECISION MW
      COMMON/BLOCKA/TBAR,WR(14),BIC,DG,BEDS(14,14),BEDP(14,14),TCONV
      COMMON/BLOCKC/XNI,CBO,MW
      COMMON NC,NCOMP
C
C.....determine surface concentrations using Ideal Adsorbed
C.....Solution Theory
C
      DO 21 I=1,NC
       IF(Y0(I) .LE. 0.0D0 .OR. CBO .EQ. 0.0D0) THEN
        CPORE(I)=0.0D0
       ELSEIF (XNI*DLOG10(Y0(I)) .LT. -20.0D0) THEN
        CPORE(I)=0.0D0
       ELSE
        CPORE(I)=Y0(I)**XNI
       ENDIF
21    CONTINUE
C
      ND=NC -1
      LL=0
      KK=0
      II=0
C
C.....Overall Liquid Phase Mass Balance
C
C
      DO 30 J=1,ND
       BB(J)=0.0D0
30    CONTINUE
      WW=0.0D0
      DO 50 J=1,ND
       II=II+1
       DO 40 K=1,NC
        KK=KK+1
        BB(J)=BB(J)+BEDS(J,K)*Y0(KK)+BEDP(J,K)*CPORE(KK)
40     CONTINUE
C
C.....Intraparticle Phase Mass Balance (excluding boundary)
C
       YDOT(II)=BB(J)*TBAR/TCONV
       WW=WW+WR(J)*YDOT(II)
       KK=KK-NC
50    CONTINUE
      LL=LL+NC
C
C.....Liquid-Solid Boundary Layer Mass Balance
C
      YDOT(LL)=((BIC*(Y0(LL+1)-CPORE(LL))-WW)/WR(NC))*TBAR/TCONV
C
C.....Overall Mass Balance
C
      YDOT(LL+1)=(-3.0D0*DG*(WW+(YDOT(LL)*WR(NC))))
C
      RETURN
      END
C
                  SUBROUTINE OBJFUN ( TD,NDATA,NP )
C
C   ***************************************************************
C   * This subroutine calculates the standard deviation between   *
C   * the predicted concentrations and experimental data, if any  *
C   * is given.  If no data is given this subroutine is ignored.  *
C   ***************************************************************
C
      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      DIMENSION TD(600)
      DOUBLE PRECISION MW
      COMMON/BLOCKB/FMIN,TP(600),CP(600),CD(600),CINT(600)
      COMMON/BLOCKC/XNI,CBO,MW
      COMMON NC,NCOMP
C
      FMIN=0.0D0
      NP1=NP-1
      DO 10 J=1,NDATA
       DO 5 I=1,NP1
        IF(TD(J) .LT. TP(I) .OR. TD(J) .GT. TP(I+1) ) GOTO 5
        CAP=CP(I)+((TD(J)-TP(I))/(TP(I+1)-TP(I)))*(CP(I+1)-CP(I))
        CINT(J)=CAP
        ZZ=0.0D0
        ZZ=((CD(J)*CBO*MW-CINT(J)*CBO*MW)**2/(CD(J)*CBO*MW)**2)**0.50D0
        FMIN=FMIN+ZZ
        GOTO 10
5      CONTINUE
10    CONTINUE
      RETURN
      END
C
                 SUBROUTINE PEDERV ( N,T,Y0,PD,N0 )
C
C       ******************************************************
C       * This subroutine is a dummy subprogram used by GEAR *
C       ******************************************************
C
      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
      COMMON NC,NCOMP
      RETURN
      END
