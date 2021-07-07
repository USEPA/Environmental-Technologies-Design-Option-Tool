!CC*********************************************************************
!CC
!CC                                PARTC
!CC             CALCULATE OCTANOL/WATER PARTITION COEFFICIENT
!CC
!CC Output Variables:
!CC    XKOW =      Octanol water partition coefficient (-)
!CC    XLGK =      log Kow (-)
!CC    JJERR =     Error flag from this routine
!CC
!CC Input Variables:
!CC    TT =        Temperature of calculation (K)
!CC    OCTDEN =
!CC    WATDEN =
!CC    NG =
!CC
!CC Authors:  M. Miller, T. Rogers, D. Hokanson (4/5/94)
!CC
!CC*********************************************************************

      SUBROUTINE PARTC (TT,OCTDEN,WATDEN,XKOW,XLGK,NG,JJERR)
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::PARTC
!MS$ ATTRIBUTES ALIAS:'_PARTC@28':: PARTC
!MS$ ATTRIBUTES REFERENCE::TT,OCTDEN,WATDEN,XKOW,XLGK,NG,JJERR

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)

      DIMENSION XGUESS(2),XSOLN(2),FF(2)
      DIMENSION X1(10),X2(10),XE(2),IE(2),MI(10,2)
      DIMENSION ACT1(10),DACT1(10,10),TACT1(10)
      DIMENSION ACT2(10),DACT2(10,10),TACT2(10)
      DOUBLE PRECISION OCTDEN, WATDEN, XKOW, TT

      COMMON /GROUP/ MS(10,10,2),NMAX
      COMMON /LIMITS/ TOL,IMAX
      COMMON /ERR/ ERRMAT(30),ERRNUM

!CC    -- INITIALIZE VARIABLES
   
      XKOW = 0.0D0
      XLGK = 0.0D0
      NDIF = 0
      NACT = 0
      NC = 2
      NI = NMAX

      DO 10 J=1,10

           DO 10 K=1,2
           
                MI(J,K) = MS(2,J,K)

                     DO 10 I=1,3

                MS(I,J,K) = 0

  10  CONTINUE

!CC    -- DEFINE ARRAYS FOR UNIFAC GROUPS {[1]-WATER AND [2]-OCTANOL}

      NMAX = 3

      MS(1,1,1) = 17
      MS(1,1,2) = 1

      MS(2,1,1) = 1
      MS(2,1,2) = 1
      MS(2,2,1) = 2
      MS(2,2,2) = 7
      MS(2,3,1) = 15
      MS(2,3,2) = 1

!CC    -- FIND OCTANOL/WATER EQUILIBRIUM
    
      CALL FGRP (NC,NG,JERR)

      IF (JERR.EQ.-1) THEN

           JJERR = -1
           CALL ERROR (ERRMAT,ERRNUM,10)
           RETURN

      END IF 

      XGUESS(1)=1.0D0
      XGUESS(2)=0.0D0

      CALL NEWTON (NC,TT,NG,IMAX,TOL,XGUESS,XSOLN,FF,IERR)

      IF (IERR.EQ.-1) THEN
 
           JJERR = -1
           CALL ERROR (ERRMAT,ERRNUM,10)
           RETURN

      END IF

!CC    -- SORT COMPOSITION (DESCENDING ORDER)

      DO 20 I=1,NC

           IF (XSOLN(I).LE.0) THEN

                JJERR = -1
                CALL ERROR (ERRMAT,ERRNUM,10)
                RETURN

           END IF

           IE(I) = 1

           DO 20 J=1,NC
   
                IF (J.EQ.I) GOTO 20

                DIFF = DABS((XSOLN(J)-XSOLN(I))/XSOLN(I))*100.0D0

                IF (DIFF.LE.0.1) THEN 

                     JJERR = -1
                     CALL ERROR (ERRMAT,ERRNUM,10)
                     RETURN

                END IF

                IF (XSOLN(I).LT.XSOLN(J)) IE(I)=IE(I)+1

  20  CONTINUE

      DO 30 I=1,NC

           XE(I) = XSOLN(IE(I))

  30  CONTINUE

!CC    -- MOLE FRACTIONS (INFINITE DILUTION OF CHEMICAL [3])

      X1(1) = XE(1)
      X1(2) = 1.0D0 - X1(1)
      X1(3) = 0.0D0
      X2(1) = XE(2)
      X2(2) = 1.0D0 - X2(1)
      X2(3) = 0.0D0

!CC    -- PARTITIONING FOR DISTRIBUTED CHEMICAL

      NC = 3

      IF (NI.GT.NMAX) NMAX = NI

      DO 35 J=1,10

           DO 35 K=1,2

                MS(NC,J,K) = MI(J,K)

  35  CONTINUE
  
      CALL FGRP (NC,NG,JERR)

      IF (JERR.EQ.-1) THEN

           JJERR = -1
           CALL ERROR (ERRMAT,ERRNUM,10)
           RETURN

      END IF

      CALL PARMS (NC,NG,TT)
      CALL UNIMOD (NDIF,NACT,NC,NG,TT,X1,ACT1,DACT1,TACT1)
      CALL UNIMOD (NDIF,NACT,NC,NG,TT,X2,ACT2,DACT2,TACT2)

!CC    -- XKOW = PARTITION COEFFICIENT
!CC    -- XLGK = BASE-10 LOGARITHM OF XKOW
 
      PHASEW = 1.0D0/(X1(1)/WATDEN + X1(2)/OCTDEN)

      PHASEO = 1.0D0/(X2(1)/WATDEN + X2(2)/OCTDEN)
       
      XKOW = (PHASEO/PHASEW)*(ACT1(3)/ACT2(3))
      XLGK = DLOG10(XKOW)

!CC    -- RESET ORIGINAL "MS" VALUES

      NC = 2
      NMAX = NI
      MS(1,1,1) = 17
      MS(1,1,2) = 1

      DO 60 J=1,10

           DO 60 K=1,2
             
                MS(NC,J,K) = MI(J,K)

  60  CONTINUE
  
      END



