!CC*********************************************************************
!CC
!CC                                   AQSOL
!CC                        CALCULATE AQEOUS SOLUBILITY
!CC
!CC Output Variables:
!CC    SOLUB =    Aqueous solubility (PPMw)
!CC    TIE =      Solubility of water in the organic phase (PPMw)
!CC    IIERR =    Error flag for this routine
!CC
!CC Input Variables:
!CC    NC =
!CC    NG =
!CC    TT =       Temperature of calculation (K)
!CC    XMW =
!CC
!CC Authors:  M. Miller, T. Rogers, and D. Hokanson (4/5/94)
!CC
!CC*********************************************************************

      SUBROUTINE AQSOL (NC,NG,TT,SOLUB,TIE,IIERR,XMW)  
!MS$ ATTRIBUTES DLLEXPORT, STDCALL::AQSOL
!MS$ ATTRIBUTES ALIAS:'_AQSOL@28':: AQSOL
!MS$ ATTRIBUTES REFERENCE::NC,NG,TT,SOLUB,TIE,IIERR,XMW

      IMPLICIT DOUBLE PRECISION (A-H,O-Z)

      DIMENSION XGUESS(2),XSOLN(2),FF(2),XMW(10)
      DIMENSION X1(10),X2(10),XE(2),IE(2)

      COMMON /LIMITS/ TOL,IMAX
      COMMON /ERR/ ERRMAT(30),ERRNUM


!CC    -- INITIALIZE VARIABLES

      TIE = 0.0D0
      KSAV = 0 
      XGUESS(1) = 1.0D+00
      XGUESS(2) = 0.0D+00
      FF(1) = 0.0D0
      FF(2) = 0.0D0
      XSOLN(1) = 0.0D0
      XSOLN(2) = 0.0D0
      IIERR = 0
      IERR = 0

      CALL NEWTON (NC,TT,NG,IMAX,TOL,XGUESS,XSOLN,FF,IERR)

      IF (IERR.EQ.-1) THEN

           IIERR = -1
           GOTO 30

      ELSE IF (IRR.EQ.-2) THEN
           IIERR = -2
           GOTO 30

      END IF 

!CC    -- SORT COMPOSITION (DESCENDING ORDER)

      DO 10 I=1,NC

           IF (XSOLN(I).LE.0) THEN

               IIERR = -3
               GOTO 30

           END IF

           IE(I) = 1

           DO 10 J=1,NC

                IF (J.EQ.I) GOTO 10
    
                DIFF = DABS((XSOLN(J)-XSOLN(I))/XSOLN(I))*100.0D0

                IF (DIFF.LE.0.1) THEN

                     IIERR = -4
                     GOTO 30
 
                END IF
 
                IF (XSOLN(I).LT.XSOLN(J)) IE(I)=IE(I)+1

  10  CONTINUE

      DO 20 I=1,NC

           XE(I) = XSOLN(IE(I))

  20  CONTINUE

      X1(1) = XE(1)
      X1(2) = 1.0D0 - X1(1)
      X2(1) = XE(2)
      X2(2) = 1.0D0 - X2(1)

!CC    -- CONVERT MOLE FRACTION TO "PPMW" --

      XMF = 1.0D0 - XE(1)
      XE(1) = 1.0D+06/(1.0D0+((1.0D0/XE(2))-1.0D0)*XMW(2)/XMW(1))
      XE(2) = 1.0D+06/(1.0D0+((1.0D0/XMF)-1.0D0)*XMW(1)/XMW(2))
      SOLUB = XE(2)
      TIE = XE(1)

  30  IF ((SOLUB.LE.TOL).AND.(KSAV.EQ.0).AND.(IIERR.GE.0)) THEN
         IIERR = -4
      END IF

      IF (IIERR.LT.0) THEN
           CALL ERROR (ERRMAT,ERRNUM,9)
           RETURN
      END IF

      END


