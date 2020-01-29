$NODEBUG
$NOFLOATCALLS
$STORAGE:4
                                  PROGRAM SEARCH_MIN
C
C                               * Golden-Section Algorithm *
C
C                                 Tony Rogers / April 11, 1995
C
      IMPLICIT REAL*8(A-H,O-Z)
      EXTERNAL  FCN, GOLDEN
C
C                                 Open I/O files:
C
C     OPEN( 5 , FILE='INPUTS.DAT', STATUS='UNKNOWN',
C    &          ACCESS='SEQUENTIAL', FORM='FORMATTED')
C     READ( 5 , * ) MAXIT, FTOL, A, DFLT
C
C                                 Initialization:
C
      MAXIT  =  10000
      FTOL   =  1.D-06
      A      =  0.D0
      DFLT   =  0.D0
      ALPHA  =  0.D0
      OBJ    =  0.D0
C
      CALL GOLDEN(A, MAXIT, FTOL, DFLT, ALPHA, OBJ)
C
      WRITE( * , 901)
      WRITE( * , 902) ALPHA, OBJ
C
      STOP
 901  FORMAT(//,' -- CALCULATED RESULTS :',/)
 902  FORMAT( /,'    Root (Zero) of Equation  =', 1PE11.3, //,
     &          '    Objective Function Value =', 1PE11.3, ///)
      END
C
C                                 Calculate Objective Function:
C
      REAL*8 FUNCTION OBJFCN(ALPHA)
      IMPLICIT REAL*8(A-H,O-Z)
C
C        - Put equation (F=0) to be solved here -
C
      OBJFCN = ALPHA**4 - 2.D0*(ALPHA**3) + 3.D0*ALPHA - 1.D0
      RETURN
      END
C
C                                 Subdivide Uncertainty Interval:
C
      REAL*8 FUNCTION SECT(XVAL,UNC)
      IMPLICIT REAL*8(A-H,O-Z)
      SECT = XVAL + (0.61803D0 * UNC)
      RETURN
      END
C
C                                 Search Subroutine:
C
      SUBROUTINE GOLDEN(A, MAXIT, TOL, DFLT, X, FX)
      IMPLICIT REAL*8(A-H,O-Z)
      EXTERNAL OBJFCN, SECT
      KFLAG = 0
      N = 0
      B = A
      F1 = OBJFCN(A)
      IF(F1 .GT. 0.D0) GOTO 998
      DO 5 I=1,MAXIT
         B = B + 1.D-03
         F2 = OBJFCN(B)
         IF(F2 .GE. 0.D0) GOTO 10
         A = B
         F1 = F2
   5  CONTINUE
      GOTO 998
  10  CONTINUE
      UNC = B - A
      IF(UNC .LE. TOL) GOTO 45
      IF(N .EQ. MAXIT) GOTO 999
      IF(N .EQ. 0) GOTO 15
      IF(KFLAG .EQ. 1) GOTO 30
      GOTO 40
  15  CONTINUE
      X1 = SECT(B,-UNC)
      FX1 = OBJFCN(X1)**2
      IF(N .GT. 0) GOTO 25
  20  CONTINUE
      X2 = SECT(A,UNC)
      FX2 = OBJFCN(X2)**2
  25  CONTINUE
      N = N + 1
      IF(FX1 .GT. FX2) GOTO 35
      KFLAG = 1
      B = X2
      GOTO 10
  30  CONTINUE
      X2 = X1
      FX2 = FX1
      GOTO 15
  35  CONTINUE
      KFLAG = 2
      A = X1
      GOTO 10
  40  CONTINUE
      X1 = X2
      FX1 = FX2
      GOTO 20
  45  CONTINUE
      FA = OBJFCN(A)**2
      FB = OBJFCN(B)**2
      IF(FA .LE. FB) THEN
         X = A
         FX = FA
      ELSE
         X = B
         FX = FB
      ENDIF
      RETURN
C
 998  WRITE( * , 801)
 999  WRITE( * , 802) N
      X = DFLT
      FX = OBJFCN(X)**2
      RETURN
C
 801  FORMAT(//,1X,'** ERROR : ROOT NOT BRACKETED !',//)
 802  FORMAT(//,1X,'** ERROR : SUBROUTINE GOLDEN DID NOT FIND THE ROOT A
     &FTER ',I6,'  ITERATIONS;',/,12X,'DEFAULT VALUE IS RETURNED',//)
      END
