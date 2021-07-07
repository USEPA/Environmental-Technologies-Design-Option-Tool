C *****************************************************************************
C *                                                                           *
C *                "GOLDEN" LINEAR EQUATION SOLVER FOR HANFAC                 *
C *                                                                           *
C *****************************************************************************

      SUBROUTINE GOLDEN (TT,NG,A,IMAX,TOL,N,X,FX,HFERR)

      IMPLICIT NONE
      DOUBLE PRECISION XMOLFR,XMOLFW,RGAS,VOLM,VH2O,DADSP,DADSPW,
     &                  A,X,FX,TT,TOL,HFERR,OBJECT,SECT,Y,B,
     &                  F1,F2,FSAVE,
     &                  UNC,FX1,X2,X1,FX2
      INTEGER N,NG,IMAX,KFLAG

      COMMON /INFO/ XMOLFR,XMOLFW,RGAS,VOLM,VH2O,DADSP,DADSPW

C    -- ENTER OBJECTIVE FUNCTION TO BE MINIMIZED

      OBJECT(X)=DABS(-DADSP/VOLM + (RGAS*TT/VOLM)*DLOG(X/XMOLFR) +
     &               DADSPW/VH2O - (RGAS*TT/VH2O)*DLOG((1.D0-X)/XMOLFW))

C    -- STATEMENT FUNCTION TO IMPLEMENT GOLDEN SECTION

      SECT(X,Y) = X + 0.618*Y

C    -- INITIALIZE VARIABLES 

      KFLAG = 0
      N = 0
      B = 0.999999
      F1 = OBJECT(A)
      F2 = OBJECT(B)
      FSAVE = F1

      IF (F2.GT.FSAVE) GOTO 10

      FSAVE = F2

  10  UNC = B-A

      IF (UNC.LE.TOL) GOTO 45

      IF (N.EQ.IMAX) THEN

           HFERR = -1
           RETURN

      END IF

      IF (N.EQ.0) GOTO 15

      IF (KFLAG.EQ.1) GOTO 30

      IF (KFLAG.EQ.2) GOTO 40

  15  X1 = SECT(B,-UNC)

      IF (X1.GE.1) THEN

            X1 = 0.99999

      END IF

      FX1 = OBJECT(X1)

      IF (N.GT.0) GOTO 25

  20  X2 = SECT(A,UNC)

      IF (X2.GE.1) THEN

            X2 = 0.99999
 
      END IF

      FX2 = OBJECT(X2)

  25  N = N+1

      IF (FX1.GT.FX2) GOTO 35

C    -- BRANCH FOR F(X1) < F(X2)

      KFLAG = 1
      B = X2
  
      GOTO 10

  30  X2 = X1
      FX2 = FX1

      GOTO 15

C    -- BRANCH FOR F(X1) > F(X2)

  35  KFLAG = 2
      A = X1

      GOTO 10

  40  X1 = X2
      FX1 = FX2

      GOTO 20

  45  X = (A+B)/2

      IF (X.GT.1) THEN

           HFERR = -1
           RETURN

      END IF

      FX = OBJECT(X)
      HFERR = 1

      END
