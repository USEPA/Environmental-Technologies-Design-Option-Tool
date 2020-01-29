C-----------------------------------------------------------------------
C  IMSL Name:  N7QNF/DN7QNF (Single/Double precision version)
C
C  Computer:   pcdsms/DOUBLE
C
C  Revised:    October 1, 1985
C
C  Purpose:
C
C  Usage:      CALL N7QNF (N, R, LR, DIAG, QTB, DELTA, X, WK1, WK2)
C
C  Arguments:
C     N      - The number of equations to be solbed and the number
C              of unknowns.  (Input)
C     R      -
C     LR     -
C     DIAG   -
C     QTB    -
C     DELTA  -
C     X      -
C     WK1    - Real work array of length N.
C     WK2    - Real work array of length N.
C
C  Copyright:  1985 by IMSL, Inc.  All Rights Reserved.
C
C  Warranty:   IMSL warrants only that IMSL testing has been applied
C              to this code.  No other warranty, expressed or implied,
C              is applicable.
C
C-----------------------------------------------------------------------
C
      SUBROUTINE DN7QNF (N, R, LR, DIAG, QTB, DELTA, X, WK1, WK2, nflag)
      integer nflag
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    N, LR
      DOUBLE PRECISION DELTA, R(*), DIAG(*), QTB(*), X(*), WK1(*),
     &           WK2(*)
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    I, J, JJ, JP1, K, L
      DOUBLE PRECISION ALPHA, BIG, BNORM, EPSMCH, GNORM, QNORM,
     &           SGNORM, SMALL, SUM, TEMP
C                                  SPECIFICATIONS FOR INTRINSICS
C     INTRINSIC  DABS,DMAX1,DMIN1,DSQRT
      INTRINSIC  DABS, DMAX1, DMIN1, DSQRT
      DOUBLE PRECISION DABS, DMAX1, DMIN1, DSQRT
C                                  SPECIFICATIONS FOR SUBROUTINES
      EXTERNAL   DAXPY, DSET
C                                  SPECIFICATIONS FOR FUNCTIONS
      EXTERNAL   DMACH, DDOT, DNRM2
      DOUBLE PRECISION DMACH, DDOT, DNRM2
C
      EPSMCH = DMACH(4, nflag)
C                                  First, calculate the GAUSS-NEWTON
C                                  direction
      JJ = (N*(N+1))/2 + 1
      DO 20  K=1, N
         J = N - K + 1
         JP1 = J + 1
         JJ = JJ - K
         L = JJ + 1
         SUM = DDOT(N-JP1+1,R(L),1,X(JP1),1)
         TEMP = R(JJ)
         IF (TEMP .EQ. 0.0D0) THEN
            L = J
            DO 10  I=1, J
               TEMP = DMAX1(TEMP,DABS(R(L)))
               L = L + N - I
   10       CONTINUE
            TEMP = EPSMCH*TEMP
            IF (TEMP .EQ. 0.0D0) TEMP = EPSMCH
         END IF
         X(J) = (QTB(J)-SUM)/TEMP
   20 CONTINUE
C                                  Test whether the GAUSS-NEWTON
C                                  direction is acceptable
      CALL DSET (N, 0.0D0, WK1, 1)
      DO 30  J=1, N
         WK2(J) = DIAG(J)*X(J)
   30 CONTINUE
      QNORM = DNRM2(N,WK2,1, nflag)
      IF (QNORM .GT. DELTA) THEN
C                                  The GAUSS-NEWTON direction is not
C                                  acceptabel. Next, calculate the
C                                  scaled gradient direction
         L = 1
         DO 40  J=1, N
            TEMP = QTB(J)
            CALL DAXPY (N-J+1, TEMP, R(L), 1, WK1(J), 1)
            L = L + N - J + 1
            WK1(J) = WK1(J)/DIAG(J)
   40    CONTINUE
C                                  Calculate the norm of the scaled
C                                  gradient and test for the special
C                                  case in wjich the scaled gradient
C                                  is zero
         GNORM = DNRM2(N,WK1,1, nflag)
         SGNORM = 0.0D0
         ALPHA = DELTA/QNORM
         SMALL = DMACH(1, nflag)
         BIG = DMACH(2, nflag)
         IF (BIG*SMALL .LT. 1.0D0) SMALL = 1.0D0/BIG
         IF (GNORM .NE. 0.0D0) THEN
C                                  Calculate the point along the scaled
C                                  gradient at which the quadratic is
C                                  minimized
            DO 50  J=1, N
               WK1(J) = (WK1(J)/GNORM)/DIAG(J)
   50       CONTINUE
            L = 1
            DO 60  J=1, N
               SUM = DDOT(N-J+1,R(L),1,WK1(J),1)
               L = L + N - J + 1
               WK2(J) = SUM
   60       CONTINUE
            TEMP = DNRM2(N,WK2,1, nflag)
            SGNORM = (GNORM/TEMP)/TEMP
C                                  Test whether the scaled gradient
C                                  direction is acceptable
            ALPHA = 0.0D0
            IF (SGNORM .LT. DELTA) THEN
C                                  The scaled gradient direction is not
C                                  acceptable. Finally, calculate the
C                                  point along the dogleg at which the
C                                  quadratic is minimized
               BNORM = DNRM2(N,QTB,1, nflag)
               TEMP = (BNORM/GNORM)*(BNORM/QNORM)*(SGNORM/DELTA)
               TEMP = TEMP - (DELTA/QNORM)*(SGNORM/DELTA)**2 +
     &                DSQRT((TEMP-(DELTA/QNORM))**2+
     &                (1.0D0-(DELTA/QNORM)**2)*(1.0D0-(SGNORM/DELTA)
     &                **2))
               ALPHA = ((DELTA/QNORM)*(1.0D0-(SGNORM/DELTA)**2))/TEMP
            END IF
         END IF
C                                  Form appropriate convex combination
C                                  of the GAUSS-NEWTON direction and the
C                                  scaled gradient direction
         TEMP = (1.0D0-ALPHA)*DMIN1(SGNORM,DELTA)
         DO 70  J=1, N
            X(J) = TEMP*WK1(J) + ALPHA*X(J)
   70    CONTINUE
      END IF
      RETURN
      END
