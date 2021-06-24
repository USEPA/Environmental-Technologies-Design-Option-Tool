C-----------------------------------------------------------------------
C  IMSL Name:  N3QNF/DN3QNF (Single/Double precision version)
C
C  Computer:   pcdsms/DOUBLE
C
C  Revised:    September 30, 1985
C
C  Purpose:
C
C  Usage:      CALL N3QNF (FCN, ERRREL, N, X, FVEC, FJAC, R, QTF,
C                          MAXFEV, ML, MU, EPSFCN, MODE, FACTOR, NPRINT,
C                          INFO, NFEV, LR, DIAG, WK1, WK2, WK3, WK4)
C
C  Arguments:
C     FCN    - A real function subroutine supplied by the
C              user.  FCN must be declared EXTERNAL in the
C              CALLING PROGRAM.  FCN specifies the system of
C              equations to be solved and should be of the
C              following form
C                            SUBROUTINE FCN(X,F,N)
C                            REAL X(*),F(*)
C                            F(1)=
C                             .
C                            F(N)=
C                            RETURN
C                            END
C              Where X is given.  FCN must not alter X.
C     ERRREL - Stopping criterion.  The root is accepted if the
C              relative error between two successive approximations
C              to this root is within ERRREL.  (Input)
C     N      - The number of equations to be solved and the number
C              of unknowns.  (Input)
C     X      - A vector of length N.  X contains the best estimate
C              of the root found by NEQNF.  (Output)
C     FVEC   - A vector of length N.  FVEC contains the functions
C              evaluated at the point X.
C     FJAC   - An N by N matrix.  FJAC contains the orthogonal
C              matrix Q produced by the QR factorization of the
C              final approximate Jacobian.
C     R      - A vector of length N*(N+1)/2.  R contains the upper
C              triangular matrix produced by the QR factorization
C              of the final approximation Jacobian.  R is stored
C              rowwise.
C     QTF    - A vector of length N.  QTF contains the vector
C              (Q transpose)*FVEC.
C     MAXFEV - Maximum number of calls to FCN. (Input)
C     ML     -
C     MU     -
C     EPSFCN -
C     MODE   -
C     FACTOR -
C     NPRINT - Number of iterates to be printed. (
C     INFO   -
C     NFEV   - Number of calls to FCN.  (Output)
C     LR     - Length of the vector R.  (Input)
C     DIAG   -
C     WK1    - Real work vector of length N.  (Output)
C     WK2    - Real work vector of length N.  (Output)
C     WK3    - Real work vector of length N.  (Output)
C     WK4    - Real work vector of length N.  (Output)
C
C  Copyright:  1985 by IMSL, Inc.  All rights reserved
C
C  Warranty:   IMSL warrants only that IMSL testing has been applied
C              to this code.  No other warranty, expressed or implied,
C              is applicable.
C
C-----------------------------------------------------------------------
C
      SUBROUTINE DN3QNF (FCN, ERRREL, N, X, FVEC, FJAC, R, QTF,
     &                   MAXFEV, ML, MU, EPSFCN, MODE, FACTOR, NPRINT,
     &                   INFO, NFEV, LR, DIAG, WK1, WK2, WK3, WK4,
     &                   nflag)
C                                  SPECIFICATIONS FOR ARGUMENTS
      integer nflag
      INTEGER    N, MAXFEV, ML, MU, MODE, NPRINT, INFO, NFEV, LR
      DOUBLE PRECISION ERRREL, EPSFCN, FACTOR, X(*), FVEC(*),
     &           FJAC(N,*), R(*), QTF(*), DIAG(*), WK1(*), WK2(*),
     &           WK3(*), WK4(*)
      EXTERNAL   FCN
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    I, IFLAG, ITER, IWA(1), J, JM1, L, MSUM, NCFAIL,
     &           NCSUC, NSLOW1, NSLOW2
      DOUBLE PRECISION ACTRED, DELTA, EPSMCH, FNORM, FNORM1, PNORM,
     &           PRERED, RATIO, SUM, TEMP, XNORM
      LOGICAL    JEVAL, SING
C                                  SPECIFICATIONS FOR INTRINSICS
C     INTRINSIC  DABS,DMAX1,DMIN1,MIN0
      INTRINSIC  DABS, DMAX1, DMIN1, MIN0
      INTEGER    MIN0
      DOUBLE PRECISION DABS, DMAX1, DMIN1
C                                  SPECIFICATIONS FOR SUBROUTINES
      EXTERNAL   E1USR, DAXPY, DCOPY, DSCAL, DN4QNF, DN5QNF, DN6QNF,
     &           DN7QNF, DN8QNF, DN9QNF
C                                  SPECIFICATIONS FOR FUNCTIONS
      EXTERNAL   DMACH, DDOT, DNRM2
      DOUBLE PRECISION DMACH, DDOT, DNRM2
C
c      write(9,*) 'Enter dn3qnf.'

      EPSMCH = DMACH(4, nflag)
      INFO = 0
      IFLAG = 0
      NFEV = 0
C                                  Check the input parameters for
C                                  errors
      IF (MODE .EQ. 2) THEN
         DO 10  J=1, N
            IF (DIAG(J) .LE. 0.0D0) GO TO 150
   10    CONTINUE
      END IF
C                                  Evaluate the function at the starting
C                                  point and calculate its norm
      IFLAG = 1
      CALL E1USR ('ON', nflag)
      CALL FCN (X, FVEC, N)
      CALL E1USR ('OFF', nflag)
      NFEV = 1
      IF (IFLAG .LT. 0) GO TO 150
      FNORM = DNRM2(N,FVEC,1, nflag)
C                                  Determine the number of calls to FCN
C                                  Needed to compute the jacobian
C                                  matrix
C
      MSUM = MIN0(ML+MU+1,N)
C
C                                  Initialize iteration counter and
C                                  monitors
      ITER = 1
      NCSUC = 0
      NCFAIL = 0
      NSLOW1 = 0
      NSLOW2 = 0
C                                  Beginning of the outer loop
   20 JEVAL = .TRUE.
C                                  Calculate the jacobian matrix
      IFLAG = 2
      CALL DN4QNF (FCN, N, X, FVEC, FJAC, IFLAG, ML, MU, EPSFCN, WK1,
     &             WK2, nflag)
      NFEV = NFEV + MSUM
      IF (IFLAG .LT. 0) GO TO 150
C                                  Compute the QR factorization of the
C                                  jacobian
      CALL DN5QNF (N, N, FJAC, .FALSE., IWA, 1, WK1, WK2, WK3, nflag)
C                                  On the first iteration and if MODE is
C                                  1, scale according to the norms of
C                                  the columns of the intial Jacobian
      IF (ITER .EQ. 1) THEN
         IF (MODE .NE. 2) THEN
            CALL DCOPY (N, WK2, 1, DIAG, 1)
            DO 30  J=1, N
               IF (WK2(J) .EQ. 0.0D0) DIAG(J) = 1.0D0
   30       CONTINUE
         END IF
C                                  On the first iteration, calculate the
C                                  norm of the scaled X and initialize
C                                  the step bound delta
         DO 40  J=1, N
            WK3(J) = DIAG(J)*X(J)
   40    CONTINUE
         XNORM = DNRM2(N,WK3,1, nflag)
         DELTA = FACTOR*XNORM
         IF (DELTA .EQ. 0.0D0) DELTA = FACTOR
      END IF
C                                  Form (Q transpose)*FVEC and store in
C                                  QTF.
      CALL DCOPY (N, FVEC, 1, QTF, 1)
      DO 50  J=1, N
         IF (FJAC(J,J) .NE. 0.0D0) THEN
            SUM = DDOT(N-J+1,FJAC(J,J),1,QTF(J),1)
            TEMP = -SUM/FJAC(J,J)
            CALL DAXPY (N-J+1, TEMP, FJAC(J,J), 1, QTF(J), 1)
         END IF
   50 CONTINUE
C                                  Copy the triangular factor of the QR
C                                  factorization into R
      SING = .FALSE.
      DO 70  J=1, N
         L = J
         JM1 = J - 1
         IF (JM1 .GE. 1) THEN
            DO 60  I=1, JM1
               R(L) = FJAC(I,J)
               L = L + N - I
   60       CONTINUE
         END IF
         R(L) = WK1(J)
         IF (WK1(J) .EQ. 0.0D0) SING = .TRUE.
   70 CONTINUE
C                                  Accumulate the orthogonal factor in
C                                  FJAC
      CALL DN6QNF (N, N, FJAC, WK1, nflag)
C                                  Rescale if necessary
      IF (MODE .NE. 2) THEN
         DO 80  J=1, N
            DIAG(J) = DMAX1(DIAG(J),WK2(J))
   80    CONTINUE
      END IF
C                                  Beginning of the inner loop
C                                  If requested, call FCN to enable
C                                  printing of iterates
   90 IF (NPRINT .LE. 0) GO TO 100
      IF (IFLAG .LT. 0) GO TO 150
  100 CONTINUE
C                                  Determine the direction P
      CALL DN7QNF (N, R, LR, DIAG, QTF, DELTA, WK1, WK2, WK3, nflag)
C                                  Store the direction P and X + P
C                                  Calculate the norm of P
      CALL DSCAL (N, -1.0D0, WK1, 1)
      DO 110  J=1, N
         WK2(J) = X(J) + WK1(J)
         WK3(J) = DIAG(J)*WK1(J)
  110 CONTINUE
      PNORM = DNRM2(N,WK3,1, nflag)
C                                  On the first iteration, adjust the
C                                  initial step bound
      IF (ITER .EQ. 1) DELTA = DMIN1(DELTA,PNORM)
C                                  Evaluate the function at X + P and
C                                  calculate its norm
      CALL E1USR ('ON', nflag)
      CALL FCN (WK2, WK4, N)
      CALL E1USR ('OFF', nflag)
      NFEV = NFEV + 1
      FNORM1 = DNRM2(N,WK4,1, nflag)
C                                  Compute the scaled actual reduction
      ACTRED = -1.0D0
      IF (FNORM1 .LT. FNORM) ACTRED = 1.0D0 - (FNORM1/FNORM)**2
C                                  Compute the scaled predicted
C                                  reduction
      L = 1
      DO 120  I=1, N
         SUM = DDOT(N-I+1,R(L),1,WK1(I),1)
         L = L + N - I + 1
         WK3(I) = QTF(I) + SUM
  120 CONTINUE
      TEMP = DNRM2(N,WK3,1, nflag)
      PRERED = 1.0D0
      IF (TEMP .LT. FNORM) PRERED = 1.0D0 - (TEMP/FNORM)**2
C                                  Compute the ratio of the actual to
C                                  the prdeicted reduction
      RATIO = 0.0D0
      IF (PRERED .GT. 0.0D0) RATIO = ACTRED/PRERED
C                                  Update the step bound
      IF (RATIO .GE. 0.1D0) THEN
         NCFAIL = 0
         NCSUC = NCSUC + 1
         IF (RATIO.GE.0.5D0 .OR. NCSUC.GT.1) DELTA =
     &       DMAX1(DELTA,PNORM/0.5D0)
         IF (DABS(RATIO-1.0D0) .LE. 0.1D0) DELTA = PNORM/0.5D0
      ELSE
         NCSUC = 0
         NCFAIL = NCFAIL + 1
         DELTA = 0.5D0*DELTA
      END IF
C                                  Test for successful iteration
      IF (RATIO .GE. 0.0001D0) THEN
C                                  Successful iteration. Update X, FVEC,
C                                  and their norms
         CALL DCOPY (N, WK2, 1, X, 1)
         CALL DCOPY (N, WK4, 1, FVEC, 1)
         DO 130  J=1, N
            WK2(J) = DIAG(J)*X(J)
  130    CONTINUE
         XNORM = DNRM2(N,WK2,1, nflag)
         FNORM = FNORM1
         ITER = ITER + 1
      END IF
C                                  Determine the progress of the
C                                  iteration
      NSLOW1 = NSLOW1 + 1
      IF (ACTRED .GE. 0.001D0) NSLOW1 = 0
      IF (JEVAL) NSLOW2 = NSLOW2 + 1
      IF (ACTRED .GE. 0.1D0) NSLOW2 = 0
C                                  Test for convergence
      IF (DELTA.LE.ERRREL*XNORM .OR. FNORM.EQ.0.0D0) INFO = 1
      IF (INFO .NE. 0) GO TO 150
C                                  Tests for termination and stringent
C                                  tolerances
      IF (NFEV .GE. MAXFEV) INFO = 2
      IF (0.1D0*DMAX1(0.1D0*DELTA,PNORM) .LE. EPSMCH*XNORM) INFO = 3
      IF (NSLOW2 .EQ. 5) INFO = 4
      IF (NSLOW1 .EQ. 10) INFO = 5
      IF (INFO .NE. 0) GO TO 150
C                                  Criterion for recalculating Jacobian
C                                  approximation by forward differences
      IF (NCFAIL .NE. 2) THEN
C                                  Calculate the rank one modification
C                                  to the Jacobian and update QTF if
C                                  necessary
         DO 140  J=1, N
            SUM = DDOT(N,FJAC(1,J),1,WK4,1)
            WK2(J) = (SUM-WK3(J))/PNORM
            WK1(J) = DIAG(J)*((DIAG(J)*WK1(J))/PNORM)
            IF (RATIO .GE. 0.0001D0) QTF(J) = SUM
  140    CONTINUE
C                                  Compute the QR factoracation of the
C                                  updated Jacobian
         CALL DN8QNF (N, N, R, WK1, WK2, WK3, SING, nflag)
         CALL DN9QNF (N, N, FJAC, N, WK2, WK3, nflag)
         CALL DN9QNF (1, N, QTF, 1, WK2, WK3, nflag)
C                                  End of the inner loop
         JEVAL = .FALSE.
         GO TO 90
      END IF
C                                  End of the outer loop
      GO TO 20
C                                  Termination, either normal or user
C                                  imposed
  150 IF (IFLAG .LT. 0) INFO = IFLAG
      IFLAG = 0

c      write(9,*) 'Exit dn3qnf.'

      RETURN
      END
