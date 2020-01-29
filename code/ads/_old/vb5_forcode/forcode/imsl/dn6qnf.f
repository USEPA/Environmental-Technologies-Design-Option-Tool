C-----------------------------------------------------------------------
C  IMSL Name:  N6QNF/DN6QNF (Single/Double precision version)
C
C  Computer:   pcdsms/DOUBLE
C
C  Revised:    October 1, 1985
C
C  Purpose:    Calculate the orthogonal factor in to solve a system of
C              nonlinear equations
C
C  Usage:      CALL N6QNF (M, N, Q, WK)
C
C  Arguments:
C     M      -
C     N      - The number of equations to be solbed and the number
C              of unknowns.  (Input)
C     Q      -
C     WK     - Real work array of length N. (
C
C  Copyright:  1985 by IMSL, Inc.  All Rights Reserved.
C
C  Warranty:   IMSL warrants only that IMSL testing has been applied
C              to this code.  No other warranty, expressed or implied,
C              is applicable.
C
C-----------------------------------------------------------------------
C
      SUBROUTINE DN6QNF (M, N, Q, WK, nflag)
      integer nflag
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    M, N
      DOUBLE PRECISION Q(N,*), WK(*)
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    J, JM1, K, L, MINMN, NP1
      DOUBLE PRECISION BIG, SMALL, SUM, TEMP
C                                  SPECIFICATIONS FOR INTRINSICS
C     INTRINSIC  MIN0
      INTRINSIC  MIN0
      INTEGER    MIN0
C                                  SPECIFICATIONS FOR SUBROUTINES
      EXTERNAL   DAXPY, DCOPY, DSET
C                                  SPECIFICATIONS FOR FUNCTIONS
      EXTERNAL   DMACH, DDOT
      DOUBLE PRECISION DMACH, DDOT
C
      MINMN = MIN0(M,N)
      IF (MINMN .GE. 2) THEN
         DO 10  J=2, MINMN
            JM1 = J - 1
            CALL DSET (JM1, 0.0D0, Q(1,J), 1)
   10    CONTINUE
      END IF
C                                  Initialize remaining columns to those
C                                  of the identity matrix
      NP1 = N + 1
      IF (M .GE. NP1) THEN
         DO 20  J=NP1, M
            CALL DSET (M, 0.0D0, Q(1,J), 1)
            Q(J,J) = 1.0D0
   20    CONTINUE
      END IF
C                                  Accumulate Q from its factored form
      DO 40  L=1, MINMN
         K = MINMN - L + 1
         CALL DCOPY (M-K+1, Q(K,K), 1, WK(K), 1)
         CALL DSET (M-K+1, 0.0D0, Q(K,K), 1)
         Q(K,K) = 1.0D0
         SMALL = DMACH(1, nflag)
         BIG = DMACH(2, nflag)
         IF (BIG*SMALL .LT. 1.0D0) SMALL = 1.0D0/BIG
         IF (WK(K) .NE. 0.0D0) THEN
            DO 30  J=K, M
               SUM = DDOT(M-K+1,Q(K,J),1,WK(K),1)
               TEMP = SUM/WK(K)
               CALL DAXPY (M-K+1, -TEMP, WK(K), 1, Q(K,J), 1)
   30       CONTINUE
         END IF
   40 CONTINUE
      RETURN
      END
