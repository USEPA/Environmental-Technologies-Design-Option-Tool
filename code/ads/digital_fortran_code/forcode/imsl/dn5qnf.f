C-----------------------------------------------------------------------
C  IMSL Name:  N5QNF/DN5QNF (Single/Double precision version)
C
C  Computer:   pcdsms/DOUBLE
C
C  Revised:    October 1, 1985
C
C  Purpose:
C
C  Usage:      CALL N5QNF (M, N, A, PIVOT, IPVT, LIPVT, RDIAG, ACNORM,
C                          WK)
C
C  Arguments:
C     M      -
C     N      - The number of equations to be solbed and the number
C              of unknowns.  (Input)
C     A      -
C     PIVOT  -
C     IPVT   -
C     LIPVT  -
C     RDIAG  -
C     ACNORM -
C     WK     - Real work array of length N.  (Output)
C
C  Copyright:  1985 by IMSL, Inc.  All Rights Reserved.
C
C  Warranty:   IMSL warrants only that IMSL testing has been applied
C              to this code.  No other warranty, expressed or implied,
C              is applicable.
C
C-----------------------------------------------------------------------
C
      SUBROUTINE DN5QNF (M, N, A, PIVOT, IPVT, LIPVT, RDIAG, ACNORM,
     &                   WK, nflag)
      integer nflag
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    M, N, LIPVT, IPVT(*)
      DOUBLE PRECISION A(N,*), RDIAG(*), ACNORM(*), WK(*)
      LOGICAL    PIVOT
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    J, JP1, K, KMAX, MINMN
      DOUBLE PRECISION AJNORM, BIG, EPSMCH, SMALL, SUM, TEMP
C                                  SPECIFICATIONS FOR INTRINSICS
C     INTRINSIC  DMAX1,MIN0,DSQRT
      INTRINSIC  DMAX1, MIN0, DSQRT
      INTEGER    MIN0
      DOUBLE PRECISION DMAX1, DSQRT
C                                  SPECIFICATIONS FOR SUBROUTINES
      EXTERNAL   DAXPY, DCOPY, DSCAL, DSWAP
C                                  SPECIFICATIONS FOR FUNCTIONS
      EXTERNAL   DMACH, DDOT, DNRM2
      DOUBLE PRECISION DMACH, DDOT, DNRM2
C

c      write(9,*) 'Enter dn5qnf.'

      EPSMCH = DMACH(4, nflag)
C                                  Compute the initial column norms and
C                                  initialize several arrays
      DO 10  J=1, N
         ACNORM(J) = DNRM2(M,A(1,J),1, nflag)
         IF (PIVOT) IPVT(J) = J
   10 CONTINUE
      CALL DCOPY (N, ACNORM, 1, RDIAG, 1)
      CALL DCOPY (N, RDIAG, 1, WK, 1)
C                                  Reduce A to R with Householder
C                                  transformations
      MINMN = MIN0(M,N)
      DO 40  J=1, MINMN
         IF (PIVOT) THEN
C                                  Bring the column of largest norm into
C                                  the pivot position
            KMAX = J
            DO 20  K=J, N
               IF (RDIAG(K) .GT. RDIAG(KMAX)) KMAX = K
   20       CONTINUE
            IF (KMAX .NE. J) THEN
               CALL DSWAP (M, A(1,J), 1, A(1,KMAX), 1)
               RDIAG(KMAX) = RDIAG(J)
               WK(KMAX) = WK(J)
               K = IPVT(J)
               IPVT(J) = IPVT(KMAX)
               IPVT(KMAX) = K
            END IF
         END IF
C                                  Compute the Householder
C                                  transformation to reduce the J-TH
C                                  column of A to a multiple of the J-TH
C                                  unit vector
         AJNORM = DNRM2(M-J+1,A(J,J),1, nflag)
         SMALL = DMACH(1, nflag)
         BIG = DMACH(2, nflag)
         IF (SMALL*BIG .LT. 1.0D0) SMALL = 1.0D0/BIG
         IF (AJNORM .NE. 0.0D0) THEN
            IF (A(J,J) .LT. 0.0D0) AJNORM = -AJNORM
            CALL DSCAL (M-J+1, 1.0D0/AJNORM, A(J,J), 1)
            A(J,J) = A(J,J) + 1.0D0
C                                  Apply the transformation to the
C                                  remaining columns and update the
C                                  norms
            JP1 = J + 1
            IF (N .GE. JP1) THEN
               DO 30  K=JP1, N
                  SUM = DDOT(M-J+1,A(J,J),1,A(J,K),1)
                  TEMP = SUM/A(J,J)
                  CALL DAXPY (M-J+1, -TEMP, A(J,J), 1, A(J,K), 1)
                  IF (PIVOT .AND. RDIAG(K).NE.0.0D0) THEN
                     TEMP = A(J,K)/RDIAG(K)
                     RDIAG(K) = RDIAG(K)*DSQRT(DMAX1(0.0D0,1.0D0-
     &                          TEMP**2))
                     IF (0.05D0*(RDIAG(K)/WK(K))**2 .LE. EPSMCH) THEN
                        RDIAG(K) = DNRM2(M-J,A(JP1,K),1, nflag)
                        WK(K) = RDIAG(K)
                     END IF
                  END IF
   30          CONTINUE
            END IF
         END IF
         RDIAG(J) = -AJNORM
   40 CONTINUE

c      write(9,*) 'Exit dn5qnf.'

      RETURN
      END
