C-----------------------------------------------------------------------
C  IMSL Name:  L2TRG/DL2TRG (Single/Double precision version)
C
C  Computer:   pcdsms/DOUBLE
C
C  Revised:    February 27, 1985
C
C  Purpose:    Compute the LU factorization of a real general matrix.
C
C  Usage:      CALL L2TRG (N, A, LDA, FAC, LDFAC, IPVT, SCALE)
C
C  Arguments:  See LFTRG/DLFTRG.
C
C  Remarks:    See LFTRG/DLFTRG.
C
C  Chapter:    MATH/LIBRARY Linear Systems
C
C  Copyright:  1985 by IMSL, Inc.  All Rights Reserved.
C
C  Warranty:   IMSL warrants only that IMSL testing has been applied
C              to this code.  No other warranty, expressed or implied,
C              is applicable.
C
C-----------------------------------------------------------------------
C
      SUBROUTINE DL2TRG (N, A, LDA, FAC, LDFAC, IPVT, SCALE,NFLAG)
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    N, LDA, LDFAC, IPVT(*),NFLAG,NFLAGBIS
      DOUBLE PRECISION A(LDA,*), FAC(LDFAC,*), SCALE(*)
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    I, INDJ, INFO, K, L
      DOUBLE PRECISION BIG, CURMAX, SMALL, T, VALUE
C                                  SPECIFICATIONS FOR INTRINSICS
C     INTRINSIC  DABS
      INTRINSIC  DABS
      DOUBLE PRECISION DABS
C                                  SPECIFICATIONS FOR SUBROUTINES
      EXTERNAL   E1MES, E1POP, E1PSH, E1STI, DGER, DSCAL, DSWAP, DCRGRG
C                                  SPECIFICATIONS FOR FUNCTIONS
      EXTERNAL   DMACH, IDAMAX
      INTEGER    IDAMAX
      DOUBLE PRECISION DMACH,DDD
C
      NFLAGBIS=0
      CALL E1PSH ('DL2TRG ',NFLAG)
      IF (NFLAG.NE.0) GOTO 9999
C
      IF (N .LE. 0) THEN
         CALL E1STI (1, N)
         CALL E1MES (5, 1, 'The order of the matrix must be '//
     &               'positive while N = %(I1) is given.',NFLAG)
         IF (NFLAG.NE.0) GOTO 9999
         NFLAGBIS=186
         GO TO 9000
      END IF
C
      IF (N .GT. LDA) THEN
         CALL E1STI (1, N)
         CALL E1STI (2, LDA)
         CALL E1MES (5, 2, 'The order of the matrix must be '//
     &               'less than or equal to its leading dimension '//
     &               'while N = %(I1) and LDA = %(I2) are given.',NFLAG)
         IF (NFLAG.NE.0) GOTO 9999
         NFLAGBIS=187
         GO TO 9000
      END IF
C
      IF (N .GT. LDFAC) THEN
         CALL E1STI (1, N)
         CALL E1STI (2, LDFAC)
         CALL E1MES (5, 3, 'The order of the matrix must be '//
     &               'less than or equal to its leading dimension '//
     &          'while N = %(I1) and LDFAC = %(I2) are given.',NFLAG)
         IF (NFLAG.NE.0) GOTO 9999
         NFLAGBIS=188
         GO TO 9000
      END IF
C                                  Preserve a copy of the input matrix
      CALL DCRGRG (N, A, LDA, FAC, LDFAC,NFLAG)
      IF (NFLAG.NE.0) GOTO 9999
C                                  Compute the infinity norm of each row
C                                  of A for scaling purpose
      DO 10  I=1, N
         INDJ = IDAMAX(N,FAC(I,1),LDFAC)
         SCALE(I) = DABS(FAC(I,INDJ))
   10 CONTINUE
C                                  Gaussian elimination with scaled
C                                  partial pivoting
      INFO = 0
      DDD=DMACH(1,NFLAG)
      IF (NFLAG.NE.0) GOTO 9999
      SMALL = DDD
      DDD=DMACH(2,NFLAG)
      IF (NFLAG.NE.0) GOTO 9999
      BIG = DDD
      IF (SMALL*BIG .LT. 1.0D0) SMALL = 1.0D0/BIG
      DO 30  K=1, N - 1
C                                  Find L = pivot index
         L = K
         CURMAX = 0.0D0
         DO 20  I=K, N
            IF (SCALE(I) .GE. SMALL) THEN
               VALUE = DABS(FAC(I,K))/SCALE(I)
            ELSE
               VALUE = DABS(FAC(I,K))
            END IF
            IF (VALUE .GT. CURMAX) THEN
               CURMAX = VALUE
               L = I
            END IF
   20    CONTINUE
         IPVT(K) = L
C                                  Zero pivot implies this column
C                                  already triangularized
         IF (FAC(L,K) .NE. 0.0D0) THEN
C                                  Interchange if necessary
            IF (L .NE. K) THEN
               T = FAC(L,K)
               FAC(L,K) = FAC(K,K)
               FAC(K,K) = T
            END IF
C                                  Compute multipliers
            IF (DABS(FAC(K,K)) .GT. SMALL) THEN
               CALL DSCAL (N-K, -1.0D0/FAC(K,K), FAC(K+1,K), 1)
            END IF
C                                  Row elimination with column indexing
            CALL DSWAP (N-K, FAC(K,K+1), LDFAC, FAC(L,K+1), LDFAC)
            CALL DGER (N-K, N-K, 1.0D0, FAC(K+1,K), 1, FAC(K,K+1),
     &                 LDFAC, FAC(K+1,K+1), LDFAC)
         ELSE
            INFO = K
         END IF
   30 CONTINUE
      IPVT(N) = N
      IF (DABS(FAC(N,N)) .LE. SMALL) INFO = N
C
      IF (INFO .NE. 0) THEN
         CALL E1MES (4, 2, 'The input matrix is singular.  '//
     &               'Some of the diagonal elements of the upper '//
     &               'triangular matrix U of the LU factorization '//
     &               'are close to zero.',NFLAG)
         IF (NFLAG.NE.0) GOTO 9999
         NFLAGBIS=189
      END IF
C
 9000 CALL E1POP ('DL2TRG ',NFLAG)
      IF (NFLAG.NE.0) GOTO 9999
      NFLAG=NFLAGBIS
 9999 RETURN
      END
