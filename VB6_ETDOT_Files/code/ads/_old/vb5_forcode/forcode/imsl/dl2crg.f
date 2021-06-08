C-----------------------------------------------------------------------
C  IMSL Name:  L2CRG/DL2CRG (Single/Double precision version)
C
C  Computer:   pcdsms/DOUBLE
C
C  Revised:    February 27, 1985
C
C  Purpose:    Compute the LU factorization of a real general matrix and
C              estimate its L1 condition number.
C
C  Usage:      CALL L2CRG (N, A, LDA, FAC, LDFAC, IPVT, RCOND, Z)
C
C  Arguments:  See LFCRG/DLFCRG.
C
C  Remarks:    See LFCRG/DLFCRG.
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
      SUBROUTINE DL2CRG (N, A, LDA, FAC, LDFAC, IPVT, RCOND, Z,NFLAG)
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    N, LDA, LDFAC, IPVT(*),NFLAG,NFLAGBIS
      DOUBLE PRECISION RCOND, A(LDA,*), FAC(LDFAC,*), Z(*)
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    J, K, KP1, L
      DOUBLE PRECISION ANORM, EK, S, SM, T, WK, WKM, YNORM
C                                  SPECIFICATIONS FOR INTRINSICS
C     INTRINSIC  DABS,DSIGN
      INTRINSIC  DABS, DSIGN
      DOUBLE PRECISION DABS, DSIGN
C                                  SPECIFICATIONS FOR SUBROUTINES
      EXTERNAL   E1MES, E1POP, E1PSH, E1STI, E1STD, DAXPY, DSCAL,
     &           DSET, DL2TRG, DNR1RR
C                                  SPECIFICATIONS FOR FUNCTIONS
      EXTERNAL   DMACH, N1RTY, DASUM, DDOT
      INTEGER    N1RTY,II
      DOUBLE PRECISION DMACH, DASUM, DDOT,DDD
C
      NFLAGBIS=0
      CALL E1PSH ('DL2CRG ',NFLAG)
      IF (NFLAG.NE.0) GOTO 9999
C
      IF (N .LE. 0) THEN
         CALL E1STI (1, N)
         CALL E1MES (5, 1, 'The order of the matrix must be '//
     &               'positive while N = %(I1) is given.',NFLAG)
         IF (NFLAG.NE.0) GOTO 9999
         NFLAGBIS=217
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
         NFLAGBIS=218
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
         NFLAGBIS=219
         GO TO 9000
      END IF
C
      RCOND = 0.0D0
C                                  COMPUTE 1-NORM OF A
      CALL DNR1RR (N, N, A, LDA, ANORM,NFLAG)
      IF (NFLAG.NE.0) GOTO 9999
C                                  FACTORIZATION STEP
C
      CALL DL2TRG (N, A, LDA, FAC, LDFAC, IPVT, Z,NFLAG)
      IF (NFLAG.NE.0) GOTO 9999
      II=N1RTY(1,NFLAG)
      IF (NFLAG.NE.0) GOTO 9999
      IF (II.EQ. 4) GO TO 9000
C                                  RCOND = 1/(NORM(A)*(ESTIMATE OF
C                                  NORM(INVERSE(A)))). ESTIMATE =
C                                  NORM(Z)/NORM(Y) WHERE A*Z = Y AND
C                                  TRANS(A)*Y = E . TRANS(A) IS THE
C                                  TRANSPOSE OF A. THE COMPONENTS OF
C                                  E ARE CHOSEN TO CAUSE MAXIMUM LO-
C                                  CAL GROWTH IN THE ELEMENTS OF W
C                                  WHERE TRANS(U)*W = E. THE VECTORS
C                                  ARE FREQUENTLY RESCALED TO AVOID
C                                  OVERFLOW. SOLVE TRANS(U)*W = E
      EK = 1.0D0
      CALL DSET (N, 0.0D0, Z, 1)
      DO 20  K=1, N
         IF (Z(K) .NE. 0.0D0) EK = DSIGN(EK,-Z(K))
         IF (DABS(EK-Z(K)) .GT. DABS(FAC(K,K))) THEN
            S = DABS(FAC(K,K))/DABS(EK-Z(K))
            CALL DSCAL (N, S, Z, 1)
            EK = S*EK
         END IF
         WK = EK - Z(K)
         WKM = -EK - Z(K)
         S = DABS(WK)
         SM = DABS(WKM)
         IF (FAC(K,K) .NE. 0.0D0) THEN
            WK = WK/FAC(K,K)
            WKM = WKM/FAC(K,K)
         ELSE
            WK = 1.0D0
            WKM = 1.0D0
         END IF
         KP1 = K + 1
         IF (KP1 .LE. N) THEN
            DO 10  J=KP1, N
               SM = SM + DABS(Z(J)+WKM*FAC(K,J))
               Z(J) = Z(J) + WK*FAC(K,J)
               S = S + DABS(Z(J))
   10       CONTINUE
            IF (S .LT. SM) THEN
               T = WKM - WK
               WK = WKM
               CALL DAXPY (N-K, T, FAC(K,KP1), LDFAC, Z(KP1), 1)
            END IF
         END IF
         Z(K) = WK
   20 CONTINUE
      S = 1.0D0/DASUM(N,Z,1)
      CALL DSCAL (N, S, Z, 1)
C                                  SOLVE TRANS(L)*Y = W
      DO 30  K=N, 1, -1
         IF (K .LT. N) Z(K) = Z(K) + DDOT(N-K,FAC(K+1,K),1,Z(K+1),1)
         IF (DABS(Z(K)) .GT. 1.0D0) THEN
            S = 1.0D0/DABS(Z(K))
            CALL DSCAL (N, S, Z, 1)
         END IF
         L = IPVT(K)
         T = Z(L)
         Z(L) = Z(K)
         Z(K) = T
   30 CONTINUE
      S = 1.0D0/DASUM(N,Z,1)
      CALL DSCAL (N, S, Z, 1)
C
      YNORM = 1.0D0
C                                  SOLVE L*V = Y
      DO 40  K=1, N
         L = IPVT(K)
         T = Z(L)
         Z(L) = Z(K)
         Z(K) = T
         IF (K .LT. N) CALL DAXPY (N-K, T, FAC(K+1,K), 1, Z(K+1), 1)
         IF (DABS(Z(K)) .GT. 1.0D0) THEN
            S = 1.0D0/DABS(Z(K))
            CALL DSCAL (N, S, Z, 1)
            YNORM = S*YNORM
         END IF
   40 CONTINUE
      S = 1.0D0/DASUM(N,Z,1)
      CALL DSCAL (N, S, Z, 1)
      YNORM = S*YNORM
C                                  SOLVE U*Z = V
      DO 50  K=N, 1, -1
         IF (DABS(Z(K)) .GT. DABS(FAC(K,K))) THEN
            S = DABS(FAC(K,K))/DABS(Z(K))
            CALL DSCAL (N, S, Z, 1)
            YNORM = S*YNORM
         END IF
         IF (FAC(K,K) .NE. 0.0D0) THEN
            Z(K) = Z(K)/FAC(K,K)
         ELSE
            Z(K) = 1.0D0
         END IF
         T = -Z(K)
         CALL DAXPY (K-1, T, FAC(1,K), 1, Z(1), 1)
   50 CONTINUE
C                                  MAKE ZNORM = 1.0
      S = 1.0D0/DASUM(N,Z,1)
      CALL DSCAL (N, S, Z, 1)
      YNORM = S*YNORM
      IF (ANORM .NE. 0.0D0) RCOND = YNORM/ANORM
C
      DDD=DMACH(4,NFLAG)
      IF (NFLAG.NE.0) GOTO 9999
      IF (RCOND .LE. DDD) THEN
         CALL E1STD (1, RCOND)
         IF (NFLAG.NE.0) GOTO 9999
         CALL E1MES (3, 1, 'The matrix is algorithmically '//
     &               'singular.  An estimate of the reciprocal '//
     &       'of its L1 condition number is RCOND = %(D1).',NFLAG)
         IF (NFLAG.NE.0) GOTO 9999
         NFLAGBIS=220
      END IF
C
 9000 CALL E1POP ('DL2CRG ',NFLAG)
      IF (NFLAG.NE.0) GOTO 9999
      NFLAG=NFLAGBIS
 9999 RETURN
      END
