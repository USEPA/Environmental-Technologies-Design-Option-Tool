C-----------------------------------------------------------------------
C  IMSL Name:  L2NRG/DL2NRG (Single/Double precision version)
C
C  Computer:   pcdsms/DOUBLE
C
C  Revised:    February 27, 1985
C
C  Purpose:    Compute the inverse of a real general matrix.
C
C  Usage:      CALL L2NRG (N, A, LDA, AINV, LDAINV, WK, IWK)
C
C  Arguments:  See LINRG/DLINRG.
C
C  Remarks:    See LINRG/DLINRG.
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
      SUBROUTINE DL2NRG (N, A, LDA, AINV, LDAINV, WK, IWK,NFLAG)
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    N, LDA, LDAINV, IWK(*),NFLAG,NFLAGBIS
      DOUBLE PRECISION A(LDA,*), AINV(LDAINV,*), WK(*)
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    I, INC, J, K, L
      DOUBLE PRECISION RCOND
C                                  SPECIFICATIONS FOR SUBROUTINES
      EXTERNAL   E1MES, E1POP, E1PSH, E1STI, E1STD, DCOPY, DGEMV,
     &           DSET, DSWAP, DL2CRG, DLINRT
C                                  SPECIFICATIONS FOR FUNCTIONS
      EXTERNAL   DMACH, N1RTY
      INTEGER    N1RTY,II
      DOUBLE PRECISION DMACH,DDD
C
      NFLAGBIS=0
      CALL E1PSH ('DL2NRG ',NFLAG)
      IF (NFLAG.NE.0) GOTO 9999
C
      IF (N .LE. 0) THEN
         CALL E1STI (1, N)
         CALL E1MES (5, 1, 'The order of the matrix must be '//
     &               'positive while N = %(I1) is given.',NFLAG)
        NFLAGBIS=NFLAG
        NFLAG=0
        IF (NFLAGBIS.NE.0) GOTO 9000
        NFLAGBIS=317
        GO TO 9000
      END IF
C
      IF (N .GT. LDA) THEN
         CALL E1STI (1, N)
         CALL E1STI (2, LDA)
         CALL E1MES (5, 2, 'The order of the matrix must be '//
     &               'less than or equal to its leading dimension '//
     &           'while N = %(I1) and LDA = %(I2) are given.',NFLAG)
         NFLAGBIS=NFLAG
         NFLAG=0
         IF (NFLAGBIS.NE.0) GOTO 9000
         NFLAGBIS=318
         GO TO 9000
      END IF
C
      IF (N .GT. LDAINV) THEN
         CALL E1STI (1, N)
         CALL E1STI (2, LDAINV)
         CALL E1MES (5, 3, 'The order of the matrix must be '//
     &               'less than or equal to its leading dimension '//
     &          'while N = %(I1) and LDAINV = %(I2) are given.',NFLAG)
         NFLAGBIS=NFLAG
         NFLAG=0
         IF (NFLAGBIS.NE.0) GOTO 9000
         NFLAGBIS=319
         GO TO 9000
      END IF
C                                  COMPUTE THE LU FACTORIZATION OF A
C                                  AND ESTIMATE ITS CONDITION NUMBER
      INC = N*(N-1)/2
      CALL DL2CRG (N, A,LDA,AINV,LDAINV,IWK,RCOND,WK(INC+1),NFLAG)
      NFLAGBIS=NFLAG
      NFLAG=0
      IF (NFLAGBIS.NE.0) GOTO 9000
      II=N1RTY(1,NFLAG)
      NFLAGBIS=NFLAG
      NFLAG=0
      IF (NFLAGBIS.NE.0) GOTO 9000
      IF (II.EQ. 4) GO TO 9000
C                                  COMPUTE INVERSE(U)
      J = INC
      K = 0
      DO 10  I=1, N - 1
         J = J - K
         CALL DCOPY (I, AINV(N-I+1,N-I), 1, WK(J), 1)
         K = I + 1
   10 CONTINUE
C
      CALL DLINRT (N, AINV, LDAINV,2,AINV,LDAINV,NFLAG)
      NFLAGBIS=NFLAG
      NFLAG=0
      IF (NFLAGBIS.NE.0) GOTO 9000
C
      J = INC
      K = 0
      DO 20  I=1, N - 1
         J = J - K
         CALL DCOPY (I, WK(J), 1, AINV(N-I+1,N-I), 1)
         K = I + 1
   20 CONTINUE
C                                  FORM INVERSE(U)*INVERSE(L)
      DO 30  K=N - 1, 1, -1
         CALL DCOPY (N-K, AINV(K+1,K), 1, WK(K+1+INC), 1)
         CALL DSET (N-K, 0.0D0, AINV(K+1,K), 1)
         CALL DGEMV ('N', N, N-K, 1.0D0, AINV(1,K+1), LDAINV,
     &               WK(K+1+INC), 1, 1.0D0, AINV(1,K), 1)
         L = IWK(K)
         IF (L .NE. K) CALL DSWAP (N, AINV(1,K), 1, AINV(1,L), 1)
   30 CONTINUE
C
      DDD=DMACH(4,NFLAG)
      NFLAGBIS=NFLAG
      NFLAG=0
      IF (NFLAGBIS.NE.0) GOTO 9000
      IF (RCOND .LE. DDD) THEN
      CALL E1STD (1, RCOND)
      NFLAGBIS=NFLAG
      NFLAG=0
      IF (NFLAGBIS.NE.0) GOTO 9000
         CALL E1MES (3, 1, 'The matrix is too ill-conditioned. '//
     &               'An estimate of the reciprocal of its L1 '//
     &               'condition number is RCOND = %(D1).  '//
     &               'The inverse might not be accurate.',NFLAG)
        NFLAGBIS=NFLAG
        NFLAG=0
        IF (NFLAGBIS.NE.0) GOTO 9000
        NFLAGBIS=320
      END IF
C
 9000 CALL E1POP ('DL2NRG ',NFLAG)
      IF (NFLAG.NE.0) GOTO 9999
      NFLAG=NFLAGBIS
 9999 RETURN
      END
