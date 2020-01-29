C-----------------------------------------------------------------------
C  IMSL Name:  LINRT/DLINRT  (Single/Double precision version)
C
C  Computer:   pcdsms/DOUBLE
C
C  Revised:    January 1, 1985
C
C  Purpose:    Compute the inverse of a real triangular matrix.
C
C  Usage:      CALL LINRT (N, A, LDA, IPATH, AINV, LDAINV)
C
C  Arguments:
C     N      - Order of the matrix.  (Input)
C     A      - N by N matrix containing the triangular matrix to be
C              inverted in full storage form.  (Input)
C              For a lower triangular matrix, only the lower triangular
C              part and diagonal of A are referenced.  For an upper
C              triangular matrix, only the upper triangular part and
C              diagonal of A are referenced.
C     LDA    - Leading dimension of A exactly as specified in the
C              dimension statement of the calling program.  (Input)
C     IPATH  - Path indicator.  (Input)
C              IPATH = 1 means A is lower triangular,
C              IPATH = 2 means A is upper triangular.
C     AINV   - N by N matrix containing the inverse of A.  (Output)
C              If A is lower triangular, AINV is also lower triangular.
C              If A is upper triangular, AINV is also upper triangular.
C              If A is not needed, A and AINV can share the same storage
C              locations.
C     LDAINV - Leading dimension of AINV exactly as specified in the
C              dimension statement of the calling program.  (Input)
C
C  GAMS:       D2a3
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
      SUBROUTINE DLINRT (N,A,LDA,IPATH,AINV,LDAINV,NFLAG)
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    N, LDA, IPATH, LDAINV,NFLAG,NFLAGBIS
      DOUBLE PRECISION A(LDA,*), AINV(LDAINV,*)
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    INFO, J, K
      DOUBLE PRECISION BIG, SMALL, TEMP
C                                  SPECIFICATIONS FOR INTRINSICS
C     INTRINSIC  DABS
      INTRINSIC  DABS
      DOUBLE PRECISION DABS
C                                  SPECIFICATIONS FOR SUBROUTINES
      EXTERNAL   E1MES, E1POP, E1PSH, E1STI, DCOPY, DGER, DSCAL, DSET
C                                  SPECIFICATIONS FOR FUNCTIONS
      EXTERNAL   DMACH, N1RCD
      INTEGER    N1RCD,II
      DOUBLE PRECISION DMACH,DDD
C
      NFLAGBIS=0
      CALL E1PSH ('DLINRT ',NFLAG)
      IF (NFLAG.NE.0) GOTO 9999
C
      IF (N .LE. 0) THEN
         CALL E1STI (1, N)
         CALL E1MES (5, 1, 'The order of the matrix must be '//
     &               'positive while N = %(I1) is given.',NFLAG)
         NFLAGBIS=NFLAG
         NFLAG=0
         IF (NFLAGBIS.NE.0) GOTO 9000
         NFLAGBIS=312
         GOTO 9000   
      ELSE IF (N .GT. LDA) THEN
         CALL E1STI (1, N)
         CALL E1STI (2, LDA)
         CALL E1MES (5, 2, 'The order of the matrix must be '//
     &               'less than or equal to its leading dimension '//
     &       'while N = %(I1) and LDA = %(I2) are given.',NFLAG)
         NFLAGBIS=NFLAG
         NFLAG=0
         IF (NFLAGBIS.NE.0) GOTO 9000
         NFLAGBIS=313
         GOTO 9000
      ELSE IF (N .GT. LDAINV) THEN
         CALL E1STI (1, N)
         CALL E1STI (2, LDAINV)
         CALL E1MES (5, 3, 'The order of the matrix must be '//
     &               'less than or equal to its leading dimension '//
     &        'while N = %(I1) and LDAINV = %(I2) are given.',NFLAG)
         NFLAGBIS=NFLAG
         NFLAG=0
         IF (NFLAGBIS.NE.0) GOTO 9000
         NFLAGBIS=314
         GOTO 9000
      ELSE IF (IPATH.NE.1 .AND. IPATH.NE.2) THEN
         CALL E1STI (1, IPATH)
         CALL E1MES (5, 4, 'IPATH must be either 1 or 2 while '//
     &               'a value of %(I1) is given.',NFLAG)
         NFLAGBIS=NFLAG
         NFLAG=0
         IF (NFLAGBIS.NE.0) GOTO 9000
         NFLAGBIS=315
         GOTO 9000
      END IF
      II=N1RCD(0,NFLAG)
      NFLAGBIS=NFLAG
      NFLAG=0
      IF (NFLAGBIS.NE.0) GOTO 9000
      IF (II.NE. 0) GO TO 9000
C
      DDD=DMACH(1,NFLAG)
      NFLAGBIS=NFLAG
      NFLAG=0
      IF (NFLAGBIS.NE.0) GOTO 9000
      SMALL = DDD
      DDD=DMACH(2,NFLAG)
      NFLAGBIS=NFLAG
      NFLAG=0
      IF (NFLAGBIS.NE.0) GOTO 9000
      BIG = DDD
      IF (SMALL*BIG .LT. 1.0D0) SMALL = 1.0D0/BIG
      IF (IPATH .EQ. 1) THEN
C                                  MAKE A COPY OF A IN AINV AND ZERO
C                                  THE STRICTLY UPPER TRIANGLE OF AINV
         DO 10  J=1, N
            CALL DSET (J-1, 0.0D0, AINV(1,J), 1)
            CALL DCOPY (N-J+1, A(J,J), 1, AINV(J,J), 1)
   10    CONTINUE
C                                  COMPUTE INVERSE OF LOWER TRIANGULAR
C                                  MATRIX
         DO 20  K=N, 1, -1
            INFO = K
            IF (DABS(AINV(K,K)) .LE. SMALL) GO TO 50
            AINV(K,K) = 1.0D0/AINV(K,K)
            TEMP = -AINV(K,K)
            IF (K .LT. N) THEN
               CALL DSCAL (N-K, TEMP, AINV(K+1,K), 1)
               CALL DGER (N-K, K-1, 1.0D0, AINV(K+1,K), 1, AINV(K,1),
     &                    LDAINV, AINV(K+1,1), LDAINV)
            END IF
            CALL DSCAL (K-1, AINV(K,K), AINV(K,1), LDAINV)
   20    CONTINUE
         INFO = 0
      ELSE IF (IPATH .EQ. 2) THEN
C                                  MAKE A COPY OF A IN AINV AND ZERO
C                                  THE STRICTLY LOWER TRIANGLE OF AINV
         DO 30  J=1, N
            IF (J .LT. N) CALL DSET (N-J, 0.0D0, AINV(J+1,J), 1)
            CALL DCOPY (J, A(1,J), 1, AINV(1,J), 1)
   30    CONTINUE
C                                  COMPUTE INVERSE OF AN UPPER
C                                  TRIANGULAR MATRIX
         DO 40  K=1, N
            INFO = K
            IF (DABS(AINV(K,K)) .LE. SMALL) GO TO 50
            AINV(K,K) = 1.0D0/AINV(K,K)
            TEMP = -AINV(K,K)
            CALL DSCAL (K-1, TEMP, AINV(1,K), 1)
            IF (K .LT. N) THEN
               CALL DGER (K-1, N-K, 1.0D0, AINV(1,K), 1, AINV(K,K+1),
     &                    LDAINV, AINV(1,K+1), LDAINV)
               CALL DSCAL (N-K, AINV(K,K), AINV(K,K+1), LDAINV)
            END IF
   40    CONTINUE
         INFO = 0
      END IF
C
   50 IF (INFO .NE. 0) THEN
         CALL E1STI (1, INFO)
         CALL E1MES (5, 5, 'The matrix to be inverted is '//
     &               'singular.  The index of the first zero '//
     &               'diagonal element of A is %(I1).',NFLAG)
        NFLAGBIS=NFLAG
        NFLAG=0
        IF (NFLAGBIS.NE.0) GOTO 9000
        NFLAGBIS=316
      END IF
C
 9000 CALL E1POP ('DLINRT ',NFLAG)
      IF (NFLAG.NE.0) GOTO 9999
      NFLAGBIS=NFLAG
 9999 RETURN
      END
