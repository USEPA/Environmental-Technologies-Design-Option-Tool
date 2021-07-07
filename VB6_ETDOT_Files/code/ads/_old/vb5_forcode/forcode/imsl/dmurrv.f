C-----------------------------------------------------------------------
C  IMSL Name:  MURRV/DMURRV (Single/Double precision version)
C
C  Computer:   pcdsms/DOUBLE
C
C  Revised:    August 8, 1985
C
C  Purpose:    Multiply a real rectangular matrix by a vector.
C
C  Usage:      CALL MURRV (NRA, NCA, A, LDA, NX, X, IPATH, NY, Y)
C
C  Arguments:
C     NRA    - Number of rows of A.  (Input)
C     NCA    - Number of columns of A.  (Input)
C     A      - Real NRA by NCA matrix in full storage mode.  (Input)
C     LDA    - Leading dimension of A exactly as specified in the
C              dimension statement of the calling program.  (Input)
C     NX     - Length of the vector X.  (Input)
C              NX must be equal to NCA if IPATH is equal to 1.
C              NX must be equal to NRA if IPATH is equal to 2.
C     X      - Real vector of length NX.  (Input)
C     IPATH  - Integer flag.  (Input)
C              IPATH = 1 means the product Y = A*X is computed.
C              IPATH = 2 means the product Y = trans(A)*X is computed
C              where trans(A) is the transpose of A.
C     NY     - Length of the vector Y.  (Input)
C              NY must be equal to NRA if IPATH is equal to 1.
C              NY must be equal to NCA if IPATH is equal to 2.
C     Y      - Real vector of length NY containing the product A*X if
C              IPATH is equal to 1 and the product trans(A)*X if IPATH
C              is equal to 2.  (Output)
C
C  GAMS:       D1b4
C
C  Chapter:    MATH/LIBRARY Basic Matrix/Vector Operations
C
C  Copyright:  1985 by IMSL, Inc.  All Rights Reserved.
C
C  Warranty:   IMSL warrants only that IMSL testing has been applied
C              to this code.  No other warranty, expressed or implied,
C              is applicable.
C
C-----------------------------------------------------------------------
C
      SUBROUTINE DMURRV (NRA, NCA,A,LDA,NX,X,IPATH,NY,Y,NFLAG)
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    NRA, NCA, LDA, NX, IPATH, NY,NFLAG,NFLAGBIS
      DOUBLE PRECISION A(LDA,*), X(*), Y(*)
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      CHARACTER  TRANS*1
C                                  SPECIFICATIONS FOR SUBROUTINES
      EXTERNAL   E1MES, E1POP, E1PSH, E1STI, DGEMV
C                                  SPECIFICATIONS FOR FUNCTIONS
      EXTERNAL   N1RTY
      INTEGER    N1RTY,II
C
      NFLAGBIS=0
      CALL E1PSH ('DMURRV ',NFLAG)
      IF (NFLAG.NE.0) GOTO 9999
C
      IF (NRA.LE.0 .OR. NCA.LE.0) THEN
         CALL E1STI (1, NRA)
         CALL E1STI (2, NCA)
         CALL E1MES (5, 1, 'Both the number of rows and the '//
     &               'number of columns of the input matrix have '//
     &               'to be positive while NRA = %(I1) and '//
     &               'NCA = %(I2) are given.',NFLAG)
         NFLAGBIS=NFLAG
         NFLAG=0
         IF (NFLAGBIS.NE.0) GOTO 9000
         NFLAGBIS=300
         GOTO 9000
      END IF
C
      IF (NRA .GT. LDA) THEN
         CALL E1STI (1, NRA)
         CALL E1STI (2, LDA)
         CALL E1MES (5, 2, 'The number of rows of the matrix '//
     &               'must be less than or equal to its leading '//
     &               'dimension while NRA = %(I1) and LDA = %(I2) '//
     &               'are given.',NFLAG)
        NFLAGBIS=NFLAG
         NFLAG=0
         IF (NFLAGBIS.NE.0) GOTO 9000
        NFLAGBIS=301
        GOTO 9000
      END IF
C
      IF (IPATH.NE.1 .AND. IPATH.NE.2) THEN
         CALL E1STI (1, IPATH)
         CALL E1MES (5, 3, 'IPATH must equal to 1 or 2 while '//
     &               'a value of %(I1) is given.',NFLAG)
         NFLAGBIS=NFLAG
         NFLAG=0
         IF (NFLAGBIS.NE.0) GOTO 9000
         NFLAGBIS=302
         GOTO 9000
      END IF
C
      II=N1RTY(0,NFLAG)
      NFLAGBIS=NFLAG
         NFLAG=0
         IF (NFLAGBIS.NE.0) GOTO 9000
      IF (II.NE. 0) GO TO 9000
C
      IF (IPATH.EQ.1 .AND. (NX.NE.NCA.OR.NY.NE.NRA)) THEN
         CALL E1STI (1, NRA)
         CALL E1STI (2, NCA)
         CALL E1STI (3, NX)
         CALL E1STI (4, NY)
         CALL E1MES (5, 4, 'When IPATH=1, the number of rows '//
     &               'in A must be the same as the length of Y '//
     &               'and the number of columns in A must be the '//
     &               'same as the length of X, but NRA=%(I1), '//
     &               'NY=%(I4), NCA=%(I2), and NX=%(I3).',NFLAG)
         NFLAGBIS=NFLAG
         NFLAG=0
         IF (NFLAGBIS.NE.0) GOTO 9000
         NFLAGBIS=303
         GOTO 9000
      END IF
C
      IF (IPATH.EQ.2 .AND. (NX.NE.NRA.OR.NY.NE.NCA)) THEN
         CALL E1STI (1, NRA)
         CALL E1STI (2, NCA)
         CALL E1STI (3, NX)
         CALL E1STI (4, NY)
         CALL E1MES (5, 5, 'When IPATH=2, the number of rows '//
     &               'in A must be the same as the length of X '//
     &               'and the number of columns in A must be the '//
     &               'same as the length of Y, but NRA=%(I1), '//
     &               'NX=%(I3), NCA=%(I2), and NY=%(I4).',NFLAG)
         NFLAGBIS=NFLAG
         NFLAG=0
         IF (NFLAGBIS.NE.0) GOTO 9000
         NFLAGBIS=304
         GOTO 9000
      END IF
C
      II=N1RTY(0,NFLAG)
      NFLAGBIS=NFLAG
         NFLAG=0
         IF (NFLAGBIS.NE.0) GOTO 9000
      IF (II.NE. 0) GO TO 9000
C                                  Set flag TRANS
      IF (IPATH .EQ. 1) THEN
         TRANS = 'N'
      ELSE IF (IPATH .EQ. 2) THEN
         TRANS = 'T'
      END IF
C                                  Compute  Y = A*X or Y = (A**T)*X
      CALL DGEMV (TRANS, NRA, NCA, 1.0D0, A, LDA, X, 1, 0.0D0, Y, 1)
C
 9000 CALL E1POP ('DMURRV ',NFLAG)
      IF (NFLAG.NE.0) GOTO 9999
      NFLAG=NFLAGBIS
 9999 RETURN
      END
