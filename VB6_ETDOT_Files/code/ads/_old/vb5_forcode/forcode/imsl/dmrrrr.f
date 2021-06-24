C-----------------------------------------------------------------------
C  IMSL Name:  MRRRR/DMRRRR (Single/Double precision version)
C
C  Computer:   pcdsms/DOUBLE
C
C  Revised:    August 8, 1985
C
C  Purpose:    Multiply two real rectangular matrices, A*B.
C
C  Usage:      CALL MRRRR (NRA, NCA, A, LDA, NRB, NCB, B, LDB,
C                          NRC, NCC, C, LDC)
C
C  Arguments:
C     NRA    - Number of rows of A.  (Input)
C     NCA    - Number of columns of A.  (Input)
C     A      - Real NRA by NCA matrix in full storage mode.  (Input)
C     LDA    - Leading dimension of A exactly as specified in the
C              dimension statement of the calling program.  (Input)
C     NRB    - Number of rows of B.  (Input)
C              NRB must be equal to NCA.
C     NCB    - Number of columns of B.  (Input)
C     B      - Real NRB by NCB matrix in full storage mode.  (Input)
C     LDB    - Leading dimension of B exactly as specified in the
C              dimension statement of the calling program.  (Input)
C     NRC    - Number of rows of C.  (Input)
C              NRC must be equal to NRA.
C     NCC    - Number of columns of C.  (Input)
C              NCC must be equal to NCB.
C     C      - Real NRC by NCC matrix containing the product A*B in full
C              storage mode.  (Output)
C     LDC    - Leading dimension of C exactly as specified in the
C              dimension statement of the calling program.  (Input)
C
C  GAMS:       D1b6
C
C  Chapters:   MATH/LIBRARY Basic Matrix/Vector Operations
C              STAT/LIBRARY Mathematical Support
C
C  Copyright:  1986 by IMSL, Inc.  All Rights Reserved.
C
C  Warranty:   IMSL warrants only that IMSL testing has been applied
C              to this code.  No other warranty, expressed or implied,
C              is applicable.
C
C-----------------------------------------------------------------------
C
      SUBROUTINE DMRRRR (NRA, NCA, A, LDA, NRB, NCB, B, LDB, NRC, NCC,
     &                   C, LDC,NFLAG)
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    NRA, NCA, LDA, NRB, NCB, LDB, NRC, NCC, LDC
      INTEGER NFLAG,NFLAGBIS
      DOUBLE PRECISION A(LDA,*), B(LDB,*), C(LDC,*)
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    IPATH, J
C                                  SPECIFICATIONS FOR SUBROUTINES
      EXTERNAL   E1MES, E1POP, E1PSH, E1STI, DMURRV
C                                  SPECIFICATIONS FOR FUNCTIONS
      EXTERNAL   N1RTY
      INTEGER    N1RTY,II
C
      NFLAGBIS=0
      CALL E1PSH ('DMRRRR ',NFLAG)
      IF (NFLAG.NE.0) GOTO 9999
C
      IF (NRA.LE.0 .OR. NCA.LE.0) THEN
         CALL E1STI (1, NRA)
         CALL E1STI (2, NCA)
         CALL E1MES (5, 1, 'Both the number of rows and the '//
     &               'number of columns of a matrix have '//
     &               'to be positive while NRA = %(I1) and '//
     &               'NCA = %(I2) are given.',NFLAG)
         NFLAGBIS=NFLAG
         NFLAG=0
         IF (NFLAGBIS.NE.0) GOTO 9000
         NFLAGBIS=305
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
        NFLAGBIS=306
        GOTO 9000
      END IF
C
      IF (NRB.LE.0 .OR. NCB.LE.0) THEN
         CALL E1STI (1, NRB)
         CALL E1STI (2, NCB)
         CALL E1MES (5, 3, 'Both the number of rows and the '//
     &               'number of columns of a matrix have '//
     &               'to be positive while NRB = %(I1) and '//
     &               'NCB = %(I2) are given.',NFLAG)
         NFLAGBIS=NFLAG
         NFLAG=0
         IF (NFLAGBIS.NE.0) GOTO 9000
         NFLAGBIS=307
         GOTO 9000
      END IF
C
      IF (NRB .GT. LDB) THEN
         CALL E1STI (1, NRB)
         CALL E1STI (2, LDB)
         CALL E1MES (5, 4, 'The number of rows of the matrix '//
     &               'must be less than or equal to its leading '//
     &               'dimension while NRB = %(I1) and LDB = %(I2) '//
     &               'are given.',NFLAG)
         NFLAGBIS=NFLAG
         NFLAG=0
         IF (NFLAGBIS.NE.0) GOTO 9000
         NFLAGBIS=308
         GOTO 9000
      END IF
C
      IF (NRC.LE.0 .OR. NCC.LE.0) THEN
         CALL E1STI (1, NRC)
         CALL E1STI (2, NCC)
         CALL E1MES (5, 5, 'Both the number of rows and the '//
     &               'number of columns of a matrix have '//
     &               'to be positive while NRC = %(I1) and '//
     &               'NCC = %(I2) are given.',NFLAG)
         NFLAGBIS=NFLAG
         NFLAG=0
         IF (NFLAGBIS.NE.0) GOTO 9000
         NFLAGBIS=309
         GOTO 9000
      END IF
C
      IF (NRC .GT. LDC) THEN
         CALL E1STI (1, NRC)
         CALL E1STI (2, LDC)
         CALL E1MES (5, 6, 'The number of rows of the matrix '//
     &               'must be less than or equal to its leading '//
     &               'dimension while NRC = %(I1) and LDC = %(I2) '//
     &               'are given.',NFLAG)
         NFLAGBIS=NFLAG
         NFLAG=0
         IF (NFLAGBIS.NE.0) GOTO 9000
         NFLAGBIS=310
         GOTO 9000
      END IF
C--------------------------------------------------
      II=N1RTY(0,NFLAG)
      NFLAGBIS=NFLAG
        NFLAG=0
        IF (NFLAGBIS.NE.0) GOTO 9000
      IF (II.NE. 0) GO TO 9000
C
      IF (NRB.NE.NCA .OR. NRC.NE.NRA .OR. NCC.NE.NCB) THEN
         CALL E1STI (1, NRA)
         CALL E1STI (2, NCA)
         CALL E1STI (3, NRB)
         CALL E1STI (4, NCB)
         CALL E1STI (5, NRC)
         CALL E1STI (6, NCC)
         CALL E1MES (5, 7, 'Some of the dimensions are not '//
     &               'consistent.  The following must hold NRB = '//
     &               'NCA, NRC = NRA and NCC = NCB while NRA = '//
     &               '%(I1), NCA = %(I2), NRB = %(I3), NCB = %(I4), '//
     &               'NRC = %(I5) and NCC = %(I6) are given.',NFLAG)
        NFLAGBIS=NFLAG
        NFLAG=0
        IF (NFLAGBIS.NE.0) GOTO 9000
        NFLAGBIS=311
        GOTO 9000
      END IF
C
      II=N1RTY(0,NFLAG)
      NFLABIS=NFLAG
        NFLAG=0
        IF (NFLAGBIS.NE.0) GOTO 9000
      IF (II.NE. 0) GO TO 9000
C
      IPATH = 1
      DO 10  J=1, NCC
         CALL DMURRV (NRA, NCA, A, LDA, NRB, B(1,J), IPATH, NRC,
     &                C(1,J),NFLAG)
      NFLAGBIS=NFLAG
      NFLAG=0
      IF (NFLAGBIS.NE.0) GOTO 9000
   10 CONTINUE
C
 9000 CALL E1POP ('DMRRRR ',NFLAG)
      IF (NFLAG.NE.0) GOTO 9999
      NFLAG=NFLAGBIS
 9999 RETURN
      END
