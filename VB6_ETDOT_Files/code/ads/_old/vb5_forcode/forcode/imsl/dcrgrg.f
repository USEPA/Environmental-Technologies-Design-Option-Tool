C-----------------------------------------------------------------------
C  IMSL Name:  CRGRG/DCRGRG (Single/Double precision version)
C
C  Computer:   pcdsms/DOUBLE
C
C  Revised:    June 5, 1985
C
C  Purpose:    Copy a real general matrix.
C
C  Usage:      CALL CRGRG (N, A, LDA, B, LDB)
C
C  Arguments:
C     N      - Order of the matrices.  (Input)
C     A      - Matrix of order N.  (Input)
C     LDA    - Leading dimension of A exactly as specified in the
C              dimension statement of the calling program.  (Input)
C     B      - Matrix of order N containing a copy of A.  (Output)
C     LDB    - Leading dimension of B exactly as specified in the
C              dimension statement of the calling program.  (Input)
C
C  GAMS:       D1b8
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
      SUBROUTINE DCRGRG (N, A, LDA, B, LDB,NFLAG)
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    N, LDA, LDB,NFLAG,NFLAGBIS
      DOUBLE PRECISION A(LDA,*), B(LDB,*)
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    J
C                                  SPECIFICATIONS FOR SUBROUTINES
      EXTERNAL   E1MES, E1POP, E1PSH, E1STI, DCOPY
C
      NFLAGBIS=0
      CALL E1PSH ('DCRGRG ',NFLAG)
      IF (NFLAG.NE.0) GOTO 9999
C                                  Check N
      IF (N .LT. 1) THEN
         CALL E1STI (1, N)
         CALL E1MES (5, 1, 'The argument N = %(I1).  It must be at '//
     &               'least 1.',NFLAG)
	 IF (NFLAG.NE.0) GOTO 9999
	 NFLAGBIS=110
         GO TO 9000
      END IF
C                                  Check LDA
      IF (LDA .LT. N) THEN
         CALL E1STI (1, LDA)
         CALL E1STI (2, N)
         CALL E1MES (5, 2, 'The argument LDA = %(I1).  It must be at '//
     &               'least as large as N = %(I2).',NFLAG)
         IF (NFLAG.NE.0) GOTO 9999
	 NFLAGBIS=111
         GO TO 9000
      END IF
C                                  Check LDB
      IF (LDB .LT. N) THEN
         CALL E1STI (1, LDB)
         CALL E1STI (2, N)
         CALL E1MES (5, 3, 'The argument LDB = %(I1).  It must be at '//
     &               'least as large as N = %(I2).',NFLAG)
         IF (NFLAG.NE.0) GOTO 9999
	 NFLAGBIS=112
         GO TO 9000
      END IF
C                                  Copy
      IF (LDA.EQ.N .AND. LDB.EQ.N) THEN
         CALL DCOPY (N*N, A, 1, B, 1)
      ELSE IF (LDA .GE. LDB) THEN
         DO 10  J=1, N
            CALL DCOPY (N, A(1,J), 1, B(1,J), 1)
   10    CONTINUE
      ELSE
         DO 20  J=N, 1, -1
            CALL DCOPY (N, A(1,J), -1, B(1,J), -1)
   20    CONTINUE
      END IF
C
 9000 CONTINUE
      CALL E1POP ('DCRGRG ',NFLAG)
      IF (NFLAG.NE.0) GOTO 9999
      NFLAG=NFLAGBIS
 9999 RETURN
      END
