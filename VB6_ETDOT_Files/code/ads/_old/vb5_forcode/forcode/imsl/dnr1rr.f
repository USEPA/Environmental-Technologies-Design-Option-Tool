C-----------------------------------------------------------------------
C  IMSL Name:  NR1RR/DNR1RR (Single/Double precision version)
C
C  Computer:   pcdsms/DOUBLE
C
C  Revised:    October 17, 1985
C
C  Purpose:    Compute the 1-norm of a real matrix.
C
C  Usage:      CALL NR1RR (NRA, NCA, A, LDA, ANORM)
C
C  Arguments:
C     NRA    - Number of rows of A.  (Input)
C     NCA    - Number of columns of A.  (Input)
C     A      - Real NRA by NCA matrix whose 1-norm is to be computed.
C              (Input)
C     LDA    - Leading dimension of A exactly as specified in the
C              dimension statement of the calling program.  (Input)
C     ANORM  - Real scalar containing the 1-norm of A.  (Output)
C
C  GAMS:       D1b2
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
      SUBROUTINE DNR1RR (NRA, NCA, A, LDA, ANORM,NFLAG)
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    NRA, NCA, LDA,NFLAG,NFLAGBIS,II
      DOUBLE PRECISION ANORM, A(LDA,*)
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    J
      DOUBLE PRECISION ANORM1
C                                  SPECIFICATIONS FOR INTRINSICS
C     INTRINSIC  DMAX1
      INTRINSIC  DMAX1
      DOUBLE PRECISION DMAX1
C                                  SPECIFICATIONS FOR SUBROUTINES
      EXTERNAL   E1MES, E1POP, E1PSH, E1STI
C                                  SPECIFICATIONS FOR FUNCTIONS
      EXTERNAL   N1RCD, DASUM
      INTEGER    N1RCD
      DOUBLE PRECISION DASUM
C
      NFLAGBIS=0
      CALL E1PSH ('DNR1RR ',NFLAG)
	IF (NFLAG.NE.0) GOTO 9999
C                                  CHECK FOR INPUT ERRORS
      IF (NRA .GT. LDA) THEN
         CALL E1STI (1, NRA)
         CALL E1STI (2, LDA)
         CALL E1MES (5, 1, 'The number of rows of the input '//
     &               'matrix must be less than or equal to the '//
     &               'leading dimension while NRA = %(I1) and '//
     &               'LDA = %(I2) are given.',NFLAG)
	IF (NFLAG.NE.0) GOTO 9999
	NFLAGBIS=144
         GO TO 9000
      END IF
C
      IF (NRA .LE. 0) THEN
         CALL E1STI (1, NRA)
         CALL E1MES (5, 2, 'The number of rows of the input '//
     &               'matrix must be greater than zero while '//
     &               'NRA = %(I1) is given.',NFLAG)
	IF (NFLAG.NE.0) GOTO 9999
	NFLAGBIS=145
      END IF
C
      IF (NCA .LE. 0) THEN
         CALL E1STI (1, NCA)
         CALL E1MES (5, 3, 'The number of columns of the input '//
     &               'matrix must be greater than zero while '//
     &               'NCA = %(I1) is given.',NFLAG)
	IF (NFLAG.NE.0) GOTO 9999
	NFLAGBIS=146
      END IF
C
	II=N1RCD(0,NFLAG)
	IF (NFLAG.NE.0) GOTO 9999
      IF (II.NE. 0) GO TO 9000
C                                  CALCULATE THE L1 NORM FOR A.
      ANORM = 0.0D0
      DO 10  J=1, NCA
         ANORM1 = DASUM(NRA,A(1,J),1)
         ANORM = DMAX1(ANORM1,ANORM)
   10 CONTINUE
C
 9000 CALL E1POP ('DNR1RR ',NFLAG)
	IF (NFLAG.NE.0) GOTO 9999
	NFLAG=NFLAGBIS
 9999 RETURN
      END
