C-----------------------------------------------------------------------
C  IMSL Name:  AMACH (Single precision version)
C
C  Computer:   pcdsms/SINGLE
C
C  Revised:    March 15, 1984
C
C  Purpose:    Retrieve single-precision machine constants.
C
C  Usage:      AMACH(N)
C
C  Arguments:
C     N      - Index of desired constant.  (Input)
C     AMACH  - Machine constant.  (Output)
C              AMACH(1) = B**(EMIN-1), the smallest positive magnitude.
C              AMACH(2) = B**EMAX*(1 - B**(-T)), the largest magnitude.
C              AMACH(3) = B**(-T), the smallest relative spacing.
C              AMACH(4) = B**(1-T), the largest relative spacing.
C              AMACH(5) = LOG10(B), the log, base 10, of the radix.
C              AMACH(6) = not-a-number.
C              AMACH(7) = positive machine infinity.
C              AMACH(8) = negative machine infinity.
C
C  GAMS:       R1
C
C  Chapters:   MATH/LIBRARY Reference Material
C              STAT/LIBRARY Reference Material
C              SFUN/LIBRARY Reference Material
C
C  Copyright:  1984 by IMSL, Inc.  All Rights Reserved.
C
C  Warranty:   IMSL warrants only that IMSL testing has been applied
C              to this code.  No other warranty, expressed or implied,
C              is applicable.
C
C-----------------------------------------------------------------------
C
      REAL FUNCTION AMACH (N,NFLAG)
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    N,NFLAG
C                                  SPECIFICATIONS FOR SAVE VARIABLES
      REAL       RMACH(8)
      SAVE       RMACH
C                                  SPECIFICATIONS FOR SUBROUTINES
      EXTERNAL   E1MES, E1POP, E1PSH, E1STI
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    IRMACH(8)
C
      EQUIVALENCE (RMACH, IRMACH)
C                                  DEFINE CONSTANTS
      DATA RMACH(1)/1.17577E-38/
      DATA RMACH(2)/3.40204E38/
      DATA RMACH(3)/5.96184E-8/
      DATA RMACH(4)/1.19237E-7/
      DATA RMACH(5)/.301029995663981195E0/
      DATA IRMACH(6)/2139091960/
      DATA RMACH(7)/3.40204E38/
      DATA RMACH(8)/-3.40204E38/
C
      IF (N.LT.1 .OR. N.GT.8) THEN
         CALL E1PSH ('AMACH ',NFLAG)
	 IF (NFLAG.NE.0) GOTO 9999
         AMACH = RMACH(6)
         CALL E1STI (1, N)
         CALL E1MES (5, 5, 'The argument must be between 1 '//
     &               'and 8 inclusive. N = %(I1)',NFLAG)	
         CALL E1POP ('AMACH ',NFLAG)
	 IF (NFLAG.NE.0) GOTO 9999
	 NFLAG=100
      ELSE
         AMACH = RMACH(N)
      END IF
C
9999  RETURN
      END
