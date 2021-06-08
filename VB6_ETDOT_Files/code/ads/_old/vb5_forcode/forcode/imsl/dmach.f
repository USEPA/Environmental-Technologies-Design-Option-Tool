C-----------------------------------------------------------------------
C  IMSL Name:  DMACH (Double precision version)
C
C  Computer:   pcdsms/DOUBLE
C
C  Revised:    March 15, 1984
C
C  Purpose:    Generate double precision machine constants.
C
C  Usage:      DMACH(N)
C
C  Arguments:
C     N      - Index of desired constant.  (Input)
C     DMACH  - Machine constant.  (Output)
C              DMACH(1) = B**(EMIN-1), the smallest positive magnitude.
C              DMACH(2) = B**EMAX*(1 - B**(-T)), the largest magnitude.
C              DMACH(3) = B**(-T), the smallest relative spacing.
C              DMACH(4) = B**(1-T), the largest relative spacing.
C              DMACH(5) = LOG10(B), the log, base 10, of the radix.
C              DMACH(6) = not-a-number.
C              DMACH(7) = positive machine infinity.
C              DMACH(8) = negative machine infinity.
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
      DOUBLE PRECISION FUNCTION DMACH (N,NFLAG)
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    N,NFLAG,NFLAGBIS
C                                  SPECIFICATIONS FOR SAVE VARIABLES
      DOUBLE PRECISION RMACH(8)
      SAVE       RMACH
C                                  SPECIFICATIONS FOR SUBROUTINES
      EXTERNAL   E1MES, E1POP, E1PSH, E1STI
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    IRMACH(16)
C
      EQUIVALENCE (RMACH, IRMACH)
C                                  DEFINE CONSTANTS
      DATA RMACH(1)/2.22559D-308/
      DATA RMACH(2)/1.79728D308/
      DATA RMACH(3)/1.11048D-16/
      DATA RMACH(4)/2.22096D-16/
      DATA RMACH(5)/.3010299956639811952137388947245D0/
      DATA IRMACH(11)/0/
      DATA IRMACH(12)/1206910591/
      DATA RMACH(7)/1.79728D308/
      DATA RMACH(8)/-1.79728D308/
C
      NFLAGBIS=0
      IF (N.LT.1 .OR. N.GT.8) THEN
         CALL E1PSH ('DMACH ',NFLAG)
         IF (NFLAG.NE.0) GOTO 1000
         DMACH = RMACH(6)
         CALL E1STI (1, N)
         CALL E1MES (5, 5, 'The argument must be between 1 '//
     &               'and 8 inclusive. N = %(I1)',NFLAG)
         NFLAGBIS=NFLAG
         NFLAG=0
         CALL E1POP ('DMACH ',NFLAG)
         IF (NFLAG.NE.0) GOTO 1000
         NFLAGBIS=8
      ELSE
         DMACH = RMACH(N)
      END IF
C
1000  IF (NFLAG.EQ.0) THEN 
       NFLAG=NFLAGBIS 
      ENDIF
      RETURN
      END
