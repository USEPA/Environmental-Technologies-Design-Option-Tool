C-----------------------------------------------------------------------
C  IMSL Name:  M1VECH
C
C  Computer:   pcdsms/SINGLE
C
C  Revised:    December 31, 1984
C
C  Purpose:    Character substring assignment.
C
C  Usage:      CALL M1VECH (STR1, LEN1, STR2, LEN2)
C
C  Arguments:
C     STR1   - Source substring.  (Input)
C              The source substring is STR1(1:LEN1).
C     LEN1   - Length of STR1.  (Input)
C     STR2   - Destination substring.  (Output)
C              The destination substring is STR2(1:LEN2).
C     LEN2   - Length of STR2.  (Input)
C
C  Copyright:  1984 by IMSL, Inc.  All rights reserved.
C
C  Warranty:   IMSL warrants only that IMSL testing has been applied
C              to this code.  No other warranty, expressed or implied,
C              is applicable.
C
C-----------------------------------------------------------------------
C
      SUBROUTINE M1VECH (STR1, LEN1, STR2, LEN2)
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    LEN1, LEN2
      CHARACTER  STR1*(*), STR2*(*)
C
      STR2(1:LEN2) = STR1(1:LEN1)
C
      RETURN
      END
