C-----------------------------------------------------------------------
C  IMSL Name:  I1ERIF
C
C  Computer:   pcdsms/SINGLE
C
C  Revised:    March 13, 1984
C
C  Purpose:    Return the position of the first element of a given
C              character array which is not an element of another
C              character array.
C
C  Usage:      I1ERIF(STR1, LEN1, STR2, LEN2)
C
C  Arguments:
C     STR1   - Character array to be searched.  (Input)
C     LEN1   - Length of STR1.  (Input)
C     STR2   - Character array to be searched for.  (Input)
C     LEN2   - Length of STR2.  (Input)
C     I1ERIF - Integer function.  (Output)
C
C  Copyright:  1984 by IMSL, Inc.  All rights reserved.
C
C  Warranty:   IMSL warrants only that IMSL testing has been applied
C              to this code.  No other warranty, expressed or implied,
C              is applicable.
C
C-----------------------------------------------------------------------
C
      INTEGER FUNCTION I1ERIF (STR1, LEN1, STR2, LEN2)
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    LEN1, LEN2
      CHARACTER  STR1(*), STR2(*)
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    I
C                                  SPECIFICATIONS FOR FUNCTIONS
      EXTERNAL   I1X
      INTEGER    I1X
C                              FIRST EXECUTABLE STATEMENT
      IF (LEN1.LE.0 .OR. LEN2.LE.0) THEN
         I1ERIF = 1
      ELSE
         DO 10  I=1, LEN1
            IF (I1X(STR2,LEN2,STR1(I),1) .EQ. 0) THEN
               I1ERIF = I
               RETURN
            END IF
   10    CONTINUE
         I1ERIF = 0
      END IF
C
      RETURN
      END
