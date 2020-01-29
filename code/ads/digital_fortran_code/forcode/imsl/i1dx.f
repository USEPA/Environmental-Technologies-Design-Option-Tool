C-----------------------------------------------------------------------
C  IMSL Name:  I1DX (Single precision version)
C
C  Computer:   pcdsms/SINGLE
C
C  Revised:    September 9, 1985
C
C  Purpose:    Determine the array subscript indicating the starting
C              element at which a key character sequence begins.
C              (Case-insensitive version)
C
C  Usage:      I1DX(CHRSTR, I1LEN, KEY, KLEN)
C
C  Arguments:
C     CHRSTR - Character array to be searched.  (Input)
C     I1LEN  - Length of CHRSTR.  (Input)
C     KEY    - Character array that contains the key sequence.  (Input)
C     KLEN   - Length of KEY.  (Input)
C     I1DX   - Integer function.  (Output)
C
C  Remarks:
C  1. Returns zero when there is no match.
C
C  2. Returns zero if KLEN is longer than ISLEN.
C
C  3. Returns zero when any of the character arrays has a negative or
C     zero length.
C
C  GAMS:       N5c
C
C  Chapter:    MATH/LIBRARY Utilities
C
C  Copyright:  1985 by IMSL, Inc.  All Rights Reserved.
C
C  Warranty:   IMSL warrants only that IMSL testing has been applied
C              to this code.  No other warranty, expressed or implied,
C              is applicable.
C
C-----------------------------------------------------------------------
C
      INTEGER FUNCTION I1DX (CHRSTR, I1LEN, KEY, KLEN)
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    I1LEN, KLEN
      CHARACTER  CHRSTR(*), KEY(*)
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    I, II, J
C                                  SPECIFICATIONS FOR FUNCTIONS
      EXTERNAL   ICASE, I1CSTR
      INTEGER    ICASE, I1CSTR
C
      I1DX = 0
      IF (KLEN.LE.0 .OR. I1LEN.LE.0) GO TO 9000
      IF (KLEN .GT. I1LEN) GO TO 9000
C
      I = 1
      II = I1LEN - KLEN + 1
   10 IF (I .LE. II) THEN
         IF (ICASE(CHRSTR(I)) .EQ. ICASE(KEY(1))) THEN
            IF (KLEN .NE. 1) THEN
               J = KLEN - 1
               IF (I1CSTR(CHRSTR(I+1),J,KEY(2),J) .EQ. 0) THEN
                  I1DX = I
                  GO TO 9000
               END IF
            ELSE
               I1DX = I
               GO TO 9000
            END IF
         END IF
         I = I + 1
         GO TO 10
      END IF
C
 9000 RETURN
      END
