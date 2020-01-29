C-----------------------------------------------------------------------
C  IMSL Name:  IMACH (Single precision version)
C
C  Computer:   pcdsms/SINGLE
C
C  Revised:    March 26, 1984
C
C  Purpose:    Retrieve integer machine constants.
C
C  Usage:      IMACH(N)
C
C  Arguments:
C     N      - Index of desired constant.  (Input)
C     IMACH  - Machine constant.  (Output)
C
C  Remark:
C     Following is a description of the assorted integer machine
C     constants.
C
C     Words
C
C        IMACH( 1) = Number of bits per integer storage unit.
C        IMACH( 2) = Number of characters per integer storage unit.
C
C     Integers
C
C        Assume integers are represented in the S-DIGIT, BASE-A form
C        SIGN ( X(S-1)*A**(S-1) + ... + X(1)*A + X(0) )
C        where 0 .LE. X(I) .LT. A for I=0,...,S-1.  Then
C
C        IMACH( 3) = A, the base.
C        IMACH( 4) = S, number of BASE-A digits.
C        IMACH( 5) = A**S - 1, largest magnitude.
C
C     Floating-point numbers
C
C        Assume floating-point numbers are represented in the T-DIGIT,
C        BASE-B form SIGN (B**E)*( (X(1)/B) + ... + (X(T)/B**T) )
C        where 0 .LE. X(I) .LT. B for I=1,...,T,
C        0 .LT. X(1), and EMIN .LE. E .LE. EMAX.  Then
C
C        IMACH( 6) = B, the base.
C
C        Single precision
C
C           IMACH( 7) = T, number of BASE-B digits.
C           IMACH( 8) = EMIN, smallest exponent E.
C           IMACH( 9) = EMAX, largest exponent E.
C
C        Double precision
C
C           IMACH(10) = T, number of BASE-B digits.
C           IMACH(11) = EMIN, smallest exponent E.
C           IMACH(12) = EMAX, largest exponent E.
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
      INTEGER FUNCTION IMACH (N,NFLAG)
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    N,NFLAG
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    NOUT,NFLAGB
C                                  SPECIFICATIONS FOR SAVE VARIABLES
      INTEGER    IMACHV(12)
      SAVE       IMACHV
      
C                                  SPECIFICATIONS FOR SUBROUTINES
      EXTERNAL   UMACH
C                                  DEFINE CONSTANTS
      DATA IMACHV(1)/32/
      DATA IMACHV(2)/4/
      DATA IMACHV(3)/2/
      DATA IMACHV(4)/31/
      DATA IMACHV(5)/2147483647/
      DATA IMACHV(6)/2/
      DATA IMACHV(7)/24/
      DATA IMACHV(8)/-125/
      DATA IMACHV(9)/128/
      DATA IMACHV(10)/53/
      DATA IMACHV(11)/-1021/
      DATA IMACHV(12)/1024/
C

      IF (N.LT.1 .OR. N.GT.12) THEN
C                                  ERROR.  INVALID RANGE FOR N.
         CALL UMACH (2, NOUT,NFLAGB)
         IF (NFLAGB.NE.0) THEN
             NFLAG=NFLAGB
             GOTO 1000
         ENDIF
c         WRITE (NOUT,99999) N
99999    FORMAT (/, ' *** TERMINAL ERROR 5 from IMACH.  The argument',
     &          /, ' ***          must be between 1 and 12 inclusive.'
     &          , /, ' ***          N = ', I6, '.', /)
         IMACH = 0
c         STOP
        NFLAG=99
C
      ELSE
         IMACH = IMACHV(N)
      END IF
C
1000  RETURN
      END
