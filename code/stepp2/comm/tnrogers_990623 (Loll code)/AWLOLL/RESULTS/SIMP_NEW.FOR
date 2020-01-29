C-----------------------------------------------------------------------
C  IMSL Name:  DTIME (Single precision version)
C
C  Computer:   pcdsms/SINGLE
C
C  Revised:    November 5, 1985
C
C  Purpose:    Get time of day.
C
C  Usage:      CALL DTIME (IHOUR, MINUTE, ISEC)
C
C  Arguments:
C     IHOUR  - Hour of the day.  (Output)
C              IHOUR is between 0 and 23 inclusive.
C     MINUTE - Minute within the hour.  (Output)
C              MINUTE is between 0 and 59 inclusive.
C     ISEC   - Second within the minute.  (Output)
C              ISEC is between 0 and 59 inclusive.
C
C  GAMS:       R
C
C  Chapter:    Reference Material - Time utilities
C
C  Copyright:  1985 by IMSL, Inc.  All Rights Reserved.
C
C  Warranty:   IMSL warrants only that IMSL testing has been applied
C              to this code.  No other warranty, expressed or implied,
C              is applicable.
C
C-----------------------------------------------------------------------
C
      SUBROUTINE DTIME (IHOUR, MINUTE, ISEC)
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    IHOUR, MINUTE, ISEC
C
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    *2 IH, IM, IS, IHS
C                                  SPECIFICATIONS FOR SUBROUTINES
      EXTERNAL   GETTIM
C
      CALL GETTIM (IH, IM, IS, IHS)
      IHOUR = IH
      MINUTE = IM
      ISEC = IS
      RETURN
      END
C-----------------------------------------------------------------------
C  IMSL Name:  ISMAX (Single precision version)
C
C  Computer:   pcdsms/SINGLE
C
C  Revised:    August 9, 1986
C
C  Purpose:    Find the smallest index of the component of a
C              single-precision vector having maximum value.
C
C  Usage:      ISMAX(N, SX, INCX)
C
C  Arguments:
C     N      - Length of vector X.  (Input)
C     SX     - Real vector of length N*INCX.  (Input)
C     INCX   - Displacement between elements of SX.  (Input)
C              X(I) is defined to be SX(1+(I-1)*INCX). INCX must be
C              greater than zero.
C     ISMAX  - The smallest index I such that X(I)  is the maximum of
C              X(J) for J=1 to N.  (Output)
C              X(I) refers to a specific element of SX. See INCX
C              argument description.
C
C  GAMS:       D1a2
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
      INTEGER FUNCTION ISMAX (N, SX, INCX)
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    N, INCX
      REAL       SX(*)
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    I, IX
      REAL       SMAX
C
      ISMAX = 0
      IF (N .GE. 1) THEN
         ISMAX = 1
         IF (N .NE. 1) THEN
            IF (INCX .NE. 1) THEN
C                                  CODE FOR INCREMENT NOT EQUAL TO 1
               IX = 1
               SMAX = SX(1)
               IX = IX + INCX
               DO 10  I=2, N
                  IF (SX(IX) .GT. SMAX) THEN
                     ISMAX = I
                     SMAX = SX(IX)
                  END IF
                  IX = IX + INCX
   10          CONTINUE
            ELSE
C                                  CODE FOR INCREMENT EQUAL TO 1
               SMAX = SX(1)
               DO 20  I=2, N
                  IF (SX(I) .GT. SMAX) THEN
                     ISMAX = I
                     SMAX = SX(I)
                  END IF
   20          CONTINUE
            END IF
         END IF
      END IF
      RETURN
      END
C-----------------------------------------------------------------------
C  IMSL Name:  ISMIN (Single precision version)
C
C  Computer:   pcdsms/SINGLE
C
C  Revised:    August 9, 1986
C
C  Purpose:    Find the smallest index of the component of a
C              single-precision vector having minimum value.
C
C  Usage:      ISMIN(N, SX, INCX)
C
C  Arguments:
C     N      - Length of vector X.  (Input)
C     SX     - Real vector of length N*INCX.  (Input)
C     INCX   - Displacement between elements of SX.  (Input)
C              X(I) is defined to be SX(1+(I-1)*INCX). INCX must be
C              greater than zero.
C     ISMIN  - The smallest index I such that X(I) is the minimum of
C              of X(J) for J=1 to N.  (Output)
C              X(I) refers to a specific element of SX. See INCX
C              argument description.
C
C  GAMS:       D1a2
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
      INTEGER FUNCTION ISMIN (N, SX, INCX)
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    N, INCX
      REAL       SX(*)
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    I, IX
      REAL       SMIN
C
      ISMIN = 0
      IF (N .GE. 1) THEN
         ISMIN = 1
         IF (N .NE. 1) THEN
            IF (INCX .NE. 1) THEN
C                                  CODE FOR INCREMENT NOT EQUAL TO 1
               IX = 1
               SMIN = SX(1)
               IX = IX + INCX
               DO 10  I=2, N
                  IF (SX(IX) .LT. SMIN) THEN
                     ISMIN = I
                     SMIN = SX(IX)
                  END IF
                  IX = IX + INCX
   10          CONTINUE
            ELSE
C                                  CODE FOR INCREMENT EQUAL TO 1
               SMIN = SX(1)
               DO 20  I=2, N
                  IF (SX(I) .LT. SMIN) THEN
                     ISMIN = I
                     SMIN = SX(I)
                  END IF
   20          CONTINUE
            END IF
         END IF
      END IF
      RETURN
      END
C-----------------------------------------------------------------------
C  IMSL Name:  R1CLK (Single precision version)
C
C  Computer:   pcdsms/SINGLE
C
C  Revised:    July 11, 1985
C
C  Purpose:    Set a value in the range (1, 2147483646) using the
C              system clock.
C
C  Usage:      CALL R1CLK (IISEED)
C
C  Argument:
C     IISEED - A value in the range (1, 2147483646) computed from a
C              call to the system clock.  (Output)
C
C  GAMS:       L6c
C
C  Chapter:    STAT/LIBRARY Not-user-callable
C
C  Copyright:  1984 by IMSL, Inc.  All Rights Reserved.
C
C  Warranty:   IMSL warrants only that IMSL testing has been applied
C              to this code.  No other warranty, expressed or implied,
C              is applicable.
C
C-----------------------------------------------------------------------
C
      SUBROUTINE R1CLK (IISEED)
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    IISEED
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    IHOUR, MINUTE, ISEC
C                                  SPECIFICATIONS FOR SUBROUTINES
      EXTERNAL   DTIME
C                                  Set seed using the system clock
      CALL DTIME (IHOUR, MINUTE, ISEC)
      IISEED = IHOUR*10000 + MINUTE*100 + ISEC + 1
      IF (IISEED .LE. 1) THEN
C                                  Seed is nonrandomly set.  Use of the
C                                  system clock has not been devloped.
         IISEED = 123457
      END IF
      RETURN
      END
C-----------------------------------------------------------------------
C  IMSL Name:  R1INS/DR1INS (Single/Double precision version)
C
C  Computer:   pcdsms/SINGLE
C
C  Revised:    April 19, 1988
C
C  Purpose:    Randomly initialize the table for the IMSL shuffled
C              random number generators.
C
C  Usage:      CALL R1INS
C
C  Arguments:  (none)
C
C  GAMS:       L6c
C
C  Chapter:    STAT/LIBRARY Random Number Generation (not documented)
C
C  Copyright:  1987 by IMSL, Inc.  All Rights Reserved.
C
C  Warranty:   IMSL warrants only that IMSL testing has been applied
C              to this code.  No other warranty, expressed or implied,
C              is applicable.
C
C-----------------------------------------------------------------------
C
      SUBROUTINE R1INS
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    I
      DOUBLE PRECISION DSEED1, DSEED2
C                                  SPECIFICATIONS FOR SAVE VARIABLES
      LOGICAL    FIRST
      DOUBLE PRECISION D2P31M
      SAVE       D2P31M, FIRST
C                                  SPECIFICATIONS FOR SPECIAL CASES
C                                  SPECIFICATIONS FOR COMMON /R2NCOM/
      COMMON     /R2NCOM/ D2P31A, DSEED, DWK, DINTTB, INDCTR, INTTB, WK
      DOUBLE PRECISION D2P31A, DSEED, DWK(128)
      REAL       WK(128)
      INTEGER    INDCTR
      LOGICAL    DINTTB, INTTB
C                                  SPECIFICATIONS FOR INTRINSICS
C     INTRINSIC  DMOD
      INTRINSIC  DMOD
      DOUBLE PRECISION DMOD
C                                  SPECIFICATIONS FOR SUBROUTINES
      EXTERNAL   R1INT
C
      DATA D2P31M/2147483647.0D0/
      DATA FIRST/.TRUE./
C
      IF (FIRST) THEN
         CALL R1INT (0)
         FIRST = .FALSE.
      END IF
C
      IF (INDCTR .EQ. 2) THEN
C                                  Use the multiplier 16807
         DO 10  I=1, 128
            DSEED = DMOD(16807.0D0*DSEED,D2P31M)
            WK(I) = DSEED/D2P31A
   10    CONTINUE
      ELSE IF (INDCTR .EQ. 4) THEN
C                                  Use the multiplier 397204094
         DO 20  I=1, 128
            DSEED1 = DMOD(32768.0D0*DSEED,D2P31M)
            DSEED2 = DMOD(23166.0D0*DSEED,D2P31M)
            DSEED = DMOD(12121.0D0*DSEED1+DSEED2,D2P31M)
            WK(I) = DSEED/D2P31A
   20    CONTINUE
      ELSE IF (INDCTR .EQ. 6) THEN
C                                  Use the multiplier 950706376
         DO 30  I=1, 128
            DSEED1 = DMOD(32768.0D0*DSEED,D2P31M)
            DSEED2 = DMOD(8392.0D0*DSEED,D2P31M)
            DSEED = DMOD(29013.0D0*DSEED1+DSEED2,D2P31M)
            WK(I) = DSEED/D2P31A
   30    CONTINUE
      END IF
      RETURN
      END
C-----------------------------------------------------------------------
C  IMSL Name:  R1INT (Single precision version)
C
C  Computer:   pcdsms/SINGLE
C
C  Revised:    July 15, 1985
C
C  Purpose:    Initialize variables in common used by the IMSL
C              multiplicative congruential random number generators.
C
C  Usage:      CALL R1INT (IOPT)
C
C  Argument:
C     IOPT   - Option indicator.  (Input)
C              IOPT       ACTION
C              Positive   INDCTR in COMMON is set to IOPT.
C               0         A check is made to see if the seed has been
C                         set.  If a seed has not been set, it is
C                         initialized using the system clock.
C              -1         The seed is being set in the calling routine.
C              -2         No action, other than the initialization (on
C                         the first call only) of the variables in
C                         COMMON that the user has no control over.
C
C  Remark:
C     On the first call to R1INT, the default indicator is set and the
C     scaling factor for (1,2**31-1) into (0,1) is determined.
C
C  GAMS:       L6c
C
C  Chapter:    STAT/LIBRARY Random Number Generation (not documented)
C
C  Copyright:  1985 by IMSL, Inc.  All Rights Reserved.
C
C  Warranty:   IMSL warrants only that IMSL testing has been applied
C              to this code.  No other warranty, expressed or implied,
C              is applicable.
C
C-----------------------------------------------------------------------
C
      SUBROUTINE R1INT (IOPT)
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    IOPT
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    IISEED
      REAL       TEMP
C                                  SPECIFICATIONS FOR SAVE VARIABLES
      LOGICAL    FIRST, SETSED
      SAVE       FIRST, SETSED
C                                  SPECIFICATIONS FOR SPECIAL CASES
C                                  SPECIFICATIONS FOR COMMON /R2NCOM/
      COMMON     /R2NCOM/ D2P31A, DSEED, DWK, DINTTB, INDCTR, INTTB, WK
      DOUBLE PRECISION D2P31A, DSEED, DWK(128)
      REAL       WK(128)
      INTEGER    INDCTR
      LOGICAL    DINTTB, INTTB
C                                  SPECIFICATIONS FOR INTRINSICS
C     INTRINSIC  DBLE
      INTRINSIC  DBLE
      DOUBLE PRECISION DBLE
C                                  SPECIFICATIONS FOR SUBROUTINES
      EXTERNAL   R1CLK
C
      DATA FIRST/.TRUE./, SETSED/.FALSE./
C
      IF (FIRST) THEN
C                                  Determine scaling factor for U(0,1)
         D2P31A = 2147483647.0D0
         TEMP = 2147483646.0D0/D2P31A
         IF (TEMP .GE. 1.0) THEN
            D2P31A = 2147483655.0D0
            TEMP = 2147483646.0D0/D2P31A
            IF (TEMP .GE. 1.0) THEN
               D2P31A = 2147483711.0D0
   10          CONTINUE
               TEMP = 2147483646.0D0/D2P31A
               IF (TEMP .GE. 1.0) THEN
                  D2P31A = D2P31A + 8.0D0
                  GO TO 10
               END IF
            END IF
         END IF
         INDCTR = 1
         INTTB = .FALSE.
         DINTTB = .FALSE.
         FIRST = .FALSE.
      END IF
      IF (IOPT .EQ. -1) THEN
C                                  Seed is being set by calling routine
         SETSED = .TRUE.
      ELSE IF (IOPT .EQ. 0) THEN
         IF (.NOT.SETSED) THEN
C                                  Set seed using the system clock
            CALL R1CLK (IISEED)
            DSEED = DBLE(IISEED)
            SETSED = .TRUE.
         END IF
      ELSE IF (IOPT .GT. 0) THEN
C                                  Select the Uniform random number
C                                  generator
         INDCTR = IOPT
      END IF
      RETURN
      END
C-----------------------------------------------------------------------
C  IMSL Name:  RNSET (Single precision version)
C
C  Computer:   pcdsms/SINGLE
C
C  Revised:    July 11, 1985
C
C  Purpose:    Initialize a random seed for use in the IMSL random
C              number generators.
C
C  Usage:      CALL RNSET (ISEED)
C
C  Argument:
C     ISEED  - The seed of the random number generator.  (Input)
C              ISEED must be in the range (0, 2147483646).  If ISEED is
C              zero, a value is computed using the system clock; and,
C              hence, the results of programs using the IMSL random
C              number generators will be different at different times.
C
C  GAMS:       L6c
C
C  Chapters:   STAT/LIBRARY Random Number Generation
C              MATH/LIBRARY Utilities
C
C  Copyright:  1984 by IMSL, Inc.  All Rights Reserved.
C
C  Warranty:   IMSL warrants only that IMSL testing has been applied
C              to this code.  No other warranty, expressed or implied,
C              is applicable.
C
C-----------------------------------------------------------------------
C
      SUBROUTINE RNSET (ISEED)
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    ISEED
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    IISEED
C                                  SPECIFICATIONS FOR SPECIAL CASES
C                                  SPECIFICATIONS FOR COMMON /R2NCOM/
      COMMON     /R2NCOM/ D2P31A, DSEED, DWK, DINTTB, INDCTR, INTTB, WK
      DOUBLE PRECISION D2P31A, DSEED, DWK(128)
      REAL       WK(128)
      INTEGER    INDCTR
      LOGICAL    DINTTB, INTTB
C                                  SPECIFICATIONS FOR INTRINSICS
C     INTRINSIC  DBLE
      INTRINSIC  DBLE
      DOUBLE PRECISION DBLE
C                                  SPECIFICATIONS FOR SUBROUTINES
      EXTERNAL   R1CLK, R1INT
C
      IF (ISEED.LT.0 .OR. ISEED.GE.2147483647) THEN
         CALL R1INT (0)
      ELSE IF (ISEED .EQ. 0) THEN
         CALL R1CLK (IISEED)
         DSEED = DBLE(IISEED)
      ELSE
         CALL R1INT (-1)
         DSEED = DBLE(ISEED)
      END IF
      RETURN
      END
C-----------------------------------------------------------------------
C  IMSL Name:  RNUNF/DRNUNF (Single/Double precision version)
C
C  Computer:   pcdsms/SINGLE
C
C  Revised:    April 19, 1988
C
C  Purpose:    Generate a pseudorandom number from a uniform (0,1)
C              distribution.
C
C  Usage:      RNUNF()
C
C  Argument:
C     RNUNF  - Function value, a random uniform (0,1) deviate.  (Output)
C
C  Remarks:
C  1. The IMSL routine RNSET can be used to initialize the seed of the
C     random number generator.  The routine RNOPT can be used to select
C     the form of the generator.
C
C  2. This function has a side effect: it changes the value of the seed,
C     which is passed through a common block.
C
C  Keywords:   Utilities; Monte Carlo; Simulation; Rectangular
C              distribution
C
C  GAMS:       L6a21
C
C  Chapters:   STAT/LIBRARY Random Number Generation
C              MATH/LIBRARY Utilities
C
C  Copyright:  1987 by IMSL, Inc.  All Rights Reserved.
C
C  Warranty:   IMSL warrants only that IMSL testing has been applied
C              to this code.  No other warranty, expressed or implied,
C              is applicable.
C
C-----------------------------------------------------------------------
C
      REAL FUNCTION RNUNF ()
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    J
      REAL       X
      DOUBLE PRECISION DSEED1, DSEED2
C                                  SPECIFICATIONS FOR SAVE VARIABLES
      LOGICAL    FIRST
      DOUBLE PRECISION D2P31M
      SAVE       D2P31M, FIRST
C                                  SPECIFICATIONS FOR SPECIAL CASES
C                                  SPECIFICATIONS FOR COMMON /R2NCOM/
      COMMON     /R2NCOM/ D2P31A, DSEED, DWK, DINTTB, INDCTR, INTTB, WK
      DOUBLE PRECISION D2P31A, DSEED, DWK(128)
      REAL       WK(128)
      INTEGER    INDCTR
      LOGICAL    DINTTB, INTTB
C                                  SPECIFICATIONS FOR INTRINSICS
C     INTRINSIC  DMOD
      INTRINSIC  DMOD
      DOUBLE PRECISION DMOD
C                                  SPECIFICATIONS FOR SUBROUTINES
      EXTERNAL   R1INT, R1INS
C
      DATA FIRST/.TRUE./
      DATA D2P31M/2147483647.0D0/
C
      IF (FIRST) THEN
         CALL R1INT (0)
         FIRST = .FALSE.
      END IF
C
      IF (INDCTR .EQ. 1) THEN
C                                  USE THE MULTIPLIER 16807 (=7**5)
         DSEED = DMOD(16807.0D0*DSEED,D2P31M)
         RNUNF = DSEED/D2P31A
      ELSE IF (INDCTR .EQ. 2) THEN
C                                  Use the multiplier 16807 with
C                                  shuffling
         IF (.NOT.INTTB) THEN
C                                  Initialize table
            CALL R1INS
            INTTB = .TRUE.
         END IF
C                                  Generate a new Uniform (0,1) deviate
         DSEED = DMOD(16807.0D0*DSEED,D2P31M)
C                                  Now shuffle
         J = DMOD(DSEED,128.0D0) + 1.0D0
         X = DSEED/D2P31A
         RNUNF = WK(J)
         WK(J) = X
      ELSE IF (INDCTR .EQ. 3) THEN
C                                  Use the multiplier 397204094
         DSEED1 = DMOD(32768.0D0*DSEED,D2P31M)
         DSEED2 = DMOD(23166.0D0*DSEED,D2P31M)
         DSEED = DMOD(12121.0D0*DSEED1+DSEED2,D2P31M)
         RNUNF = DSEED/D2P31A
      ELSE IF (INDCTR .EQ. 4) THEN
C                                  Use the multiplier 397204094 with
C                                  shuffling
         IF (.NOT.INTTB) THEN
C                                  Initialize table
            CALL R1INS
            INTTB = .TRUE.
         END IF
C                                  Generate a new Uniform (0,1) deviate
         DSEED1 = DMOD(32768.0D0*DSEED,D2P31M)
         DSEED2 = DMOD(23166.0D0*DSEED,D2P31M)
         DSEED = DMOD(12121.0D0*DSEED1+DSEED2,D2P31M)
C                                  Now shuffle
         J = DMOD(DSEED,128.0D0) + 1.0D0
         X = DSEED/D2P31A
         RNUNF = WK(J)
         WK(J) = X
      ELSE IF (INDCTR .EQ. 5) THEN
C                                  Use the multiplier 950706376
         DSEED1 = DMOD(32768.0D0*DSEED,D2P31M)
         DSEED2 = DMOD(8392.0D0*DSEED,D2P31M)
         DSEED = DMOD(29013.0D0*DSEED1+DSEED2,D2P31M)
         RNUNF = DSEED/D2P31A
      ELSE IF (INDCTR .EQ. 6) THEN
C                                  Use the multiplier 950706376 with
C                                  shuffling
         IF (.NOT.INTTB) THEN
C                                  Initialize table
            CALL R1INS
            INTTB = .TRUE.
         END IF
C                                  Generate a new Uniform (0,1) deviate
         DSEED1 = DMOD(32768.0D0*DSEED,D2P31M)
         DSEED2 = DMOD(8392.0D0*DSEED,D2P31M)
         DSEED = DMOD(29013.0D0*DSEED1+DSEED2,D2P31M)
C                                  Now shuffle
         J = DMOD(DSEED,128.0D0) + 1.0D0
         X = DSEED/D2P31A
         RNUNF = WK(J)
         WK(J) = X
      END IF
      RETURN
      END
C-----------------------------------------------------------------------
C  IMSL Name:  SCOPY (Single precision version)
C
C  Computer:   pcdsms/SINGLE
C
C  Revised:    August 9, 1986
C
C  Purpose:    Copy a vector X to a vector Y, both single precision.
C
C  Usage:      CALL SCOPY (N, SX, INCX, SY, INCY)
C
C  Arguments:
C     N      - Length of vectors X and Y.  (Input)
C     SX     - Real vector of length MAX(N*IABS(INCX),1).  (Input)
C     INCX   - Displacement between elements of SX.  (Input)
C              X(I) is defined to be.. SX(1+(I-1)*INCX) if INCX .GE. 0
C              or SX(1+(I-N)*INCX) if INCX .LT. 0.
C     SY     - Real vector of length MAX(N*IABS(INCY),1).  (Output)
C              SCOPY copies X(I) to Y(I) for I=1,...,N. X(I) and Y(I)
C              refer to specific elements of SX and SY, respectively.
C              See INCX and INCY argument descriptions.
C     INCY   - Displacement between elements of SY.  (Input)
C              Y(I) is defined to be.. SY(1+(I-1)*INCY) if INCY .GE. 0
C              or SY(1+(I-N)*INCY) if INCY .LT. 0.
C
C  GAMS:       D1a
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
      SUBROUTINE SCOPY (N, SX, INCX, SY, INCY)
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    N, INCX, INCY
      REAL       SX(*), SY(*)
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    I, IX, IY, M, MP1
C                                  SPECIFICATIONS FOR SPECIAL CASES
C     INTRINSIC  MOD
      INTRINSIC  MOD
      INTEGER    MOD
C
      IF (N .GT. 0) THEN
         IF (INCX.NE.1 .OR. INCY.NE.1) THEN
C                                  CODE FOR UNEQUAL INCREMENTS
            IX = 1
            IY = 1
            IF (INCX .LT. 0) IX = (-N+1)*INCX + 1
            IF (INCY .LT. 0) IY = (-N+1)*INCY + 1
            DO 10  I=1, N
               SY(IY) = SX(IX)
               IX = IX + INCX
               IY = IY + INCY
   10       CONTINUE
         ELSE
C                                  CODE FOR BOTH INCREMENTS EQUAL TO 1
            M = MOD(N,7)
C                                  CLEAN-UP LOOP
            DO 30  I=1, M
               SY(I) = SX(I)
   30       CONTINUE
            MP1 = M + 1
            DO 40  I=MP1, N, 7
               SY(I) = SX(I)
               SY(I+1) = SX(I+1)
               SY(I+2) = SX(I+2)
               SY(I+3) = SX(I+3)
               SY(I+4) = SX(I+4)
               SY(I+5) = SX(I+5)
               SY(I+6) = SX(I+6)
   40       CONTINUE
         END IF
      END IF
      RETURN
      END
C-----------------------------------------------------------------------
C  IMSL Name:  SSUM (Single precision version)
C
C  Computer:   pcdsms/SINGLE
C
C  Revised:    August 9, 1986
C
C  Purpose:    Sum the values of a single precision vector.
C
C  Usage:      SSUM(N, SX, INCX)
C
C  Arguments:
C     N      - Length of vectors X.  (Input)
C     SX     - Real vector of length N*INCX.  (Input)
C     INCX   - Displacement between elements of SX.  (Input)
C              X(I) is defined to be SX(1+(I-1)*INCX). INCX must be
C              greater than 0.
C     SSUM   - Single precision sum from I=1 to N of X(I).  (Output)
C              X(I) refers to a specific element of SX.
C
C  GAMS:       D1a
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
      REAL FUNCTION SSUM (N, SX, INCX)
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    N, INCX
      REAL       SX(*)
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    I, M, MP1, NINCX
C                                  SPECIFICATIONS FOR SPECIAL CASES
C     INTRINSIC  MOD
      INTRINSIC  MOD
      INTEGER    MOD
C
      SSUM = 0.0E0
      IF (N .GT. 0) THEN
         IF (INCX .NE. 1) THEN
C                                  CODE FOR INCREMENT NOT EQUAL TO 1
            NINCX = N*INCX
            DO 10  I=1, NINCX, INCX
               SSUM = SSUM + SX(I)
   10       CONTINUE
         ELSE
C                                  CODE FOR INCREMENT EQUAL TO 1
            M = MOD(N,6)
C                                  CLEAN-UP LOOP
            DO 30  I=1, M
               SSUM = SSUM + SX(I)
   30       CONTINUE
            MP1 = M + 1
            DO 40  I=MP1, N, 6
               SSUM = SSUM + SX(I) + SX(I+1) + SX(I+2) + SX(I+3) +
     &                SX(I+4) + SX(I+5)
   40       CONTINUE
         END IF
      END IF
      RETURN
      END
C-----------------------------------------------------------------------
C  IMSL Name:  U2POL/DU2POL (Single/Double precision version)
C
C  Computer:   pcdsms/SINGLE
C
C  Revised:    November 1, 1984
C
C  Purpose:    Minimize a function of N variables using a direct search
C              polytope algorithm.
C
C  Usage:      CALL U2POL (FCN, N, XGUESS, S, FTOL, MAXFCN, X, FVALUE,
C                          WK)
C
C  Arguments:  (See UMPOL)
C
C  Chapter:    MATH/LIBRARY Optimization
C
C  Copyright:  1984 by IMSL, Inc.  All Rights Reserved.
C
C  Warranty:   IMSL warrants only that IMSL testing has been applied
C              to this code.  No other warranty, expressed or implied,
C              is applicable.
C
C-----------------------------------------------------------------------
C
      SUBROUTINE U2POL (FCN, N, XGUESS, S, FTOL, MAXFCN, X, FVALUE, WK)
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    N, MAXFCN
      REAL       S, FTOL, FVALUE, XGUESS(*), X(*), WK(*)
      EXTERNAL   FCN
C                                  SPECIFICATIONS FOR SUBROUTINES
      EXTERNAL   E1MES, E1POP, E1PSH, E1STI, E1STR, SCOPY, U3POL
C                                  FIRST EXECUTABLE STATEMENT
      CALL E1PSH ('U2POL ')
C                                  Check N
      IF (N .LE. 0) THEN
         CALL E1STI (1, N)
         CALL E1MES (5, 1, 'The number of variables, N = %(I1), must '//
     &               'be positive.')
         GO TO 9000
      END IF
C                                  Check FTOL
      IF (FTOL .LT. 0.0E0) THEN
         CALL E1STR (1, FTOL)
         CALL E1MES (5, 2, 'The function tolerance must be '//
     &               'nonnegative while FTOL = %(R1) is given.')
         GO TO 9000
      END IF
C                                  Check S
      IF (S .LE. 0.0E0) THEN
         CALL E1STR (1, S)
         CALL E1MES (6, 1, 'The length of the sides of the '//
     &               'initial simplex must be positive while S = '//
     &               '%(R1) is given.  The algorithm will use '//
     &               'a random number generator to generate the '//
     &               'starting simplex from the initial guess.')
      END IF
C
      CALL SCOPY (N, XGUESS, 1, WK(4*N+2), 1)
C
      CALL U3POL (FCN, N, S, FTOL, MAXFCN, WK(4*N+2), WK(3*N+1),
     &            WK(1), WK(N+1), WK(2*N+1))
C
      FVALUE = WK(3*N+1)
      CALL SCOPY (N, WK(4*N+2), 1, X, 1)
C
 9000 CALL E1POP ('U2POL ')
C
      RETURN
      END
C-----------------------------------------------------------------------
C  IMSL Name:  U3POL/DU3POL (Single/Double precision version)
C
C  Computer:   pcdsms/SINGLE
C
C  Revised:    November 1, 1984
C
C  Purpose:    Minimize a function in N variables using a direct search
C              polytope algorithm.
C
C  Usage:      CALL U3POL (FCN, N, S, FTOL, MAXFCN, X, F, C, XR, XE)
C
C  Arguments:
C     FCN    - See UMPOL.
C     N      - See UMPOL.
C     S      - See UMPOL.
C     FTOL   - See UMPOL.
C     MAXFCN - See UMPOL.
C     X      - Real N by N+1 work array containing the vertices of the
C              simplex.
C     F      - Real work vector of length N+1 containing the function
C              values at the simplices.
C     C      - Real work vector of length N.
C     XR     - Real work vector of length N.
C     XE     - Real work vector of length N.
C
C  Chapter:    MATH/LIBRARY Optimization
C
C  Copyright:  1984 by IMSL, Inc.  All Rights Reserved.
C
C  Warranty:   IMSL warrants only that IMSL testing has been applied
C              to this code.  No other warranty, expressed or implied,
C              is applicable.
C
C-----------------------------------------------------------------------
C
      SUBROUTINE U3POL (FCN, N, S, FTOL, MAXFCN, X, F, C, XR, XE)
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    N, MAXFCN
      REAL       S, FTOL, X(N,*), F(*), C(*), XR(*), XE(*)
      EXTERNAL   FCN
C                                  SPECIFICATIONS FOR PARAMETERS
      REAL       ALFA, BETA, GAMMA
      PARAMETER  (ALFA=1.0E0, BETA=2.0E0, GAMMA=0.5E0)
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    I, I2HIGH, IHIGH, ILOW, J, K, NFCN, NP1
      REAL       DIST, FAVG, FE, FR, FTOL2, FXC, RN, RNP1, STOPF,
     &           STOPX, SUM, VALUE
C                                  SPECIFICATIONS FOR INTRINSICS
C     INTRINSIC  ABS,AMIN1,FLOAT,SQRT
      INTRINSIC  ABS, AMIN1, FLOAT, SQRT
      REAL       ABS, AMIN1, FLOAT, SQRT
C                                  SPECIFICATIONS FOR SUBROUTINES
      EXTERNAL   E1MES, E1POP, E1PSH, E1STI, E1USR, RNSET, SCOPY
C                                  SPECIFICATIONS FOR FUNCTIONS
      EXTERNAL   ISMAX, ISMIN, SSUM, RNUNF
      INTEGER    ISMAX, ISMIN
      REAL       SSUM, RNUNF
C
      COMMON /O/ IOF(8),NPRNT,ITER,IPASS,IREST,ISUB
C
      CALL E1PSH ('U3POL ')
C                                  Initialization.
      CALL RNSET (123456)
      NPAR = N
      NP1 = N + 1
      RN = FLOAT(N)
      RNP1 = FLOAT(NP1)
      IDLW = 1
      NFCN = 0
      ITER = 0
      LDUM = 0
      IPASS = 0
      JPR = 0
      RMSE = 0.0
      FTOL2 = (FTOL**2)*RNP1
C                                  Restore saved run information.
C
      IF(IREST.NE.0) THEN
         IF(ISUB.NE.0) THEN
            DO 3 LL=1,N
               X(LL,(NP1+1)) = X(LL,1)
    3       CONTINUE
            CALL E1USR ('ON')
            CALL FCN (N, X(1,(NP1+1)), FSUB)
            CALL E1USR ('OFF')
         END IF
         OPEN(IOF(7),FILE='RESTART.DAT',STATUS='UNKNOWN',
     &               ACCESS='SEQUENTIAL',FORM='FORMATTED')
         READ(IOF(7),9201) NHEAD
         READ(IOF(7),*) LDUM,NFCN,IDLW,NPAR
         READ(IOF(7),9201) NHEAD
         READ(IOF(7),*) RMSE
         DO 5 K=1,NP1
            READ(IOF(7),9201) NHEAD
            READ(IOF(7),*) F(K)
            READ(IOF(7),9201) NHEAD
            READ(IOF(7),*) (X(L,K),L=1,N)
    5    CONTINUE
         ITER = LDUM
         IPASS = NFCN
         IF(ISUB.NE.0) THEN
            IDHI = ISMAX(NP1,F,1)
            IF(F(IDHI) .LE. FSUB) GO TO 45
            DO 7 MM=1,N
               X(MM,IDHI) = X(MM,(NP1+1))
    7       CONTINUE
            F(IDHI) = FSUB
            IDLW = ISMIN(NP1,F,1)
         END IF
         GO TO 45
      END IF
C                                  Generate the remaining N points
C                                  using a random number generator or
C                                  the initial length S if S is
C                                  positive.
      IF (S .GT. 0.0E0) THEN
         DO 10  I=2, NP1
            CALL SCOPY (N, X(1,1), 1, X(1,I), 1)
            X(I-1,I) = X(I-1,1) + S
   10    CONTINUE
      ELSE IF (S .LE. 0.0E0) THEN
         DO 30  J=2, NP1
            DO 20  I=1, N
               X(I,J) = X(I,1)*(1.0E0+RNUNF())
   20       CONTINUE
   30    CONTINUE
      END IF
C                                  Evaluate function at the current N+1
C                                  points.
      DO 40  I=1, NP1
         CALL E1USR ('ON')
         CALL FCN (N, X(1,I), F(I))
         CALL E1USR ('OFF')
   40 CONTINUE
      NFCN = NFCN + NP1
C                                  Write initial parameter guesses.
   45 CONTINUE

         WRITE(IOF(6),'(/)')
C         WRITE(  *   ,'(/)')
         WRITE(IOF(6),9100)
C         WRITE(  *   ,9100)
         WRITE(IOF(6),9101) (X(J,IDLW),J=1,N)
C         WRITE(  *   ,9101) (X(J,IDLW),J=1,N)
         WRITE(IOF(6),'(//)')
C         WRITE(  *   ,'(//)')
C                                  Main loop Find the indices of the
C                                  best (ILOW), worst (IHIGH), and
C                                  second worst point (I2HIGH).
   50 CONTINUE
      ITER = ITER + 1
      ILOW = ISMIN(NP1,F,1)
      IHIGH = ISMAX(NP1,F,1)
      I2HIGH = 1
      DO 60  I=2, NP1
         IF (I .EQ. IHIGH) GO TO 60
         IF (F(I) .GT. F(I2HIGH)) I2HIGH = I
   60 CONTINUE
C                                  Check for relative error in the
C                                  function value. Stop if F(IHIGH) -
C                                  F(ILOW) < FTOL * ( 1 +
C                                  ABS(F(ILOW)))
      STOPF = F(IHIGH) - F(ILOW)
      IF (STOPF .LT. FTOL*(1.0E0+ABS(F(ILOW)))) GO TO 150
C
C                                  Compute the centroid of the best N
C                                  vertices. Perform the reflection
C                                  step and compute F at the trial
C                                  point.
      DO 80  I=1, N
         C(I) = 0.0E0
         DO 70  J=1, NP1
            IF (J .NE. IHIGH) C(I) = C(I) + X(I,J)
   70    CONTINUE
         C(I) = C(I)/RN
         XR(I) = C(I) + ALFA*(C(I)-X(I,IHIGH))
   80 CONTINUE
      CALL E1USR ('ON')
      CALL FCN (N, XR, FR)
      CALL E1USR ('OFF')
      NFCN = NFCN + 1
C                                  XR is not either a new best point of
C                                  a new worst point. XR replaces
C                                  worst point and the next iteration
C                                  is begun.
C
      IF (F(ILOW).LE.FR .AND. FR.LE.F(I2HIGH)) THEN
         F(IHIGH) = FR
         CALL SCOPY (N, XR, 1, X(1,IHIGH), 1)
C                                  XR is the new best point. Perform
C                                  the expansion step by computing
C                                  the expanded point XE (stored in
C                                  XE) and its function value FE. If
C                                  FE < FR the expansion is
C                                  successful, XE replaces worst
C                                  point. else, the expansion has
C                                  failed and worst point is replaced
C                                  by XR.
      ELSE IF (FR .LT. F(ILOW)) THEN
         DO 90  I=1, N
            XE(I) = C(I) + BETA*(XR(I)-C(I))
   90    CONTINUE
         CALL E1USR ('ON')
         CALL FCN (N, XE, FE)
         CALL E1USR ('OFF')
         NFCN = NFCN + 1
         IF (FE .LT. FR) THEN
            CALL SCOPY (N, XE, 1, X(1,IHIGH), 1)
            F(IHIGH) = FE
         ELSE
            CALL SCOPY (N, XR, 1, X(1,IHIGH), 1)
            F(IHIGH) = FR
         END IF
C                                  The current polytope is assumed to
C                                  be too large and should be
C                                  contracted.
      ELSE IF (FR .GT. F(I2HIGH)) THEN
         IF (FR .GE. F(IHIGH)) THEN
            DO 100  I=1, N
               XE(I) = C(I) + GAMMA*(X(I,IHIGH)-C(I))
  100       CONTINUE
         ELSE IF (FR .LT. F(IHIGH)) THEN
            DO 110  I=1, N
               XE(I) = C(I) + GAMMA*(XR(I)-C(I))
  110       CONTINUE
         END IF
         CALL E1USR ('ON')
         CALL FCN (N, XE, FXC)
         CALL E1USR ('OFF')
         NFCN = NFCN + 1
         VALUE = AMIN1(FR,F(IHIGH))
         IF (FXC .LT. VALUE) THEN
            F(IHIGH) = FXC
            CALL SCOPY (N, XE, 1, X(1,IHIGH), 1)
         ELSE
C                                  The contraction step is
C                                  unsuccessful, the polytope is
C                                  shrunk by moving the vertices
C                                  half-way toward current best
C                                  point.
            DO 130  J=1, NP1
               IF (J .NE. ILOW) THEN
                  DO 120  I=1, N
                     X(I,J) = 0.5E0*(X(I,J)+X(I,ILOW))
  120             CONTINUE
                  CALL E1USR ('ON')
                  CALL FCN (N, X(1,J), F(J))
                  CALL E1USR ('OFF')
               END IF
  130       CONTINUE
            NFCN = NFCN + N
         END IF
      END IF
C                                  Test for convergence. Convergence is
C                                  assumed when the variance of the
C                                  Fi's is less than a prescribed
C                                  tolerance.
      IF (NFCN .GT. MAXFCN) THEN
         CALL E1STI (1, MAXFCN)
         CALL E1MES (4, 1, 'Maximum number of function evaluations '//
     &               'exceeded MAXFCN = %(I1).')
         MAXFCN = NFCN
         GO TO 9000
      END IF
         FAVG = SSUM(NP1,F,1)/RNP1
         STOPX = 0.0E0
         DO 140  K=1, NP1
            STOPX = STOPX + (F(K)-FAVG)**2
  140    CONTINUE
C                                  Write iteration results
C                                  to screen.
         IDLW = ISMIN(NP1,F,1)
         RMSE = SQRT(STOPX/RNP1)
            WRITE(  *   ,9102) ITER,NFCN
            WRITE(  *   ,9103)
            WRITE(  *   ,9101) (X(J,IDLW),J=1,N)
            WRITE(  *   ,9104) F(IDLW),RMSE
C
C                                  Increment print counter (JPR) and
C                                  test whether to print iteration
C                                  results.
         JPR = JPR + 1
         IF((ITER-LDUM).EQ.1) GO TO 144
         IF(JPR-NPRNT) 148,142,142
  142    JPR = 0
  144    CONTINUE
            WRITE(IOF(6),9102) ITER,NFCN
            WRITE(IOF(6),9103)
            WRITE(IOF(6),9101) (X(J,IDLW),J=1,N)
            WRITE(IOF(6),9104) F(IDLW),RMSE

C                                  Update restart log file.
C
         OPEN(IOF(8),FILE='RESTART.LOG',STATUS='UNKNOWN',
     &               ACCESS='SEQUENTIAL',FORM='FORMATTED')
         ENDFILE IOF(8)
         REWIND IOF(8)
            WRITE(IOF(8),9203)
            WRITE(IOF(8),9202) ITER,NFCN,IDLW,NPAR
            WRITE(IOF(8),9204)
C            WRITE(IOF(8),9101) RMSE
            WRITE(IOF(8),9106) RMSE
         DO 146 K=1,NP1
            WRITE(IOF(8),9205) K
C            WRITE(IOF(8),9101) F(K)
            WRITE(IOF(8),9106) F(K)
            WRITE(IOF(8),9206) K,NPAR
C            WRITE(IOF(8),9101) (X(L,K),L=1,N)
            WRITE(IOF(8),9106) (X(L,K),L=1,N)
  146    CONTINUE
         CLOSE(IOF(8))
  148    CONTINUE
         IF (STOPX .GT. FTOL2) GO TO 50
  150 CONTINUE
C                                  Compute centroid of the final N+1
C                                  points and the average distance
C                                  from the vertices to the centroid.
C                                  On return overwrite S with the
C                                  average distance and MAXFCN with
C                                  number of function evaluations.
      MAXFCN = NFCN
      DO 160  J=1, N
         C(J) = SSUM(NP1,X(J,1),N)/RNP1
  160 CONTINUE
C
      DIST = 0.0E0
      DO 180  I=1, NP1
         SUM = 0.0E0
         DO 170  J=1, N
            SUM = SUM + (X(J,I)-C(J))**2
  170    CONTINUE
         DIST = DIST + SQRT(SUM)
  180 CONTINUE
      S = DIST/RNP1
C                                  Return best point and best function
C                                  value.
      ILOW = ISMIN(NP1,F,1)
      CALL SCOPY (N, X(1,ILOW), 1, X(1,1), 1)
      F(1) = F(ILOW)
C                                  Write final solution.
      WRITE(IOF(6),9105) ITER
C      WRITE(  *   ,9105) ITER
      WRITE(IOF(6),9101) (X(J,ILOW),J=1,N)
C      WRITE(  *   ,9101) (X(J,ILOW),J=1,N)
      WRITE(IOF(6),9104) F(ILOW),RMSE
C      WRITE(  *   ,9104) F(ILOW),RMSE
C
 9000 CALL E1POP ('U3POL ')
      RETURN
C                                  Format statements:
C
9100  FORMAT(2X,'INITIAL PARAMETER GUESSES:',/)
C***9101  FORMAT(' A = ',1PE15.8, 3X,' B = ',1PE15.8,3X,' C = ',1PE15.8,/)  
C**9101  FORMAT(' A = ',1PE15.8, 3X,' B = ',1PE15.8,3X,/)  
9101  FORMAT(4(1PE15.5))
9102  FORMAT(/,2X,'ITERATION',I5,',',
     &         5X,'Number of Function Evaluations:',I5)
9103  FORMAT(2X,'Current Parameter Values:')
9104  FORMAT(2X,'F-OBJ = ',1PE15.8,',',5X,'RMSE = ',E15.8,/)
9105  FORMAT(//,2X,'FINAL PARAMETER VALUES:',10X,'TOTAL',
     &          1X,'ITERATIONS =',I5)
9106  FORMAT(4(1PE15.5))
9201  FORMAT(40A2)
9202  FORMAT(15I5)
9203  FORMAT(' ITER NCFN  ILOW NPAR')
9204  FORMAT(2X,'RMSE =')
9205  FORMAT(2X,'Function Value, F(',I2,') =')
9206  FORMAT(2X,'X(I,',I2,'), I = 1,',I2,'...')
      END
C-----------------------------------------------------------------------
C  IMSL Name:  UMPOL/DUMPOL (Single/Double precision version)
C
C  Computer:   pcdsms/SINGLE
C
C  Revised:    October 10, 1985
C
C  Purpose:    Minimize a function of N variables using a direct search
C              polytope algorithm.
C
C  Usage:      CALL UMPOL (FCN, N, XGUESS, S, FTOL, MAXFCN, X, FVALUE)
C
C  Arguments:
C     FCN    - User-supplied SUBROUTINE to evaluate the function to be
C              minimized.  The usage is
C              CALL FCN (N, X, F), where
C              N      - Length of X.  (Input)
C              X      - Vector of length N at which the function is
C                       evaluated.  (Input)
C                       X should not be changed by FCN.
C              F      - The computed function value at the point X.
C                       (Output)
C              FCN must be declared EXTERNAL in the calling program.
C     N      - Dimension of the problem.  (Input)
C     XGUESS - Real vector of length N which contains an initial guess
C              to the minimum.  (Input)
C     S      - On input, real scalar containing the length of each
C              side of the intitial simplex.  (Input/Output)
C              If no reasonable information about S is known, S could
C              be set to a number less than or equal to zero and UMPOL
C              will generate the starting simplex from the initial
C              guess with a random number generator.  On output, the
C              average distance from the vertices to the centroid
C              which is taken to be the solution, see Remark 4.
C     FTOL   - First convergence criterion.  (Input)
C              The algorithm stops when a relative error in the function
C              values is less than FTOL, i.e. when (F(worst) - F(best))
C              < FTOL*(1+ABS(F(best))) where F(worst) and F(best) are
C              the function values of the current worst and best point
C              respectively.
C              Second convergence criterion. The algorithm stops when
C              the standard deviation of the function values at the N+1
C              current points is less than FTOL.
C              If the subroutine terminates prematurely, try again with
C              a smaller value for FTOL.
C     MAXFCN - On input, maximum allowed number of function evaluations.
C              (Input/Output)
C              On output, actual number of function evaluations needed.
C     X      - Real vector of length N containing the best estimate
C              of the minimum found.  (Output)
C     FVALUE - Function value at the computed solution.  (Output)
C
C  Remarks:
C  1. Automatic workspace usage is
C              UMPOL     N**2 + 5*N + 1  units, or
C              DUMPOL    2*N**2 + 10*N + 2 units.
C     Workspace may be explicitly provided, if desired, by use of
C     U2POL/DU2POL.  The reference is
C              CALL U2POL (FCN, N, XGUESS, S, FTOL, MAXFCN, X, FVALUE,
C                          WK)
C     The additional argument is
C     WK     - Real work vector of length LWK.
C
C  2. Informational error
C     Type Code
C       4   1  Maximum number of function evaluations exceeded.
C
C  3. Since UMPOL uses only function value information at each step
C     to determine a new approximate minimum, it could be quite ineffi-
C     cient on smooth problems compared to other methods such as those
C     implemented in routine UMINF which takes into account derivative
C     information at each iteration. Hence subroutine UMPOL should only
C     be used as a last resort. Briefly, a set of N + 1 points in an
C     N-dimensional space is called a simplex. The minimization process
C     iterates by replacing the point with the largest function value by
C     a new point with a smaller function value. The iteration continues
C     until all the points cluster sufficiently close to a minimum.
C
C  4. The value returned in S is useful for assessing the flatness of
C     the function near the computed minimum. The larger its value for a
C     given value of FTOL, the flatter the function tends to be in the
C     neighborhood of the returned point.
C
C  Keywords:   Discontinuous; Direct search; Pattern search
C
C  GAMS:       G1b2
C
C  Chapter:    MATH/LIBRARY Optimization
C
C  Copyright:  1985 by IMSL, Inc.  All Rights Reserved.
C
C  Warranty:   IMSL warrants only that IMSL testing has been applied
C              to this code.  No other warranty, expressed or implied,
C              is applicable.
C
C-----------------------------------------------------------------------
C
      SUBROUTINE UMPOL (FCN, N, XGUESS, S, FTOL, MAXFCN, X, FVALUE)
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    N, MAXFCN
      REAL       S, FTOL, FVALUE, XGUESS(*), X(*)
      EXTERNAL   FCN
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    INDR, LWK
C                                  SPECIFICATIONS FOR SPECIAL CASES
C                                  SPECIFICATIONS FOR COMMON /WORKSP/
      REAL       RWKSP(5000)
      REAL       RDWKSP(5000)
      DOUBLE PRECISION DWKSP(2500)
      COMPLEX    CWKSP(2500)
      COMPLEX    CZWKSP(2500)
      COMPLEX    *16 ZWKSP(1250)
      INTEGER    IWKSP(5000)
      LOGICAL    LWKSP(5000)
      EQUIVALENCE (DWKSP(1), RWKSP(1))
      EQUIVALENCE (CWKSP(1), RWKSP(1)), (ZWKSP(1), RWKSP(1))
      EQUIVALENCE (IWKSP(1), RWKSP(1)), (LWKSP(1), RWKSP(1))
      EQUIVALENCE (RDWKSP(1), RWKSP(1)), (CZWKSP(1), RWKSP(1))
      COMMON     /WORKSP/ RWKSP
C                                  SPECIFICATIONS FOR SUBROUTINES
      EXTERNAL   E1MES, E1POP, E1PSH, E1STI, U2POL
C                                  SPECIFICATIONS FOR FUNCTIONS
      EXTERNAL   I1KGT, N1RTY
      INTEGER    I1KGT, N1RTY
C
      CALL E1PSH ('UMPOL ')
C                                  Check N
      IF (N .LE. 0) THEN
         CALL E1STI (1, N)
         CALL E1MES (5, 1, 'The number of variables, N = %(I1), must '//
     &               'be positive.')
         GO TO 9000
      END IF
C                                  Allocate workspace
      LWK = N**2 + 5*N + 1
      INDR = I1KGT(LWK,3)
      IF (N1RTY(0) .NE. 0) THEN
         CALL E1MES (5, -1, ' ')
C  MOVED E1STI (1, N) UP ON SAME LINE AS CALL         
         CALL E1STI (1, N)
         CALL E1MES (5, 2, 'The workspace requirement is based on '//
     &               'N = %(I1). ')
         GO TO 9000
      END IF
C
      CALL U2POL (FCN, N, XGUESS, S, FTOL, MAXFCN, X, FVALUE,
     &            RDWKSP(INDR))
C
 9000 CALL E1POP ('UMPOL ')
C
      RETURN
      END
