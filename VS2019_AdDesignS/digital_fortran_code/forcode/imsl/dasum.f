C-----------------------------------------------------------------------
C  IMSL Name:  DASUM (Double precision version)
C
C  Computer:   pcdsms/DOUBLE
C
C  Revised:    August 9, 1986
C
C  Purpose:    Sum the absolute values of the components of a
C              double precision vector.
C
C  Usage:      DASUM(N, DX, INCX)
C
C  Arguments:
C     N      - Length of vectors X.  (Input)
C     DX     - Double precision vector of length N*INCX.  (Input)
C     INCX   - Displacement between elements of DX.  (Input)
C              X(I) is defined to be DX(1+(I-1)*INCX).  INCX must be
C              greater than 0.
C     DASUM  - Double precision sum from I=1 to N of DABS(X(I)).
C              (Output)
C              X(I) refers to a specific element of DX.
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
      DOUBLE PRECISION FUNCTION DASUM (N, DX, INCX)
C                                  SPECIFICATIONS FOR ARGUMENTS
      INTEGER    N, INCX
      DOUBLE PRECISION DX(*)
C                                  SPECIFICATIONS FOR LOCAL VARIABLES
      INTEGER    I, M, MP1, NINCX
C                                  SPECIFICATIONS FOR SPECIAL CASES
C     INTRINSIC  MOD
      INTRINSIC  MOD
      INTEGER    MOD
C                                  SPECIFICATIONS FOR INTRINSICS
C     INTRINSIC  DABS
      INTRINSIC  DABS
      DOUBLE PRECISION DABS
C
      DASUM = 0.0D0
      IF (N .GT. 0) THEN
         IF (INCX .NE. 1) THEN
C                                  CODE FOR INCREMENT NOT EQUAL TO 1
            NINCX = N*INCX
            DO 10  I=1, NINCX, INCX
               DASUM = DASUM + DABS(DX(I))
   10       CONTINUE
         ELSE
C                                  CODE FOR INCREMENT EQUAL TO 1
            M = MOD(N,6)
C                                  CLEAN-UP LOOP
            DO 30  I=1, M
               DASUM = DASUM + DABS(DX(I))
   30       CONTINUE
            MP1 = M + 1
            DO 40  I=MP1, N, 6
               DASUM = DASUM + DABS(DX(I)) + DABS(DX(I+1)) +
     &                 DABS(DX(I+2)) + DABS(DX(I+3)) + DABS(DX(I+4)) +
     &                 DABS(DX(I+5))
   40       CONTINUE
         END IF
      END IF
      RETURN
      END
